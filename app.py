import io
import json
import os
import sqlite3
from datetime import datetime, timezone

import requests
from docx import Document
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS

# WeasyPrint тянет системные библиотеки (Pango/Cairo). Если их нет на хосте — импорт падает
# с OSError. Раньше это падало на старте всего приложения и валило ВЕСЬ бэкенд, включая
# транскрипцию файлов. Теперь импортируем лениво, только когда реально нужен PDF-экспорт —
# остальной функционал продолжает работать, даже если PDF временно недоступен.
_weasyprint_error = None
try:
    from weasyprint import HTML
except Exception as e:  # noqa: BLE001
    HTML = None
    _weasyprint_error = str(e)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 500 * 1024 * 1024  # 500 МБ, как заявлено в UI
# MVP: разрешаем любые origin (расширение работает без бэкенд-аутентификации).
# После публикации можно сузить до конкретного chrome-extension://<ID>.
CORS(app, resources={r"/api/*": {"origins": "*"}})

ASSEMBLYAI_API_KEY = os.environ.get("ASSEMBLYAI_API_KEY")
ASSEMBLYAI_BASE = "https://api.assemblyai.com/v2"

FREE_LIMIT_SECONDS = 30 * 60  # 30 минут — держим то же значение, что и в extension/config.js

# На Railway подключи Volume и примонтируй его в /data, чтобы SQLite не терялась при redeploy.
DB_PATH = os.environ.get("DB_PATH", "/data/app.db")


def get_db():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    conn = sqlite3.connect(DB_PATH, timeout=10)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")  # снижает блокировки при параллельных воркерах gunicorn
    return conn


def init_db():
    conn = get_db()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS usage (
            device_id TEXT PRIMARY KEY,
            seconds_used REAL NOT NULL DEFAULT 0,
            updated_at TEXT
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS jobs (
            transcript_id TEXT PRIMARY KEY,
            device_id TEXT NOT NULL,
            title TEXT,
            counted INTEGER NOT NULL DEFAULT 0,
            created_at TEXT,
            text TEXT,
            utterances TEXT,
            duration REAL
        )
    """)

    # Миграция со времён, когда была YouTube-ссылка: колонка source_type была NOT NULL,
    # без значения по умолчанию — INSERT в неё сейчас падает. CREATE TABLE IF NOT EXISTS
    # не трогает уже существующую таблицу, поэтому чиним руками, если колонка ещё жива.
    existing_cols = [row[1] for row in conn.execute("PRAGMA table_info(jobs)").fetchall()]
    if "source_type" in existing_cols:
        conn.execute("ALTER TABLE jobs RENAME TO jobs_old")
        conn.execute("""
            CREATE TABLE jobs (
                transcript_id TEXT PRIMARY KEY,
                device_id TEXT NOT NULL,
                title TEXT,
                counted INTEGER NOT NULL DEFAULT 0,
                created_at TEXT,
                text TEXT,
                utterances TEXT,
                duration REAL
            )
        """)
        old_cols = [row[1] for row in conn.execute("PRAGMA table_info(jobs_old)").fetchall()]
        target_cols = ["transcript_id", "device_id", "title", "counted", "created_at", "text", "utterances", "duration"]
        # если в старой таблице какой-то колонки не было — подставляем NULL, а не падаем
        select_list = ", ".join(c if c in old_cols else "NULL" for c in target_cols)
        conn.execute(f"INSERT INTO jobs ({', '.join(target_cols)}) SELECT {select_list} FROM jobs_old")
        conn.execute("DROP TABLE jobs_old")

    conn.commit()
    conn.close()


init_db()


def get_seconds_used(device_id):
    conn = get_db()
    row = conn.execute("SELECT seconds_used FROM usage WHERE device_id=?", (device_id,)).fetchone()
    conn.close()
    return row["seconds_used"] if row else 0


def add_seconds_used(device_id, seconds):
    conn = get_db()
    conn.execute("""
        INSERT INTO usage (device_id, seconds_used, updated_at)
        VALUES (?, ?, ?)
        ON CONFLICT(device_id) DO UPDATE SET
            seconds_used = seconds_used + excluded.seconds_used,
            updated_at = excluded.updated_at
    """, (device_id, seconds, datetime.now(timezone.utc).isoformat()))
    conn.commit()
    conn.close()


def save_job(transcript_id, device_id, title):
    conn = get_db()
    conn.execute(
        "INSERT INTO jobs (transcript_id, device_id, title, created_at) VALUES (?, ?, ?, ?)",
        (transcript_id, device_id, title, datetime.now(timezone.utc).isoformat()),
    )
    conn.commit()
    conn.close()


def get_job(transcript_id):
    conn = get_db()
    row = conn.execute("SELECT * FROM jobs WHERE transcript_id=?", (transcript_id,)).fetchone()
    conn.close()
    return row


def mark_job_counted(transcript_id):
    conn = get_db()
    conn.execute("UPDATE jobs SET counted=1 WHERE transcript_id=?", (transcript_id,))
    conn.commit()
    conn.close()


def save_job_result(transcript_id, text, utterances, duration):
    conn = get_db()
    conn.execute(
        "UPDATE jobs SET text=?, utterances=?, duration=? WHERE transcript_id=?",
        (text, json.dumps(utterances, ensure_ascii=False), duration, transcript_id),
    )
    conn.commit()
    conn.close()


def build_diarized_text(data):
    """Собираем текст с метками спикеров, если AssemblyAI вернул диаризацию."""
    utterances = data.get("utterances") or []
    if not utterances:
        return data.get("text", ""), []
    lines = [f"Speaker {u['speaker']}: {u['text']}" for u in utterances]
    return "\n\n".join(lines), utterances


def ms_to_srt_timestamp(ms):
    total_seconds, millis = divmod(int(ms), 1000)
    hours, remainder = divmod(total_seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    return f"{hours:02d}:{minutes:02d}:{seconds:02d},{millis:03d}"


def build_srt(utterances, fallback_text, duration_seconds):
    if not utterances:
        # Нет данных о таймингах реплик — отдаём один блок на весь ролик.
        end_ts = ms_to_srt_timestamp((duration_seconds or 0) * 1000)
        return f"1\n00:00:00,000 --> {end_ts}\n{fallback_text}\n"
    blocks = []
    for i, u in enumerate(utterances, start=1):
        start_ts = ms_to_srt_timestamp(u["start"])
        end_ts = ms_to_srt_timestamp(u["end"])
        blocks.append(f"{i}\n{start_ts} --> {end_ts}\nSpeaker {u['speaker']}: {u['text']}\n")
    return "\n".join(blocks)


def build_docx(title, text, utterances):
    doc = Document()
    doc.add_heading(title, level=1)
    if utterances:
        for u in utterances:
            p = doc.add_paragraph()
            run = p.add_run(f"Speaker {u['speaker']}: ")
            run.bold = True
            p.add_run(u["text"])
    else:
        for para in text.split("\n"):
            doc.add_paragraph(para)
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


def build_pdf(title, text, utterances):
    if HTML is None:
        raise RuntimeError(f"PDF export unavailable: {_weasyprint_error}")
    if utterances:
        body = "".join(
            f"<p><b>Speaker {u['speaker']}:</b> {u['text']}</p>" for u in utterances
        )
    else:
        body = "".join(f"<p>{para}</p>" for para in text.split("\n") if para.strip())
    html = f"""
    <html><head><meta charset="utf-8"><style>
        body {{ font-family: sans-serif; font-size: 13px; line-height: 1.5; }}
        h1 {{ font-size: 18px; }}
    </style></head>
    <body><h1>{title}</h1>{body}</body></html>
    """
    buf = io.BytesIO()
    HTML(string=html).write_pdf(buf)
    buf.seek(0)
    return buf


def assemblyai_headers():
    return {"authorization": ASSEMBLYAI_API_KEY}


@app.route("/api/usage", methods=["GET"])
def usage():
    device_id = request.args.get("device_id", "")
    return jsonify({"seconds_used": get_seconds_used(device_id), "limit_seconds": FREE_LIMIT_SECONDS})


@app.route("/api/transcribe/file", methods=["POST"])
def transcribe_file():
    device_id = request.form.get("device_id", "")
    if get_seconds_used(device_id) >= FREE_LIMIT_SECONDS:
        return jsonify({"error": "limit_reached"}), 402

    uploaded = request.files.get("file")
    if not uploaded:
        return jsonify({"error": "no_file"}), 400

    # Заливаем файл напрямую в AssemblyAI (потоково, без лишнего диска)
    upload_res = requests.post(
        f"{ASSEMBLYAI_BASE}/upload",
        headers=assemblyai_headers(),
        data=uploaded.stream.read(),
    )
    upload_res.raise_for_status()
    audio_url = upload_res.json()["upload_url"]

    transcript_res = requests.post(
        f"{ASSEMBLYAI_BASE}/transcript",
        headers=assemblyai_headers(),
        json={"audio_url": audio_url, "speaker_labels": True},
    )
    transcript_res.raise_for_status()
    transcript_id = transcript_res.json()["id"]

    save_job(transcript_id, device_id, uploaded.filename)
    return jsonify({"job_id": transcript_id})


@app.route("/api/status/<transcript_id>", methods=["GET"])
def status(transcript_id):
    device_id = request.args.get("device_id", "")
    job = get_job(transcript_id)
    if not job:
        return jsonify({"error": "job_not_found"}), 404

    res = requests.get(f"{ASSEMBLYAI_BASE}/transcript/{transcript_id}", headers=assemblyai_headers())
    res.raise_for_status()
    data = res.json()

    aai_status = data["status"]  # queued | processing | completed | error

    if aai_status == "completed":
        duration = data.get("audio_duration") or 0
        diarized_text, utterances = build_diarized_text(data)
        save_job_result(transcript_id, diarized_text, utterances, duration)
        if not job["counted"]:
            add_seconds_used(device_id, duration)
            mark_job_counted(transcript_id)
        return jsonify({
            "status": "completed",
            "text": diarized_text,
            "audio_duration": duration,
            "title": job["title"],
        })

    if aai_status == "error":
        return jsonify({"status": "error", "detail": data.get("error")})

    return jsonify({"status": aai_status})


@app.route("/api/export/<transcript_id>", methods=["GET"])
def export(transcript_id):
    fmt = request.args.get("format", "txt")
    job = get_job(transcript_id)
    if not job:
        return jsonify({"error": "job_not_found"}), 404

    text = job["text"] or ""
    title = (job["title"] or "transcript").rsplit(".", 1)[0]
    utterances = json.loads(job["utterances"]) if job["utterances"] else []
    safe_name = "".join(c for c in title if c.isalnum() or c in " _-").strip() or "transcript"

    if fmt == "txt":
        buf = io.BytesIO(text.encode("utf-8"))
        return send_file(buf, mimetype="text/plain", as_attachment=True, download_name=f"{safe_name}.txt")

    if fmt == "srt":
        srt_content = build_srt(utterances, text, job["duration"] or 0)
        buf = io.BytesIO(srt_content.encode("utf-8"))
        return send_file(buf, mimetype="application/x-subrip", as_attachment=True, download_name=f"{safe_name}.srt")

    if fmt == "docx":
        buf = build_docx(title, text, utterances)
        return send_file(
            buf,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            as_attachment=True,
            download_name=f"{safe_name}.docx",
        )

    if fmt == "pdf":
        try:
            buf = build_pdf(title, text, utterances)
        except RuntimeError as e:
            return jsonify({"error": "pdf_unavailable", "detail": str(e)}), 503
        return send_file(buf, mimetype="application/pdf", as_attachment=True, download_name=f"{safe_name}.pdf")

    return jsonify({"error": "unknown_format"}), 400


@app.route("/", methods=["GET"])
def health():
    return jsonify({"ok": True, "pdf_available": HTML is not None})


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)
