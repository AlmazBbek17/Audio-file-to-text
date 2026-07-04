import os
import sqlite3
import tempfile
from datetime import datetime, timezone

import requests
import yt_dlp
from flask import Flask, request, jsonify
from flask_cors import CORS

app = Flask(__name__)
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
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
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
            source_type TEXT NOT NULL,
            title TEXT,
            counted INTEGER NOT NULL DEFAULT 0,
            created_at TEXT
        )
    """)
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


def save_job(transcript_id, device_id, source_type, title):
    conn = get_db()
    conn.execute(
        "INSERT INTO jobs (transcript_id, device_id, source_type, title, created_at) VALUES (?, ?, ?, ?, ?)",
        (transcript_id, device_id, source_type, title, datetime.now(timezone.utc).isoformat()),
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
        json={"audio_url": audio_url},
    )
    transcript_res.raise_for_status()
    transcript_id = transcript_res.json()["id"]

    save_job(transcript_id, device_id, "file", uploaded.filename)
    return jsonify({"job_id": transcript_id})


@app.route("/api/transcribe/youtube", methods=["POST"])
def transcribe_youtube():
    data = request.get_json(force=True)
    device_id = data.get("device_id", "")
    url = data.get("url", "")

    if get_seconds_used(device_id) >= FREE_LIMIT_SECONDS:
        return jsonify({"error": "limit_reached"}), 402
    if not url:
        return jsonify({"error": "no_url"}), 400

    # Извлекаем ПРЯМУЮ ссылку на аудиопоток, ничего не скачивая на сервер.
    ydl_opts = {
        "format": "bestaudio/best",
        "quiet": True,
        "noplaylist": True,
        "skip_download": True,
    }
    try:
        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            info = ydl.extract_info(url, download=False)
    except Exception as e:
        return jsonify({"error": "youtube_extract_failed", "detail": str(e)}), 400

    direct_audio_url = info.get("url")
    title = info.get("title", url)

    if not direct_audio_url:
        return jsonify({"error": "no_audio_stream"}), 400

    transcript_res = requests.post(
        f"{ASSEMBLYAI_BASE}/transcript",
        headers=assemblyai_headers(),
        json={"audio_url": direct_audio_url},
    )
    transcript_res.raise_for_status()
    transcript_id = transcript_res.json()["id"]

    save_job(transcript_id, device_id, "youtube", title)
    return jsonify({"job_id": transcript_id, "title": title})


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
        if not job["counted"]:
            add_seconds_used(device_id, duration)
            mark_job_counted(transcript_id)
        return jsonify({
            "status": "completed",
            "text": data.get("text", ""),
            "audio_duration": duration,
            "source_type": job["source_type"],
            "title": job["title"],
        })

    if aai_status == "error":
        return jsonify({"status": "error", "detail": data.get("error")})

    return jsonify({"status": aai_status})


@app.route("/", methods=["GET"])
def health():
    return jsonify({"ok": True})


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)
