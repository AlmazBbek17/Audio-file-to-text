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
SPEECH_MODEL = "universal"  # самая дешёвая модель AssemblyAI ($0.15/час) — фиксируем явно,
                            # чтобы себестоимость была предсказуемой, а не "что дадут по умолчанию"

FREE_LIMIT_SECONDS = 30 * 60  # 30 минут — держим то же значение, что и в extension/config.js

DODO_API_KEY = os.environ.get("DODO_PAYMENTS_API_KEY")
DODO_ENVIRONMENT = os.environ.get("DODO_ENVIRONMENT", "live_mode")  # "test_mode" на время тестов
DODO_WEBHOOK_KEY = os.environ.get("DODO_WEBHOOK_KEY")

# Три продукта-подписки, которые нужно создать в дашборде Dodo — id проставь в Railway Variables.
# minutes — это то, что должно быть настроено как "credits issued" на самом продукте в Dodo
# (Credit entitlement "minutes", precision 0) — тут только для отображения в /api/usage и UI.
PLANS = {
    "weekly": {"product_id": os.environ.get("DODO_PRODUCT_WEEKLY"), "label": "Weekly", "minutes": 600},
    "monthly": {"product_id": os.environ.get("DODO_PRODUCT_MONTHLY"), "label": "Monthly", "minutes": 1800},
    # Годовой биллится одним платежом в год — credits логично выдавать сразу на весь год
    # (1800/мес × 12), а не рассчитывать на помесячный reissue при годовом billing cycle.
    "yearly": {"product_id": os.environ.get("DODO_PRODUCT_YEARLY"), "label": "Yearly", "minutes": 1800 * 12},
}
# Событие, которое шлём в Meter после каждой завершённой транскрипции — должно ТОЧНО совпадать
# с Event Name, указанным при создании Meter в дашборде Dodo.
USAGE_EVENT_NAME = "transcription_completed"

_dodo_client = None
def dodo_client():
    global _dodo_client
    if _dodo_client is None:
        from dodopayments import DodoPayments
        _dodo_client = DodoPayments(bearer_token=DODO_API_KEY, environment=DODO_ENVIRONMENT)
    return _dodo_client

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
        CREATE TABLE IF NOT EXISTS customers (
            device_id TEXT PRIMARY KEY,
            dodo_customer_id TEXT NOT NULL,
            updated_at TEXT
        )
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS subscriptions (
            device_id TEXT PRIMARY KEY,
            subscription_id TEXT,
            plan TEXT,
            status TEXT,
            remaining_balance REAL,
            next_billing_date TEXT,
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

    # Более старая миграция: колонка paid_seconds_balance из версии с разовыми платежами ($1/час)
    # больше не нужна — теперь баланс живёт в Dodo и приходит через вебхук подписки.
    usage_cols = [row[1] for row in conn.execute("PRAGMA table_info(usage)").fetchall()]
    if "paid_seconds_balance" in usage_cols:
        conn.execute("ALTER TABLE usage RENAME TO usage_old")
        conn.execute("""
            CREATE TABLE usage (
                device_id TEXT PRIMARY KEY,
                seconds_used REAL NOT NULL DEFAULT 0,
                updated_at TEXT
            )
        """)
        conn.execute("INSERT INTO usage (device_id, seconds_used, updated_at) SELECT device_id, seconds_used, updated_at FROM usage_old")
        conn.execute("DROP TABLE usage_old")

    conn.commit()
    conn.close()


init_db()


def get_usage_row(device_id):
    conn = get_db()
    row = conn.execute("SELECT seconds_used FROM usage WHERE device_id=?", (device_id,)).fetchone()
    conn.close()
    return {"seconds_used": row["seconds_used"] if row else 0}


def get_subscription(device_id):
    conn = get_db()
    row = conn.execute("SELECT * FROM subscriptions WHERE device_id=?", (device_id,)).fetchone()
    conn.close()
    return row


def has_credit(device_id):
    """Бесплатные 30 минут ИЛИ активная подписка (превышение лимита Dodo биллит сама —
    поэтому наличие активной подписки достаточно, отдельно считать остаток не нужно)."""
    if get_usage_row(device_id)["seconds_used"] < FREE_LIMIT_SECONDS:
        return True
    sub = get_subscription(device_id)
    return bool(sub and sub["status"] == "active")


def add_free_seconds_used(device_id, seconds):
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


def upsert_customer(device_id, dodo_customer_id):
    conn = get_db()
    conn.execute("""
        INSERT INTO customers (device_id, dodo_customer_id, updated_at)
        VALUES (?, ?, ?)
        ON CONFLICT(device_id) DO UPDATE SET
            dodo_customer_id = excluded.dodo_customer_id,
            updated_at = excluded.updated_at
    """, (device_id, dodo_customer_id, datetime.now(timezone.utc).isoformat()))
    conn.commit()
    conn.close()


def get_dodo_customer_id(device_id):
    conn = get_db()
    row = conn.execute("SELECT dodo_customer_id FROM customers WHERE device_id=?", (device_id,)).fetchone()
    conn.close()
    return row["dodo_customer_id"] if row else None


def upsert_subscription(device_id, subscription_id, plan, status, remaining_balance, next_billing_date):
    conn = get_db()
    conn.execute("""
        INSERT INTO subscriptions (device_id, subscription_id, plan, status, remaining_balance, next_billing_date, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(device_id) DO UPDATE SET
            subscription_id = excluded.subscription_id,
            plan = excluded.plan,
            status = excluded.status,
            remaining_balance = excluded.remaining_balance,
            next_billing_date = excluded.next_billing_date,
            updated_at = excluded.updated_at
    """, (device_id, subscription_id, plan, status, remaining_balance, next_billing_date, datetime.now(timezone.utc).isoformat()))
    conn.commit()
    conn.close()


def send_usage_event(device_id, transcript_id, duration_seconds):
    """Шлём событие в Dodo, чтобы Meter списал минуты с баланса подписки (или выставил
    overage, если баланс исчерпан). Если у устройства ещё нет клиента Dodo (никогда не
    оформляло подписку — всё ещё на бесплатных 30 минутах), просто ничего не отправляем."""
    dodo_customer_id = get_dodo_customer_id(device_id)
    if not dodo_customer_id:
        return
    minutes = max(1, int((duration_seconds + 59) // 60))  # округляем вверх до целой минуты
    try:
        dodo_client().usage_events.ingest(events=[{
            "event_id": transcript_id,  # тот же id при retry — по идее не задвоит событие
            "customer_id": dodo_customer_id,
            "event_name": USAGE_EVENT_NAME,
            "timestamp": datetime.now(timezone.utc).isoformat(),
            "metadata": {"minutes": str(minutes)},
        }])
    except Exception as e:  # noqa: BLE001
        print(f"[warn] failed to send usage event for {transcript_id}: {e}")


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
    u = get_usage_row(device_id)
    sub = get_subscription(device_id)
    return jsonify({
        "seconds_used": u["seconds_used"],
        "limit_seconds": FREE_LIMIT_SECONDS,
        "subscription": {
            "plan": sub["plan"],
            "status": sub["status"],
            "remaining_balance": sub["remaining_balance"],
            "next_billing_date": sub["next_billing_date"],
        } if sub else None,
    })


@app.route("/api/transcribe/file", methods=["POST"])
def transcribe_file():
    device_id = request.form.get("device_id", "")
    if not has_credit(device_id):
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
        json={"audio_url": audio_url, "speaker_labels": True, "speech_model": SPEECH_MODEL},
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
            # Бесплатные 30 минут по-прежнему считаем сами; если они уже исчерпаны и есть
            # активная подписка — send_usage_event списывает минуты через Dodo (включая overage).
            u = get_usage_row(device_id)
            remaining_free = max(0, FREE_LIMIT_SECONDS - u["seconds_used"])
            from_free = min(duration, remaining_free)
            if from_free > 0:
                add_free_seconds_used(device_id, from_free)
            from_subscription = duration - from_free
            if from_subscription > 0:
                send_usage_event(device_id, transcript_id, from_subscription)
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


@app.route("/api/subscribe", methods=["POST"])
def subscribe():
    data = request.get_json(force=True)
    device_id = data.get("device_id", "")
    plan = data.get("plan", "")

    if not device_id:
        return jsonify({"error": "no_device_id"}), 400
    if plan not in PLANS:
        return jsonify({"error": "invalid_plan"}), 400
    product_id = PLANS[plan]["product_id"]
    if not DODO_API_KEY or not product_id:
        return jsonify({"error": "payments_not_configured"}), 503

    try:
        session = dodo_client().checkout_sessions.create(
            product_cart=[{"product_id": product_id, "quantity": 1}],
            return_url=f"{request.url_root.rstrip('/')}/payment-success",
            metadata={"device_id": device_id, "plan": plan},
        )
    except Exception as e:  # noqa: BLE001
        return jsonify({"error": "dodo_error", "detail": str(e)}), 502

    checkout_url = session.checkout_url
    if not checkout_url:
        return jsonify({"error": "no_checkout_url"}), 502

    return jsonify({"checkout_url": checkout_url})


@app.route("/api/webhooks/dodo", methods=["POST"])
def dodo_webhook():
    raw_body = request.get_data()
    if DODO_WEBHOOK_KEY:
        try:
            from standardwebhooks.webhooks import Webhook
            wh = Webhook(DODO_WEBHOOK_KEY)
            wh.verify(raw_body, {
                "webhook-id": request.headers.get("webhook-id", ""),
                "webhook-signature": request.headers.get("webhook-signature", ""),
                "webhook-timestamp": request.headers.get("webhook-timestamp", ""),
            })
        except Exception as e:  # noqa: BLE001
            return jsonify({"error": "invalid_signature", "detail": str(e)}), 401

    payload = request.get_json(force=True)
    event_type = payload.get("type", "")

    # Обрабатываем ЛЮБОЕ событие subscription.* одинаково — в разных версиях доков Dodo
    # встречаются subscription.active / subscription.updated / subscription.renewed и т.д.,
    # но структура полезной нагрузки (data) одна и та же, включая remaining_balance.
    if event_type.startswith("subscription."):
        sub_data = payload.get("data", {})
        metadata = sub_data.get("metadata", {})
        device_id = metadata.get("device_id")
        plan = metadata.get("plan")
        customer = sub_data.get("customer", {})
        dodo_customer_id = customer.get("customer_id") if isinstance(customer, dict) else None

        if device_id and dodo_customer_id:
            upsert_customer(device_id, dodo_customer_id)

            remaining_balance = None
            credit_cart = sub_data.get("credit_entitlement_cart") or []
            if credit_cart:
                try:
                    remaining_balance = float(credit_cart[0].get("remaining_balance"))
                except (TypeError, ValueError):
                    remaining_balance = None

            upsert_subscription(
                device_id,
                sub_data.get("subscription_id"),
                plan,
                sub_data.get("status"),
                remaining_balance,
                sub_data.get("next_billing_date"),
            )

    return jsonify({"ok": True})


@app.route("/payment-success", methods=["GET"])
def payment_success():
    return """
    <html><body style="font-family:sans-serif;text-align:center;padding:60px 20px;">
        <h2>Оплата прошла успешно ✓</h2>
        <p>Можно закрыть эту вкладку и вернуться в расширение — баланс обновится за пару секунд.</p>
    </body></html>
    """


@app.route("/", methods=["GET"])
def health():
    return jsonify({"ok": True, "pdf_available": HTML is not None})


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)
