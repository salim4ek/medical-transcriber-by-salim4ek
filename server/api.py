"""
HTTP API для транскрибера: загрузка аудио, статус, скачивание docx.
Bearer-токен авторизация (api_shared_key в config.json).
"""
import os
import re
import json
import sys
import io
import secrets
import logging
from pathlib import Path
from datetime import datetime, timezone

from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Depends
from fastapi.responses import FileResponse
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from typing import Optional

if hasattr(sys.stdout, "buffer"):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

BASE_DIR = Path(__file__).parent
CONFIG_PATH = BASE_DIR / "config.json"
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"
STATE_PATH = BASE_DIR / "api_state.json"
TRANSCRIBER_LOG = BASE_DIR / "transcriber.log"
BOT_STATE_PATH = BASE_DIR / "bot_state.json"  # shared с ботом для маппинга username→chat_id

ALLOWED_EXTS = {".mp3", ".wav", ".m4a", ".ogg", ".webm", ".mp4", ".aac", ".flac", ".opus"}
MAX_FILE_BYTES = 500 * 1024 * 1024  # 500 MB

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger("api")

INPUT_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

config = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
API_KEY = config.get("api_shared_key") or os.environ.get("API_SHARED_KEY")
if not API_KEY:
    raise RuntimeError("api_shared_key not in config.json and API_SHARED_KEY env not set")

app = FastAPI(title="Transcriber API", version="1.0")
security = HTTPBearer(auto_error=True)


def auth(creds: HTTPAuthorizationCredentials = Depends(security)):
    if not secrets.compare_digest(creds.credentials, API_KEY):
        raise HTTPException(status_code=401, detail="Invalid token")


def load_state() -> dict:
    if STATE_PATH.exists():
        try:
            return json.loads(STATE_PATH.read_text(encoding="utf-8"))
        except Exception as e:
            log.warning(f"failed to load state: {e}")
    return {"jobs": {}}


def save_state(s: dict):
    STATE_PATH.write_text(json.dumps(s, ensure_ascii=False, indent=2), encoding="utf-8")


def _list_docx(job_id: str):
    """Все docx в output/ для данного job_id (сводка + расшифровка)."""
    if not OUTPUT_DIR.exists():
        return []
    return sorted(f for f in OUTPUT_DIR.glob("*.docx") if f"api_{job_id}" in f.name)


def find_docx(job_id: str):
    """Возвращает docx-СВОДКУ (не расшифровку). Детерминированно: сначала
    отбрасываем файлы с суффиксом _расшифровка, берём первый из оставшихся."""
    files = _list_docx(job_id)
    summaries = [f for f in files if "_расшифровка" not in f.name]
    if summaries:
        return summaries[0]
    return files[0] if files else None


def find_transcript(job_id: str):
    """Возвращает docx-РАСШИФРОВКУ (companion-файл), если она уже создана."""
    for f in _list_docx(job_id):
        if "_расшифровка" in f.name:
            return f
    return None


STATE_LABELS = {
    "queued": "В очереди",
    "stt": "Расшифровка",
    "llm": "Обработка ИИ",
    "done": "Готово",
    "failed": "Ошибка",
    "cancelled": "Отменено",
}


def queue_ahead_count(my_filename: str) -> int:
    """How many other audio files are alphabetically before this one (i.e. ahead in queue).
    Transcriber processes sorted by name, names include timestamps, so this is a usable
    queue-position estimate. 0 = this file is currently being processed (or next)."""
    if not INPUT_DIR.exists():
        return 0
    count = 0
    for f in INPUT_DIR.iterdir():
        if not f.is_file():
            continue
        if f.suffix.lower() not in ALLOWED_EXTS:
            continue
        if f.name == my_filename:
            continue
        if f.name < my_filename:
            count += 1
    return count


def detect_state(job_id: str, filename: str) -> dict:
    """Определяет состояние обработки задачи."""
    docx = find_docx(job_id)
    if docx:
        return {"state": "done", "docx_name": docx.name, "queue_ahead": 0}

    if not TRANSCRIBER_LOG.exists():
        return {"state": "queued", "queue_ahead": queue_ahead_count(filename)}

    try:
        text = TRANSCRIBER_LOG.read_text(encoding="utf-8", errors="replace")
    except OSError:
        return {"state": "queued", "queue_ahead": queue_ahead_count(filename)}

    stem = Path(filename).stem  # api_<job_id>
    idx = text.rfind(stem)
    if idx < 0:
        # transcriber ещё не взял файл в работу
        return {"state": "queued", "queue_ahead": queue_ahead_count(filename)}

    # Берём tail С НАЧАЛА СТРОКИ где упомянут файл — иначе пропустим маркер
    # state'а, который идёт в той же строке ДО имени файла
    # (например: "[LLM] Обрабатываю файл «api_xxx.m4a» ...")
    line_start = text.rfind("\n", 0, idx) + 1
    tail = text[line_start:]
    if "[ГОТОВО] Обработка завершена" in tail:
        return {"state": "done", "queue_ahead": 0}
    # Note: больше не возвращаем failed для "Не медицинская консультация" /
    # "Нет консультаций в ответе" — теперь транскрибер создаёт fallback docx с сырой
    # расшифровкой, и docx найдётся find_docx() выше как state=done.
    if "[ОШИБКА" in tail:
        return {
            "state": "failed", "queue_ahead": 0,
            "error_reason": "Ошибка обработки файла",
        }
    if "[!] Модель не вернула" in tail or "Все попытки исчерпаны" in tail:
        return {
            "state": "failed", "queue_ahead": 0,
            "error_reason": "Модель не вернула валидный JSON после нескольких попыток",
        }
    # Generic "moved on to next file" marker after our file but no docx -> terminal
    # (catches edge cases we didn't enumerate)
    if "[МОНИТОРИНГ] Продолжаю" in tail:
        return {
            "state": "failed", "queue_ahead": 0,
            "error_reason": "Файл обработан, но результат не сформирован",
        }
    if "[LLM] Обрабатываю" in tail:
        return {"state": "llm", "queue_ahead": 0}
    if "[STT]" in tail:
        return {"state": "stt", "queue_ahead": 0}
    return {"state": "queued", "queue_ahead": queue_ahead_count(filename)}


def _register_tg_delivery(filename: str, tg_username: str) -> Optional[int]:
    """Резолвит @username в chat_id через bot_state.json (там бот записывает
    всех кто ему когда-либо писал). Если найден — добавляет pending-запись в
    bot_state, и бот автоматически отправит docx-сводку + расшифровку этому
    юзеру когда они появятся в output/. Возвращает chat_id или None."""
    if not BOT_STATE_PATH.exists():
        log.warning("bot_state.json не найден — пропускаю tg-доставку")
        return None
    target = tg_username.lstrip("@").lower().strip()
    if not target:
        return None
    try:
        with open(BOT_STATE_PATH, "r", encoding="utf-8") as f:
            state = json.load(f)
    except Exception as e:
        log.warning(f"не смог прочитать bot_state.json: {e}")
        return None
    users = state.get("users", {}) or {}
    chat_id = None
    for cid, info in users.items():
        if (info.get("username") or "").lower() == target:
            try:
                chat_id = int(cid)
            except ValueError:
                continue
            break
    if chat_id is None:
        log.warning(
            f"tg_username @{target} не найден в bot.users — "
            "пользователь сначала должен написать боту (/start)"
        )
        return None
    state.setdefault("pending", {})[filename] = {
        "chat_id": chat_id,
        "user_msg_id": 0,
        "ack_msg_id": 0,
        "received_at": datetime.now(timezone.utc).isoformat(),
        "notified": ["received"],
        "tg_username": tg_username,
    }
    try:
        with open(BOT_STATE_PATH, "w", encoding="utf-8") as f:
            json.dump(state, f, ensure_ascii=False, indent=2)
        log.info(f"tg доставка: {filename} → @{target} (chat {chat_id})")
        return chat_id
    except Exception as e:
        log.error(f"не смог записать bot_state.json: {e}")
        return None


@app.get("/ping")
def ping():
    """Health-check без авторизации."""
    return {"ok": True, "time": datetime.now(timezone.utc).isoformat()}


@app.post("/upload")
async def upload(
    file: UploadFile = File(...),
    tg_username: Optional[str] = Form(None),
    _=Depends(auth),
):
    """Принимает аудио, возвращает job_id.
    Если передан tg_username (@username) — регистрирует доставку через Telegram-бот,
    при условии что пользователь раньше писал боту (любое сообщение)."""
    orig_name = file.filename or "upload"
    ext = Path(orig_name).suffix.lower()
    if ext not in ALLOWED_EXTS:
        raise HTTPException(
            400,
            f"Расширение {ext!r} не поддерживается. Допустимы: {sorted(ALLOWED_EXTS)}",
        )

    # job_id: 12 случайных hex-символов, гарантированно уникально
    job_id = secrets.token_hex(6)
    target_name = f"api_{job_id}{ext}"
    target = INPUT_DIR / target_name

    written = 0
    try:
        with open(target, "wb") as out:
            while True:
                chunk = await file.read(1024 * 1024)
                if not chunk:
                    break
                written += len(chunk)
                if written > MAX_FILE_BYTES:
                    out.close()
                    target.unlink(missing_ok=True)
                    raise HTTPException(
                        413,
                        f"Файл превышает {MAX_FILE_BYTES // 1024 // 1024} МБ",
                    )
                out.write(chunk)
    except HTTPException:
        raise
    except Exception as e:
        target.unlink(missing_ok=True)
        log.error(f"upload failed: {type(e).__name__}: {e}")
        raise HTTPException(500, f"Ошибка сохранения файла: {type(e).__name__}")

    state = load_state()
    state.setdefault("jobs", {})[job_id] = {
        "filename": target_name,
        "original_name": orig_name,
        "size": written,
        "received_at": datetime.now(timezone.utc).isoformat(),
        "tg_username": tg_username,
    }
    save_state(state)
    log.info(f"upload OK: job={job_id} file={target_name} size={written}")

    # Если запрошена TG-доставка — пытаемся зарегистрировать
    tg_chat_id = None
    if tg_username:
        tg_chat_id = _register_tg_delivery(target_name, tg_username)

    return {
        "job_id": job_id,
        "filename": target_name,
        "size": written,
        "tg_delivery": "enabled" if tg_chat_id else ("not_registered" if tg_username else None),
    }


@app.get("/status/{job_id}")
def status(job_id: str, _=Depends(auth)):
    state = load_state()
    job = state.get("jobs", {}).get(job_id)
    if not job:
        raise HTTPException(404, "Неизвестный job_id")

    if job.get("cancelled"):
        s = {"state": "cancelled", "queue_ahead": 0}
    else:
        s = detect_state(job_id, job["filename"])

    s["job_id"] = job_id
    s["original_name"] = job.get("original_name")
    s["received_at"] = job.get("received_at")
    s["size"] = job.get("size")
    s["label"] = STATE_LABELS.get(s["state"], s["state"])
    return s


@app.delete("/jobs/{job_id}")
def cancel_job(job_id: str, _=Depends(auth)):
    """Cancel a job: remove the input audio (if not yet processed), purge output docx,
    mark cancelled in state. Already-running transcriber stages can't be force-killed
    but their output is discarded."""
    state = load_state()
    job = state.get("jobs", {}).get(job_id)
    if not job:
        raise HTTPException(404, "Неизвестный job_id")

    audio_path = INPUT_DIR / job["filename"]
    if audio_path.exists():
        try:
            audio_path.unlink()
        except OSError as e:
            log.warning(f"cannot unlink {audio_path}: {e}")

    # Remove any output that already landed
    for d in OUTPUT_DIR.glob(f"*api_{job_id}*.docx"):
        try:
            d.unlink()
        except OSError:
            pass

    job["cancelled"] = True
    save_state(state)
    log.info(f"job {job_id} cancelled")
    return {"ok": True, "job_id": job_id, "state": "cancelled"}


@app.get("/result/{job_id}")
def result(job_id: str, part: str = "summary", _=Depends(auth)):
    """Отдаёт docx результата.
    part=summary (по умолчанию) — структурированная сводка (без расшифровки);
    part=transcript — сырая расшифровка speech2text отдельным файлом.
    Старые клиенты без параметра part получают сводку (обратная совместимость)."""
    state = load_state()
    job = state.get("jobs", {}).get(job_id)
    if not job:
        raise HTTPException(404, "Неизвестный job_id")

    if part == "transcript":
        docx = find_transcript(job_id)
        if not docx or not docx.exists():
            raise HTTPException(404, "Расшифровка ещё не готова или была удалена очисткой")
    else:
        docx = find_docx(job_id)
        if not docx or not docx.exists():
            raise HTTPException(404, "Результат ещё не готов или был удалён очисткой")

    return FileResponse(
        path=str(docx),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=docx.name,
    )
