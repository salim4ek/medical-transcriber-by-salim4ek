"""
Telegram bot для транскрибера.
Принимает аудио → кладёт в input/ → транскрибер обрабатывает → бот шлёт docx обратно.
"""
import os
import re
import sys
import io
import json
import asyncio
import logging
import secrets
from pathlib import Path
from datetime import datetime, timezone

from telegram import Update, Message
from telegram.ext import (
    Application, CommandHandler, MessageHandler, ContextTypes, filters,
)
from telegram.request import HTTPXRequest

TG_MAX_LEN = 4000          # лимит сообщения Telegram (~4096 за вычетом запаса)

if hasattr(sys.stdout, "buffer"):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "buffer"):
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

BASE_DIR = Path(__file__).parent
CONFIG_PATH = BASE_DIR / "config.json"
STATE_PATH = BASE_DIR / "bot_state.json"
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"
TRANSCRIBER_LOG = BASE_DIR / "transcriber.log"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler(BASE_DIR / "bot.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
# Тише httpx/telegram debug-логов
logging.getLogger("httpx").setLevel(logging.WARNING)
logging.getLogger("telegram").setLevel(logging.WARNING)
log = logging.getLogger("bot")

ALLOWED_EXTS = {".mp3", ".wav", ".m4a", ".ogg", ".webm", ".mp4", ".aac", ".flac", ".opus"}
MAX_SIZE_MB = 20  # Telegram Bot API лимит для getFile — 20 МБ.
                  # Чтобы пропускать больше — нужен self-hosted telegram-bot-api (лимит 2 ГБ).


def load_state() -> dict:
    if STATE_PATH.exists():
        try:
            return json.loads(STATE_PATH.read_text(encoding="utf-8"))
        except Exception as e:
            log.warning(f"failed to load state: {e}")
    return {"pending": {}, "sent": {}}


def save_state(state: dict):
    STATE_PATH.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")


def _record_user(state: dict, update: Update):
    """Запоминаем username и chat_id любого, кто пишет боту. Нужно api.py чтобы
    смочь резолвить @username (из настроек desktop-клиента) в chat_id для доставки."""
    msg = update.message if update else None
    if not msg or not msg.from_user or not msg.chat_id:
        return
    user = msg.from_user
    users = state.setdefault("users", {})
    users[str(msg.chat_id)] = {
        "user_id": user.id,
        "username": (user.username or "").lower(),
        "first_name": user.first_name or "",
        "last_seen": datetime.now(timezone.utc).isoformat(),
    }
    save_state(state)


# ────────────────────── /start ────────────────────────────────────

async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    _record_user(load_state(), update)
    await update.message.reply_text(
        "👋 Привет! Это бот-транскрибер медицинских консультаций.\n\n"
        "Просто пришли аудиозапись медицинской консультации — распознаю речь, "
        "структурирую текст и пришлю готовый Word-документ.\n\n"
        "✅ Поддерживаемые форматы: mp3, wav, m4a, ogg, opus, webm, mp4, aac\n"
        f"📏 Максимальный размер: {MAX_SIZE_MB} МБ (ограничение Telegram)\n\n"
        "⏱ Обработка обычно занимает 1–5 минут."
    )


def _get_audio_info(message: Message):
    """Извлекает (file_id, suggested_ext, size) из сообщения если это аудио."""
    if message.voice:
        return message.voice.file_id, ".ogg", message.voice.file_size or 0
    if message.audio:
        ext = ".mp3"
        if message.audio.file_name:
            suffix = Path(message.audio.file_name).suffix.lower()
            if suffix in ALLOWED_EXTS:
                ext = suffix
        return message.audio.file_id, ext, message.audio.file_size or 0
    if message.document:
        suffix = ""
        if message.document.file_name:
            suffix = Path(message.document.file_name).suffix.lower()
        if suffix in ALLOWED_EXTS:
            return message.document.file_id, suffix, message.document.file_size or 0
        if message.document.mime_type and message.document.mime_type.startswith("audio/"):
            return message.document.file_id, suffix or ".mp3", message.document.file_size or 0
    if message.video_note:
        return message.video_note.file_id, ".mp4", message.video_note.file_size or 0
    return None, None, 0


async def on_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = update.message
    if not msg:
        return
    chat_id = msg.chat_id

    state = load_state()
    _record_user(state, update)

    file_id, ext, size = _get_audio_info(msg)
    if not file_id:
        await msg.reply_text(
            "❌ Я принимаю только аудиофайлы (mp3, wav, m4a, ogg, opus…).\n"
            "Пришли голосовое сообщение или аудиозапись."
        )
        return

    size_mb = size / (1024 * 1024) if size else 0
    if size_mb > MAX_SIZE_MB:
        await msg.reply_text(
            f"❌ Файл {size_mb:.1f} МБ — это больше лимита Telegram для ботов ({MAX_SIZE_MB} МБ).\n\n"
            "Как обойти:\n"
            "• 🎙 Записать прямо в Telegram как голосовое (зажми микрофон в чате) — он автосжимает\n"
            "• 📉 Сжать аудио до Opus/M4A 64 kbps в приложении-конвертере (1 час → ~15–20 МБ)\n"
            "• ✂️ Разбить запись на 2–3 части и прислать по очереди"
        )
        return

    # Microsecond timestamp + 4-hex-char random suffix → файлы не коллизятся даже при залпе
    ts = int(datetime.now().timestamp() * 1_000_000)
    suffix = secrets.token_hex(2)
    filename = f"tg_{chat_id}_{ts}_{suffix}{ext}"
    target = INPUT_DIR / filename

    try:
        tg_file = await context.bot.get_file(file_id)
        await tg_file.download_to_drive(custom_path=str(target))
    except Exception as e:
        log.error(f"download failed for chat {chat_id}: {e}")
        await msg.reply_text(f"⚠ Не смог скачать файл: {type(e).__name__}")
        return

    actual_size = target.stat().st_size / (1024 * 1024)
    log.info(f"received audio from chat {chat_id}: {filename} ({actual_size:.1f} MB)")

    ack = await msg.reply_text(f"📥 Принял файл ({actual_size:.1f} МБ).")

    state = load_state()
    state.setdefault("pending", {})[filename] = {
        "chat_id": chat_id,
        "user_msg_id": msg.message_id,
        "ack_msg_id": ack.message_id,
        "received_at": datetime.now(timezone.utc).isoformat(),
        "notified": ["received"],
    }
    save_state(state)


def _find_docxs_for(stem: str) -> list:
    """Ищет ВСЕ docx в output/ с указанным stem в имени.
    Возвращает список (может быть summary + расшифровка)."""
    if not OUTPUT_DIR.exists():
        return []
    return sorted(f for f in OUTPUT_DIR.glob("*.docx") if stem in f.name)


async def watch_output(app: Application):
    """Бэкграунд-задача: сканирует output/ и отправляет готовые docx.
    Когда находит docx-сводку — ждёт пока появится и docx-расшифровка (или таймаут),
    после чего отправляет ВСЕ найденные файлы одним пакетом."""
    log.info("output watcher started")
    while True:
        try:
            state = load_state()
            pending = dict(state.get("pending", {}))
            for audio_name, info in pending.items():
                stem = Path(audio_name).stem  # tg_<chat>_<ts> или api_<job>
                docxs = _find_docxs_for(stem)
                if not docxs:
                    continue
                # Если есть хотя бы одна сводка но НЕТ расшифровки — ждём появления
                # (транскрибер пишет расшифровку сразу после summary, разница в секундах).
                names = [d.name for d in docxs]
                has_summary = any("_расшифровка" not in n for n in names)
                has_transcript = any("_расшифровка" in n for n in names)
                if has_summary and not has_transcript:
                    # подождём один цикл — может транскрипт-файл только что создаётся
                    info_pending = info.setdefault("wait_transcript_cycles", 0)
                    if info_pending < 3:
                        info["wait_transcript_cycles"] = info_pending + 1
                        save_state(state)
                        continue
                    # после 3 циклов (≈24с) — отправляем что есть

                chat_id = info["chat_id"]
                sent_files = []
                try:
                    for docx in docxs:
                        log.info(f"sending {docx.name} to chat {chat_id}")
                        is_transcript = "_расшифровка" in docx.name
                        caption = "📝 Сырая расшифровка" if is_transcript else "✅ Сводка приёма"
                        with open(docx, "rb") as f:
                            await app.bot.send_document(
                                chat_id=chat_id,
                                document=f,
                                filename=docx.name,
                                caption=caption,
                            )
                        sent_files.append(docx.name)
                    state["pending"].pop(audio_name, None)
                    state.setdefault("sent", {})[audio_name] = {
                        "chat_id": chat_id,
                        "files": sent_files,
                        "sent_at": datetime.now(timezone.utc).isoformat(),
                    }
                    save_state(state)
                except Exception as e:
                    log.error(f"send_document failed for chat {chat_id}: {type(e).__name__}: {e}")
        except Exception as e:
            log.error(f"watch loop error: {type(e).__name__}: {e}")
        await asyncio.sleep(8)


# Маркеры состояний из transcriber.log → (state_key, telegram-сообщение)
STATE_MESSAGES = {
    "stt_sent": "📤 Отправил на расшифровку.",
    "llm_start": "🤖 Распознавание готово, отправил в обработку.\n⏱ Ожидайте до 15 мин.",
}

# Позиция чтения transcriber.log (между итерациями), на первом запуске = EOF (не уведомлять про старое)
_log_pos = {"value": None}


async def _maybe_notify(app: Application, state: dict, filename: str, key: str):
    pending = state.get("pending", {})
    info = pending.get(filename)
    if not info:
        return False
    notified = info.setdefault("notified", [])
    if key in notified:
        return False
    msg = STATE_MESSAGES.get(key)
    if not msg:
        return False
    try:
        await app.bot.send_message(chat_id=info["chat_id"], text=msg)
        notified.append(key)
        log.info(f"notified {info['chat_id']} for {filename}: {key}")
        return True
    except Exception as e:
        log.warning(f"notify failed for chat {info['chat_id']}: {e}")
        return False


async def watch_log(app: Application):
    """Tail transcriber.log: detect state transitions per pending file, notify chat."""
    log.info("log watcher started")
    while True:
        try:
            if not TRANSCRIBER_LOG.exists():
                await asyncio.sleep(5)
                continue
            size = TRANSCRIBER_LOG.stat().st_size
            # Initialize at EOF on first run so we don't replay history
            if _log_pos["value"] is None:
                _log_pos["value"] = size
                await asyncio.sleep(3)
                continue
            # Log was truncated (cleanup.sh) — reset
            if size < _log_pos["value"]:
                _log_pos["value"] = 0

            with open(TRANSCRIBER_LOG, "rb") as f:
                f.seek(_log_pos["value"])
                new_bytes = f.read()
                _log_pos["value"] = f.tell()
            if not new_bytes:
                await asyncio.sleep(3)
                continue
            new_text = new_bytes.decode("utf-8", errors="replace")

            state = load_state()
            changed = False
            current_file = None
            for line in new_text.splitlines():
                # Detect file context из заголовка "Файл: X"
                m = re.search(r"Файл:\s+(\S.+?\.\S+)", line)
                if m:
                    current_file = m.group(1).strip()
                    continue
                # STT Отправляю: <filename> ... — извлекаем имя из самой строки
                m = re.search(r"\[STT\] Отправляю:\s+(\S.+?\.\S+?)\s+\(", line)
                if m:
                    fname = m.group(1).strip()
                    if await _maybe_notify(app, state, fname, "stt_sent"):
                        changed = True
                    continue
                # [LLM] Обрабатываю файл «<filename>» ... — параллельный режим, явное имя в строке
                m = re.search(r"\[LLM\] Обрабатываю файл «(.+?)»", line)
                if m:
                    fname = m.group(1).strip()
                    if await _maybe_notify(app, state, fname, "llm_start"):
                        changed = True
                    continue
                # Старый формат без имени — fallback на current_file (sequential режим)
                if "[LLM] Обрабатываю" in line and current_file:
                    if await _maybe_notify(app, state, current_file, "llm_start"):
                        changed = True
            if changed:
                save_state(state)
        except Exception as e:
            log.error(f"watch_log error: {type(e).__name__}: {e}")
        await asyncio.sleep(3)


async def post_init(app: Application):
    app.create_task(watch_output(app))
    app.create_task(watch_log(app))


def main():
    INPUT_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    if not CONFIG_PATH.exists():
        log.error(f"config.json not found at {CONFIG_PATH}")
        sys.exit(1)
    config = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    token = config.get("telegram_bot_token")
    if not token:
        log.error("telegram_bot_token not in config.json")
        sys.exit(1)

    builder = Application.builder().token(token).post_init(post_init)

    # Если в config задан proxy_url — пускаем и HTTP-клиента, и polling через прокси
    # (нужно для серверов где api.telegram.org заблокирован — типичный РФ-случай).
    # Если TELEGRAM_TRY_DIRECT=1 в env и прокси мёртв — пробуем прямое соединение.
    proxy_url = config.get("telegram_proxy")
    use_proxy = bool(proxy_url)
    if proxy_url and os.environ.get("TELEGRAM_TRY_DIRECT") == "1":
        # Проверяем доступность прокси (быстрая проба)
        try:
            host = proxy_url.split("@", 1)[-1].split("/")[0].split(":")
            ph, pp = host[0], int(host[1]) if len(host) > 1 else 3128
            import socket as _s
            _s.create_connection((ph, pp), timeout=2).close()
            log.info(f"proxy {ph}:{pp} reachable, using it")
        except Exception as e:
            log.warning(f"proxy {proxy_url} unreachable ({e}); falling back to direct")
            use_proxy = False
    if use_proxy:
        # маскируем креды в логе
        safe = proxy_url.split("@", 1)[-1] if "@" in proxy_url else proxy_url
        log.info(f"using Telegram proxy: {safe}")
        builder = (
            builder
            .request(HTTPXRequest(proxy=proxy_url, connection_pool_size=8))
            .get_updates_request(HTTPXRequest(proxy=proxy_url, connection_pool_size=8))
        )
    else:
        log.info("using direct Telegram connection (no proxy)")

    app = builder.build()
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("help", cmd_start))
    app.add_handler(MessageHandler(filters.ALL & ~filters.COMMAND, on_message))

    log.info("Bot started. Polling...")
    app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()
