"""
Transcriber Client (PyQt6) — manual + auto-watch mode.

Запуск:        python main.py
Сборка .exe:   pyinstaller --noconfirm --onefile --windowed --name "TranscriberClient" main.py
"""
import json
import math
import sys
import time
from dataclasses import dataclass, field, asdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional, List, Dict

import requests
from PyQt6.QtCore import (
    Qt, QThread, QSettings, pyqtSignal, QTimer, QPropertyAnimation,
    QEasingCurve, QPoint, QObject,
)
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QFileDialog, QProgressBar, QStatusBar,
    QMessageBox, QDialog, QLineEdit, QFormLayout, QDialogButtonBox,
    QFrame, QGraphicsOpacityEffect, QCheckBox, QListWidget,
    QListWidgetItem, QToolButton, QInputDialog, QSystemTrayIcon,
    QMenu, QStyle,
)
from PyQt6.QtGui import QAction, QPainter, QColor, QIcon

APP_NAME = "Transcriber Client v4"
APP_VERSION = "4.0"
ORG_NAME = "Medical Transcriber"
DEFAULT_SERVER = "http://YOUR_SERVER_HOST:8000"
DEFAULT_API_KEY = "YOUR_API_SHARED_KEY"

# Пароль для разблокировки правки/копирования настроек сервера.
# Защищает от случайных правок коллегой; не криптостойкая защита.
# Задаётся при сборке клиента — в репозитории заглушка.
UNLOCK_PASSWORD = "CHANGE_ME"

ALLOWED_EXTS = {".mp3", ".wav", ".m4a", ".ogg", ".opus", ".webm", ".mp4", ".aac", ".flac"}
ALLOWED_FILTERS = (
    "Аудио (*.mp3 *.wav *.m4a *.ogg *.opus *.webm *.mp4 *.aac *.flac);;Все файлы (*.*)"
)

# Polling
WATCH_INTERVAL_MS = 10_000      # how often to scan watch folder
STABLE_CHECK_DELAY_S = 3        # wait this long to confirm file isn't still being written
POLL_STATUS_MS = 5_000          # how often to poll /status during processing

STATE_LABELS = {
    "preparing": "Подготовка",
    "uploading": "Загрузка",
    "queued":    "В очереди",
    "stt":       "Распознавание",
    "llm":       "Обработка ИИ",
    "downloading": "Скачивание",
    "done":      "Готово",
    "failed":    "Ошибка",
    "cancelled": "Отменено",
    "restoring": "Проверка статуса…",
}

# job states that can be cancelled
CANCELLABLE_STATES = {"preparing", "uploading", "queued", "stt", "llm", "downloading", "restoring"}
# states that mean job is no longer active
TERMINAL_STATES = {"done", "failed", "cancelled"}

# ────────────────────── Style ──────────────────────

QSS = """
/* Палитра: тёмно-зелёная (фирменная клиники NN+).
   Основной  #1B7C70, темнее #145E55, светлый фон #ECFDF5,
   акцент-светлый #D1FAE5, текст-зелёный-тёмный #065F46. */
* { font-family: "Segoe UI", "SF Pro Display", system-ui, sans-serif; color: #1F2937; }
QMainWindow, QWidget#root { background: #F4F7F5; }
QLabel#header { font-size: 22px; font-weight: 600; color: #064E3B; }
QLabel#subheader { font-size: 13px; color: #6B7280; }
QLabel#sectionTitle { font-size: 13px; font-weight: 600; color: #065F46; }
QLabel#counter { font-size: 12px; color: #6B7280; }

QFrame#card { background: #FFFFFF; border-radius: 12px; border: 1px solid #D1E7E0; }
QFrame#dropZone { background: #FAFBFA; border: 2px dashed #D1D5DB; border-radius: 12px; }
QFrame#dropZone[hover="true"] { background: #ECFDF5; border: 2px dashed #1B7C70; }
QFrame#dropZone[fileLoaded="true"] { background: #D1FAE5; border: 2px solid #1B7C70; }
QFrame#autoPanel { background: #FFFFFF; border: 1px solid #D1E7E0; border-radius: 10px; }
QFrame#autoPanel[active="true"] { background: #ECFDF5; border: 1px solid #1B7C70; }

QLabel#dropLabel { font-size: 15px; color: #4B5563; }
QLabel#fileNameLabel { font-size: 16px; font-weight: 600; color: #064E3B; }

QPushButton {
    background: #FFFFFF; border: 1px solid #D1D5DB; border-radius: 8px;
    padding: 10px 18px; font-size: 14px; font-weight: 500; color: #374151;
}
QPushButton:hover { background: #F9FAFB; border-color: #9CA3AF; }
QPushButton:pressed { background: #F3F4F6; }
QPushButton:disabled { background: #F3F4F6; color: #9CA3AF; border-color: #E5E7EB; }

QPushButton#primary {
    background: #1B7C70; color: #FFFFFF; border: 1px solid #1B7C70; font-weight: 600;
}
QPushButton#primary:hover { background: #145E55; border-color: #145E55; }
QPushButton#primary:pressed { background: #0F4A42; }
QPushButton#primary:disabled { background: #A3D5CE; color: #FFFFFF; border-color: #A3D5CE; }

QPushButton#toggleOn {
    background: #1B7C70; color: #FFFFFF; border: 1px solid #1B7C70; font-weight: 600;
}
QPushButton#toggleOn:hover { background: #145E55; }

QProgressBar {
    background: #E5E7EB; border: none; border-radius: 5px;
    text-align: center; height: 10px; color: transparent;
}
QProgressBar::chunk { background: #1B7C70; border-radius: 5px; }

QStatusBar { background: #FFFFFF; border-top: 1px solid #D1E7E0;
             color: #6B7280; font-size: 12px; padding: 4px 8px; }

QMenuBar { background: #FFFFFF; border-bottom: 1px solid #D1E7E0; padding: 4px 6px; }
QMenuBar::item { background: transparent; padding: 6px 12px; border-radius: 6px; }
QMenuBar::item:selected { background: #ECFDF5; color: #065F46; }
QMenu { background: #FFFFFF; border: 1px solid #D1E7E0; border-radius: 8px; padding: 4px; }
QMenu::item { padding: 8px 24px; border-radius: 4px; }
QMenu::item:selected { background: #ECFDF5; color: #065F46; }

QLineEdit {
    background: #FFFFFF; border: 1px solid #D1D5DB; border-radius: 6px;
    padding: 8px 10px; font-size: 13px;
}
QLineEdit:focus { border-color: #1B7C70; }

QLabel#statusBubble {
    padding: 12px 16px; border-radius: 8px; font-size: 13px; background: transparent;
}
QLabel#statusBubble[kind="info"]    { background: #ECFDF5; color: #065F46; }
QLabel#statusBubble[kind="success"] { background: #D1FAE5; color: #064E3B; }
QLabel#statusBubble[kind="error"]   { background: #FEF2F2; color: #991B1B; }
QLabel#statusBubble[kind="warn"]    { background: #FFFBEB; color: #92400E; }

QListWidget {
    background: #FFFFFF; border: 1px solid #D1E7E0; border-radius: 8px;
    padding: 4px; font-size: 12px;
}
QListWidget::item { padding: 6px 8px; border-radius: 4px; color: #1F2937; }
QListWidget::item:selected { background: #D1FAE5; color: #065F46; }
QListWidget::item:hover { background: #F0FDF4; }

QDialog { background: #F4F7F5; }
"""


# ────────────────────── Settings ──────────────────────

@dataclass
class Settings:
    server_url: str
    api_key: str
    watch_folder: str
    output_folder: str
    auto_mode: bool
    minimize_to_tray: bool
    autostart_windows: bool
    tg_username: str

    @classmethod
    def load(cls) -> "Settings":
        s = QSettings(ORG_NAME, APP_NAME)
        return cls(
            server_url=s.value("server_url", DEFAULT_SERVER, type=str),
            api_key=s.value("api_key", DEFAULT_API_KEY, type=str),
            watch_folder=s.value("watch_folder", "", type=str),
            output_folder=s.value("output_folder", "", type=str),
            auto_mode=s.value("auto_mode", False, type=bool),
            minimize_to_tray=s.value("minimize_to_tray", True, type=bool),
            autostart_windows=s.value("autostart_windows", False, type=bool),
            tg_username=s.value("tg_username", "", type=str),
        )

    def save(self):
        s = QSettings(ORG_NAME, APP_NAME)
        s.setValue("server_url", self.server_url)
        s.setValue("api_key", self.api_key)
        s.setValue("watch_folder", self.watch_folder)
        s.setValue("output_folder", self.output_folder)
        s.setValue("auto_mode", self.auto_mode)
        s.setValue("minimize_to_tray", self.minimize_to_tray)
        s.setValue("autostart_windows", self.autostart_windows)
        s.setValue("tg_username", self.tg_username)


# ────────────────────── Autostart helpers ──────────────────────

AUTOSTART_REG_NAME = "TranscriberClient"
AUTOSTART_REG_PATH = r"Software\Microsoft\Windows\CurrentVersion\Run"


def is_frozen_exe() -> bool:
    """True when running from a PyInstaller-built .exe (not python main.py)."""
    return getattr(sys, "frozen", False)


def get_exe_path() -> str:
    """Path to the running exe (or python.exe in dev mode)."""
    return sys.executable


def set_windows_autostart(enable: bool) -> tuple[bool, str]:
    """Add/remove TranscriberClient from Windows startup. Returns (ok, message)."""
    if sys.platform != "win32":
        return False, "Только для Windows"
    if not is_frozen_exe():
        return False, "Автозапуск работает только из собранного .exe (не из dev-режима)"
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, AUTOSTART_REG_PATH, 0,
                            winreg.KEY_SET_VALUE) as key:
            if enable:
                exe = get_exe_path()
                value = f'"{exe}" --minimized'
                winreg.SetValueEx(key, AUTOSTART_REG_NAME, 0, winreg.REG_SZ, value)
                return True, f"Автозапуск включён: {exe}"
            else:
                try:
                    winreg.DeleteValue(key, AUTOSTART_REG_NAME)
                except FileNotFoundError:
                    pass
                return True, "Автозапуск выключен"
    except Exception as e:
        return False, f"Ошибка реестра: {type(e).__name__}: {e}"


def is_windows_autostart_set() -> bool:
    if sys.platform != "win32":
        return False
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, AUTOSTART_REG_PATH, 0,
                            winreg.KEY_READ) as key:
            winreg.QueryValueEx(key, AUTOSTART_REG_NAME)
            return True
    except (FileNotFoundError, OSError):
        return False


class SettingsDialog(QDialog):
    # Mask string shown when fields are locked
    LOCKED_DISPLAY = "••••••••••••••••"

    def __init__(self, parent, settings: Settings):
        super().__init__(parent)
        self.setWindowTitle("Настройки")
        self.setMinimumWidth(560)

        self._settings = settings
        self._unlocked = False

        layout = QFormLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(20, 20, 20, 20)

        # ── Server URL (locked by default) ──
        self.server_edit = QLineEdit(self.LOCKED_DISPLAY)
        self.server_edit.setReadOnly(True)
        self.server_copy_btn = QPushButton("📋")
        self.server_copy_btn.setToolTip("Скопировать URL")
        self.server_copy_btn.setEnabled(False)
        self.server_copy_btn.setMaximumWidth(36)
        self.server_copy_btn.clicked.connect(self._copy_server)
        srv_row = QHBoxLayout()
        srv_row.addWidget(self.server_edit, 1)
        srv_row.addWidget(self.server_copy_btn, 0)
        srv_wrap = QWidget(); srv_wrap.setLayout(srv_row)
        layout.addRow("URL сервера:", srv_wrap)

        # ── API key (locked by default) ──
        self.key_edit = QLineEdit(self.LOCKED_DISPLAY)
        self.key_edit.setReadOnly(True)
        self.key_copy_btn = QPushButton("📋")
        self.key_copy_btn.setToolTip("Скопировать API ключ")
        self.key_copy_btn.setEnabled(False)
        self.key_copy_btn.setMaximumWidth(36)
        self.key_copy_btn.clicked.connect(self._copy_key)
        key_row = QHBoxLayout()
        key_row.addWidget(self.key_edit, 1)
        key_row.addWidget(self.key_copy_btn, 0)
        key_wrap = QWidget(); key_wrap.setLayout(key_row)
        layout.addRow("API ключ:", key_wrap)

        # ── Unlock button ──
        self.unlock_btn = QPushButton("🔒  Изменить настройки сервера (нужен пароль)")
        self.unlock_btn.clicked.connect(self._try_unlock)
        layout.addRow("", self.unlock_btn)

        # ── Watch folder (always editable) ──
        self.watch_edit = QLineEdit(settings.watch_folder)
        watch_row = QHBoxLayout()
        watch_btn = QPushButton("Обзор…")
        watch_btn.clicked.connect(self._pick_watch)
        watch_row.addWidget(self.watch_edit, 1)
        watch_row.addWidget(watch_btn, 0)
        watch_wrap = QWidget(); watch_wrap.setLayout(watch_row)
        layout.addRow("Watch-папка:", watch_wrap)

        # ── Output folder (always editable) ──
        self.output_edit = QLineEdit(settings.output_folder)
        out_row = QHBoxLayout()
        out_btn = QPushButton("Обзор…")
        out_btn.clicked.connect(self._pick_output)
        out_row.addWidget(self.output_edit, 1)
        out_row.addWidget(out_btn, 0)
        out_wrap = QWidget(); out_wrap.setLayout(out_row)
        layout.addRow("Папка для docx:", out_wrap)

        hint = QLabel(
            "Watch-папка: куда врач кладёт аудио. Клиент сканит её каждые 10 сек.\n"
            "Папка для docx: куда автоматически сохраняются готовые расшифровки.\n"
            "После успешной обработки аудио в watch-папке удаляется."
        )
        hint.setStyleSheet("color: #6B7280; font-size: 12px;")
        hint.setWordWrap(True)
        layout.addRow(hint)

        # ── Telegram-доставка ──
        sep_tg = QLabel("─" * 40)
        sep_tg.setStyleSheet("color: #D1D5DB;")
        layout.addRow(sep_tg)

        self.tg_username_edit = QLineEdit(settings.tg_username)
        self.tg_username_edit.setPlaceholderText("@username (например @your_username)")
        layout.addRow("Telegram username:", self.tg_username_edit)

        tg_hint = QLabel(
            "Если задан — все статусы и готовые файлы будут дублироваться в Telegram-бот "
            "@NNTRANS_bot. Перед использованием напиши боту /start один раз, "
            "чтобы он узнал твой chat_id."
        )
        tg_hint.setStyleSheet("color: #6B7280; font-size: 11px;")
        tg_hint.setWordWrap(True)
        layout.addRow(tg_hint)

        # ── Background mode + autostart ──
        sep = QLabel("─" * 40)
        sep.setStyleSheet("color: #D1D5DB;")
        layout.addRow(sep)

        self.tray_cb = QCheckBox("Сворачивать в трей при закрытии окна")
        self.tray_cb.setChecked(settings.minimize_to_tray)
        layout.addRow("", self.tray_cb)

        self.autostart_cb = QCheckBox("Запускать с Windows (в трее)")
        self.autostart_cb.setChecked(settings.autostart_windows)
        if not is_frozen_exe():
            self.autostart_cb.setEnabled(False)
            self.autostart_cb.setToolTip(
                "Автозапуск работает только из собранного .exe, не из dev-режима"
            )
        layout.addRow("", self.autostart_cb)

        bg_hint = QLabel(
            "В фоновом режиме клиент продолжает следить за watch-папкой и принимать ответы\n"
            "с сервера. Чтобы выйти полностью — правый клик на иконке в трее → Выход."
        )
        bg_hint.setStyleSheet("color: #6B7280; font-size: 11px;")
        bg_hint.setWordWrap(True)
        layout.addRow(bg_hint)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addRow(buttons)

    # ── Unlock flow ──
    def _try_unlock(self):
        if self._unlocked:
            return
        pwd, ok = QInputDialog.getText(
            self, "Пароль", "Введи пароль для разблокировки настроек сервера:",
            QLineEdit.EchoMode.Password,
        )
        if not ok:
            return
        if pwd != UNLOCK_PASSWORD:
            QMessageBox.warning(self, "Неверный пароль",
                                "Неправильный пароль. Доступ закрыт.")
            return
        # Unlocked
        self._unlocked = True
        self.server_edit.setText(self._settings.server_url)
        self.server_edit.setReadOnly(False)
        self.key_edit.setText(self._settings.api_key)
        self.key_edit.setReadOnly(False)
        self.key_edit.setEchoMode(QLineEdit.EchoMode.Password)
        self.server_copy_btn.setEnabled(True)
        self.key_copy_btn.setEnabled(True)
        self.unlock_btn.setText("🔓  Настройки сервера разблокированы")
        self.unlock_btn.setEnabled(False)

    def _copy_server(self):
        if not self._unlocked: return
        QApplication.clipboard().setText(self.server_edit.text())
        self.server_copy_btn.setText("✓")
        QTimer.singleShot(1200, lambda: self.server_copy_btn.setText("📋"))

    def _copy_key(self):
        if not self._unlocked: return
        QApplication.clipboard().setText(self.key_edit.text())
        self.key_copy_btn.setText("✓")
        QTimer.singleShot(1200, lambda: self.key_copy_btn.setText("📋"))

    def _pick_watch(self):
        d = QFileDialog.getExistingDirectory(self, "Выбери watch-папку", self.watch_edit.text())
        if d:
            self.watch_edit.setText(d)

    def _pick_output(self):
        d = QFileDialog.getExistingDirectory(self, "Выбери папку для docx", self.output_edit.text())
        if d:
            self.output_edit.setText(d)

    def get_settings(self, prev: Settings) -> Settings:
        # If locked, keep server URL and API key from previous settings
        if self._unlocked:
            server_url = self.server_edit.text().strip().rstrip("/")
            api_key = self.key_edit.text().strip()
        else:
            server_url = prev.server_url
            api_key = prev.api_key
        # Нормализуем username — с ведущим @
        tg_un = self.tg_username_edit.text().strip()
        if tg_un and not tg_un.startswith("@"):
            tg_un = "@" + tg_un
        return Settings(
            server_url=server_url,
            api_key=api_key,
            watch_folder=self.watch_edit.text().strip(),
            output_folder=self.output_edit.text().strip(),
            auto_mode=prev.auto_mode,
            minimize_to_tray=self.tray_cb.isChecked(),
            autostart_windows=self.autostart_cb.isChecked(),
            tg_username=tg_un,
        )


# ────────────────────── Job + Queue ──────────────────────

@dataclass
class Job:
    file_path: Optional[Path]   # None for restored jobs
    source: str                 # "manual" | "watch" | "restored"
    state: str = "preparing"
    job_id: Optional[str] = None
    queue_ahead: int = 0
    progress: int = 0
    error: Optional[str] = None
    docx_saved_to: Optional[str] = None
    transcript_path: Optional[str] = None   # путь к скачанной сырой расшифровке (companion)
    cancelled: bool = False
    added_at: float = field(default_factory=lambda: time.time())

    @property
    def display_name(self) -> str:
        if self.file_path:
            return self.file_path.name
        return f"восстановлено: {self.job_id}"


# ────────────────────── Animated dots ──────────────────────

class PulsingDots(QWidget):
    def __init__(self, parent=None, color="#1B7C70"):
        super().__init__(parent)
        self._color = QColor(color)
        self._phase = 0.0
        self.setFixedSize(60, 16)
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._tick)

    def start(self):
        self._timer.start(50)
        self.show()

    def stop(self):
        self._timer.stop()
        self.hide()

    def _tick(self):
        self._phase = (self._phase + 0.06) % (2.0 * math.pi)
        self.update()

    def paintEvent(self, _e):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        radius = 5
        gap = 14
        for i in range(3):
            phase_offset = i * 0.6
            alpha = 0.3 + 0.7 * (0.5 + 0.5 * math.sin(self._phase - phase_offset))
            c = QColor(self._color); c.setAlphaF(alpha)
            p.setBrush(c); p.setPen(Qt.PenStyle.NoPen)
            p.drawEllipse(QPoint(8 + i * gap, 8), radius, radius)


# ────────────────────── ProcessedTracker ──────────────────────

class ProcessedTracker:
    """Persistent dedup index next to the watch folder."""
    def __init__(self, watch_folder: Path):
        self.path = watch_folder / ".processed.json"
        self.data: Dict[str, Dict] = {}
        self._load()

    def _load(self):
        if self.path.exists():
            try:
                self.data = json.loads(self.path.read_text(encoding="utf-8"))
            except Exception:
                self.data = {}

    def _save(self):
        try:
            self.path.write_text(json.dumps(self.data, ensure_ascii=False, indent=2),
                                 encoding="utf-8")
        except Exception:
            pass

    def is_done(self, filename: str) -> bool:
        return self.data.get(filename, {}).get("status") == "success"

    def is_failed(self, filename: str) -> bool:
        return self.data.get(filename, {}).get("status") == "failed"

    def mark_success(self, filename: str, docx: str):
        self.data[filename] = {
            "status": "success", "docx": docx,
            "at": datetime.now(timezone.utc).isoformat(),
        }
        self._save()

    def mark_failed(self, filename: str, error: str):
        self.data[filename] = {
            "status": "failed", "error": error,
            "at": datetime.now(timezone.utc).isoformat(),
        }
        self._save()


# ────────────────────── JobRunner (thread) ──────────────────────

class JobRunner(QThread):
    """Sequential worker. Processes jobs one-by-one from a shared queue."""
    job_progress = pyqtSignal(int, int)      # (job_index, percent)
    job_state    = pyqtSignal(int, str, int) # (job_index, state, queue_ahead)
    job_done     = pyqtSignal(int, str)      # (job_index, docx_path)
    job_failed   = pyqtSignal(int, str)      # (job_index, error)

    def __init__(self, settings: Settings):
        super().__init__()
        self.settings = settings
        self.queue: List[Job] = []
        self._stop = False
        self._current_idx: Optional[int] = None

    def add(self, job: Job) -> int:
        self.queue.append(job)
        return len(self.queue) - 1

    def stop(self):
        self._stop = True

    def cancel_job(self, idx: int):
        """Mark a job as cancelled. Worker checks the flag at safe points."""
        if 0 <= idx < len(self.queue):
            self.queue[idx].cancelled = True

    def _check_cancel(self, idx: int, job: Job) -> bool:
        if job.cancelled:
            self.job_state.emit(idx, "cancelled", 0)
            # notify server, best-effort
            if job.job_id:
                try:
                    requests.delete(
                        f"{self.settings.server_url}/jobs/{job.job_id}",
                        headers={"Authorization": f"Bearer {self.settings.api_key}"},
                        timeout=10,
                    )
                except Exception:
                    pass
            return True
        return False

    def run(self):
        while not self._stop:
            # find next unprocessed
            idx = None
            for i, j in enumerate(self.queue):
                if j.state == "preparing":
                    idx = i; break
            if idx is None:
                self.msleep(500)
                continue
            self._current_idx = idx
            try:
                if self.queue[idx].source == "restored":
                    self._process_restored(idx, self.queue[idx])
                else:
                    self._process(idx, self.queue[idx])
            except Exception as e:
                self.job_failed.emit(idx, f"{type(e).__name__}: {e}")
            self._current_idx = None

    def _process(self, idx: int, job: Job):
        # 1. Upload
        if self._check_cancel(idx, job): return
        self.job_state.emit(idx, "uploading", 0)
        # Файл мог оказаться заблокированным между сканом и выгрузкой (дописывается
        # рекордером / OneDrive). Пробуем открыть с несколькими ретраями, чтобы не
        # падать с PermissionError, а дождаться разблокировки.
        fh = None
        for _att in range(6):
            if self._check_cancel(idx, job): return
            try:
                fh = open(job.file_path, "rb"); break
            except PermissionError:
                time.sleep(2)
            except OSError as e:
                self.job_failed.emit(idx, f"Не удалось открыть файл: {e}"); return
        if fh is None:
            self.job_failed.emit(idx, "Файл занят другим приложением (ещё копируется или открыт). Закройте его — обработаю автоматически."); return
        try:
            files = {"file": (job.file_path.name, fh)}
            # Опционально — Telegram username для дублирования статусов в TG
            data = {}
            if (self.settings.tg_username or "").strip():
                data["tg_username"] = self.settings.tg_username.strip()
            resp = requests.post(
                f"{self.settings.server_url}/upload",
                headers={"Authorization": f"Bearer {self.settings.api_key}"},
                files=files,
                data=data if data else None,
                timeout=(10, 600),
            )
        except requests.exceptions.ConnectionError:
            self.job_failed.emit(idx, "Сервер недоступен"); return
        except requests.exceptions.Timeout:
            self.job_failed.emit(idx, "Таймаут загрузки"); return
        finally:
            fh.close()

        if self._check_cancel(idx, job): return
        if resp.status_code != 200:
            detail = resp.text[:300]
            try:
                detail = resp.json().get("detail", detail)
            except Exception:
                pass
            self.job_failed.emit(idx, f"HTTP {resp.status_code}: {detail}"); return

        job.job_id = resp.json().get("job_id")
        if not job.job_id:
            self.job_failed.emit(idx, "Сервер не вернул job_id"); return

        # 2 + 3. Poll then download
        self._poll_then_download(idx, job)

    def _process_restored(self, idx: int, job: Job):
        """Restored job — already uploaded, we just need to poll + download."""
        if not job.job_id:
            self.job_failed.emit(idx, "Восстановленная задача без job_id"); return
        if self._check_cancel(idx, job): return
        self.job_state.emit(idx, "restoring", 0)
        self._poll_then_download(idx, job)

    def _poll_then_download(self, idx: int, job: Job):
        self.job_state.emit(idx, "queued", 0)
        while not self._stop:
            if self._check_cancel(idx, job): return
            self.msleep(POLL_STATUS_MS)
            if self._check_cancel(idx, job): return
            try:
                r = requests.get(
                    f"{self.settings.server_url}/status/{job.job_id}",
                    headers={"Authorization": f"Bearer {self.settings.api_key}"},
                    timeout=10,
                )
                if r.status_code == 404:
                    self.job_failed.emit(idx, "Сервер не помнит этой задачи (очищена?)")
                    return
                if r.status_code != 200:
                    continue
                data = r.json()
                state = data.get("state", "queued")
                qa = data.get("queue_ahead", 0)
                self.job_state.emit(idx, state, qa)
                if state == "done":
                    break
                if state == "failed":
                    reason = data.get("error_reason") or "Сервер не смог обработать файл"
                    self.job_failed.emit(idx, reason)
                    return
                if state == "cancelled":
                    job.cancelled = True
                    return
            except requests.exceptions.RequestException:
                continue
        if self._stop or self._check_cancel(idx, job):
            return

        # Download — две части: сводка (summary) + сырая расшифровка (transcript)
        self.job_state.emit(idx, "downloading", 0)

        def _fetch(part: str):
            """Скачивает часть результата. Возвращает (bytes, filename) или (None, None)."""
            try:
                params = None if part == "summary" else {"part": part}
                rr = requests.get(
                    f"{self.settings.server_url}/result/{job.job_id}",
                    headers={"Authorization": f"Bearer {self.settings.api_key}"},
                    params=params, timeout=120, stream=True,
                )
            except requests.exceptions.RequestException:
                return None, None
            if rr.status_code != 200:
                return None, None
            fn = None
            cd = rr.headers.get("Content-Disposition", "")
            if "filename=" in cd:
                fn = cd.split("filename=", 1)[1].strip().strip('"').strip("'") or None
            buf = bytearray()
            for chunk in rr.iter_content(chunk_size=64 * 1024):
                if chunk:
                    buf += chunk
            return bytes(buf), fn

        summary_bytes, summary_fn = _fetch("summary")
        if summary_bytes is None:
            self.job_failed.emit(idx, "Не удалось скачать сводку")
            return
        # Расшифровка — companion-файл; может ещё не появиться, тогда просто пропускаем
        transcript_bytes, _ = _fetch("transcript")

        suggest = summary_fn or "transcription.docx"

        if job.source == "watch":
            if not self.settings.output_folder:
                self.job_failed.emit(idx, "Не задана output-папка в Настройках")
                return
            out_dir = Path(self.settings.output_folder)
            out_dir.mkdir(parents=True, exist_ok=True)
            # Имя сводки = имя исходного аудио (с .docx)
            stem = job.file_path.stem if job.file_path else Path(suggest).stem
            save_path = out_dir / f"{stem}.docx"
            save_path.write_bytes(summary_bytes)
            if transcript_bytes:
                tpath = out_dir / f"{stem}_расшифровка.docx"
                tpath.write_bytes(transcript_bytes)
                job.transcript_path = str(tpath)
            job.docx_saved_to = str(save_path)
            self.job_done.emit(idx, str(save_path))
        else:
            # manual / restored — сохраняем во временную папку, UI спросит куда класть
            tmp_dir = Path.home() / ".transcriber_tmp"
            tmp_dir.mkdir(parents=True, exist_ok=True)
            tmp = tmp_dir / f"{job.job_id}_{suggest}"
            tmp.write_bytes(summary_bytes)
            if transcript_bytes:
                ttmp = tmp_dir / f"{job.job_id}_расшифровка.docx"
                ttmp.write_bytes(transcript_bytes)
                job.transcript_path = str(ttmp)
            job.docx_saved_to = str(tmp)
            self.job_done.emit(idx, str(tmp))


# ────────────────────── Folder watcher ──────────────────────

class FolderWatcher(QObject):
    new_file = pyqtSignal(Path)  # detected new stable file

    def __init__(self, parent=None):
        super().__init__(parent)
        self.timer = QTimer(self)
        self.timer.timeout.connect(self._scan)
        self.tracker: Optional[ProcessedTracker] = None
        self.folder: Optional[Path] = None
        # transient size memory: filename → (size, seen_at)
        self._size_memory: Dict[str, tuple] = {}
        # set of filenames already queued in this session
        self._queued: set = set()

    def start(self, folder: Path, tracker: ProcessedTracker):
        self.folder = folder
        self.tracker = tracker
        self._size_memory.clear()
        self.timer.start(WATCH_INTERVAL_MS)
        # also do an immediate scan
        QTimer.singleShot(500, self._scan)

    def stop(self):
        self.timer.stop()
        self.folder = None
        self._size_memory.clear()

    def mark_queued(self, name: str):
        self._queued.add(name)

    def _scan(self):
        if not self.folder or not self.folder.exists() or not self.tracker:
            return
        now = time.time()
        for f in self.folder.iterdir():
            if not f.is_file():
                continue
            if f.suffix.lower() not in ALLOWED_EXTS:
                continue
            name = f.name
            if self.tracker.is_done(name) or self.tracker.is_failed(name):
                continue
            if name in self._queued:
                continue
            try:
                size = f.stat().st_size
            except OSError:
                continue
            prev = self._size_memory.get(name)
            self._size_memory[name] = (size, now)
            if prev is None:
                continue
            prev_size, prev_seen = prev
            # require: size unchanged AND last seen ≥ STABLE_CHECK_DELAY_S ago
            if size == prev_size and (now - prev_seen) >= STABLE_CHECK_DELAY_S:
                # Файл может быть размер-стабилен, но всё ещё ЗАБЛОКИРОВАН: его
                # дописывает рекордер, либо это «онлайн-только» файл OneDrive,
                # ещё не скачанный локально. В таком случае open() даст
                # PermissionError при выгрузке. Поэтому проверяем, что файл реально
                # открывается на чтение; если нет — пропускаем до следующего цикла.
                try:
                    with open(f, "rb") as _chk:
                        _chk.read(1)
                except (PermissionError, OSError):
                    continue
                self.new_file.emit(f)


# ────────────────────── Drop zone (manual) ──────────────────────

class DropZone(QFrame):
    file_dropped = pyqtSignal(Path)
    clicked = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("dropZone")
        self.setProperty("hover", False)
        self.setProperty("fileLoaded", False)
        self.setMinimumHeight(130)
        self.setAcceptDrops(True)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(4)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.icon_label = QLabel("📁")
        self.icon_label.setStyleSheet("font-size: 30px;")
        self.icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.title_label = QLabel("Перетащи аудио сюда или нажми, чтобы выбрать")
        self.title_label.setObjectName("dropLabel")
        self.title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.subtitle_label = QLabel("mp3, wav, m4a, ogg, opus · до 500 МБ")
        self.subtitle_label.setStyleSheet("color: #9CA3AF; font-size: 11px;")
        self.subtitle_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        layout.addWidget(self.icon_label)
        layout.addWidget(self.title_label)
        layout.addWidget(self.subtitle_label)

    def mousePressEvent(self, e):
        if e.button() == Qt.MouseButton.LeftButton:
            self.clicked.emit()

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            self._set_hover(True); e.acceptProposedAction()

    def dragLeaveEvent(self, _e):
        self._set_hover(False)

    def dropEvent(self, e):
        self._set_hover(False)
        urls = e.mimeData().urls()
        if urls:
            p = Path(urls[0].toLocalFile())
            if p.is_file():
                self.file_dropped.emit(p)

    def _set_hover(self, on: bool):
        self.setProperty("hover", on)
        self.style().unpolish(self); self.style().polish(self)


# ────────────────────── Status banner ──────────────────────

class StatusBanner(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0); layout.setSpacing(10)
        self.label = QLabel("")
        self.label.setObjectName("statusBubble")
        self.label.setProperty("kind", "info")
        self.label.setWordWrap(True)
        self.dots = PulsingDots(); self.dots.hide()
        layout.addWidget(self.label, 1)
        layout.addWidget(self.dots, 0)
        self.opacity = QGraphicsOpacityEffect(self.label)
        self.opacity.setOpacity(1.0)
        self.label.setGraphicsEffect(self.opacity)
        self._anim = QPropertyAnimation(self.opacity, b"opacity")
        self._anim.setDuration(180)
        self._anim.setEasingCurve(QEasingCurve.Type.InOutQuad)
        self._pending_text = ""
        self._pending_kind = "info"
        self.hide()

    def show_message(self, text: str, kind: str = "info", show_dots: bool = False):
        self._anim.stop()
        self._anim.setStartValue(self.opacity.opacity())
        self._anim.setEndValue(0.0)
        try:
            self._anim.finished.disconnect(self._swap)
        except TypeError:
            pass
        self._anim.finished.connect(self._swap)
        self._pending_text = text
        self._pending_kind = kind
        self._anim.start()
        if show_dots:
            self.dots.start()
        else:
            self.dots.stop()
        self.show()

    def _swap(self):
        try:
            self._anim.finished.disconnect(self._swap)
        except TypeError:
            pass
        self.label.setText(self._pending_text)
        self.label.setProperty("kind", self._pending_kind)
        self.label.style().unpolish(self.label)
        self.label.style().polish(self.label)
        self._anim.setStartValue(0.0); self._anim.setEndValue(1.0)
        self._anim.start()


# ────────────────────── Auto-mode panel ──────────────────────

class AutoPanel(QFrame):
    toggle_clicked = pyqtSignal()
    open_settings  = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("autoPanel")
        self.setProperty("active", False)
        layout = QVBoxLayout(self); layout.setContentsMargins(14, 12, 14, 12); layout.setSpacing(6)

        top = QHBoxLayout(); top.setSpacing(8)
        self.title = QLabel("🤖  Авто-режим: <b>выключен</b>")
        self.title.setStyleSheet("font-size: 14px;")
        self.toggle_btn = QPushButton("Включить")
        self.toggle_btn.clicked.connect(self.toggle_clicked.emit)
        top.addWidget(self.title, 1); top.addWidget(self.toggle_btn, 0)
        layout.addLayout(top)

        self.folder_label = QLabel("Watch-папка не задана")
        self.folder_label.setStyleSheet("color: #6B7280; font-size: 12px;")
        layout.addWidget(self.folder_label)

        bottom = QHBoxLayout()
        self.counters = QLabel("обработано: 0   •   ошибок: 0   •   в очереди: 0")
        self.counters.setObjectName("counter")
        bottom.addWidget(self.counters, 1)
        cfg = QToolButton(); cfg.setText("⚙ Настроить папки")
        cfg.setAutoRaise(True); cfg.setStyleSheet("font-size: 12px; color: #1B7C70;")
        cfg.clicked.connect(self.open_settings.emit)
        bottom.addWidget(cfg, 0)
        layout.addLayout(bottom)

    def set_state(self, active: bool, watch_folder: str, counters: str):
        self.setProperty("active", active)
        self.style().unpolish(self); self.style().polish(self)
        self.title.setText(
            f"🤖  Авто-режим: <b style='color:#1B7C70'>включён</b>" if active
            else "🤖  Авто-режим: <b>выключен</b>"
        )
        self.toggle_btn.setText("Выключить" if active else "Включить")
        self.toggle_btn.setObjectName("toggleOn" if active else "")
        self.toggle_btn.style().unpolish(self.toggle_btn)
        self.toggle_btn.style().polish(self.toggle_btn)
        self.folder_label.setText(
            f"📂 {watch_folder}" if watch_folder else "Watch-папка не задана (нажми ⚙)"
        )
        self.counters.setText(counters)


# ────────────────────── Main window ──────────────────────

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME + "  ·  сводка + расшифровка")
        self.resize(680, 720); self.setMinimumSize(620, 560)

        self.settings = Settings.load()
        # If auto-mode was enabled at last quit but watch_folder is gone — disable
        if self.settings.auto_mode and not self.settings.watch_folder:
            self.settings.auto_mode = False; self.settings.save()

        self.runner = JobRunner(self.settings)
        self.runner.job_state.connect(self._on_state)
        self.runner.job_done.connect(self._on_done)
        self.runner.job_failed.connect(self._on_failed)
        self.runner.start()

        self.watcher = FolderWatcher(self)
        self.watcher.new_file.connect(self._on_watch_new_file)
        self.tracker: Optional[ProcessedTracker] = None

        # counters
        self.counter_done = 0
        self.counter_failed = 0
        self.counter_queue = 0

        # current manual job idx (the last one waiting for save dialog)
        self._manual_jobs_pending_save: set = set()

        root = QWidget(); root.setObjectName("root"); self.setCentralWidget(root)
        outer = QVBoxLayout(root); outer.setContentsMargins(24, 20, 24, 16); outer.setSpacing(14)

        header = QLabel("🎤 Транскрибер консультаций"); header.setObjectName("header")
        sub = QLabel("Аудио → распознавание → языковая модель → Word"); sub.setObjectName("subheader")
        outer.addWidget(header); outer.addWidget(sub); outer.addSpacing(4)

        # Manual drop
        outer.addWidget(self._mk_section_label("Ручная отправка"))
        self.drop_zone = DropZone()
        self.drop_zone.clicked.connect(self._pick_file)
        self.drop_zone.file_dropped.connect(lambda p: self._add_manual(p))
        outer.addWidget(self.drop_zone)

        # Auto-mode
        outer.addSpacing(4)
        outer.addWidget(self._mk_section_label("Авто-режим"))
        self.auto_panel = AutoPanel()
        self.auto_panel.toggle_clicked.connect(self._toggle_auto)
        self.auto_panel.open_settings.connect(self._open_settings)
        outer.addWidget(self.auto_panel)

        # Activity list + cancel button
        outer.addSpacing(4)
        outer.addWidget(self._mk_section_label("Активность"))
        self.activity = QListWidget()
        self.activity.setMaximumHeight(180)
        self.activity.itemSelectionChanged.connect(self._update_cancel_btn)
        outer.addWidget(self.activity, 1)

        cancel_row = QHBoxLayout()
        cancel_row.addStretch()
        self.cancel_btn = QPushButton("✖  Отменить выбранную задачу")
        self.cancel_btn.setEnabled(False)
        self.cancel_btn.clicked.connect(self._cancel_selected)
        cancel_row.addWidget(self.cancel_btn, 0)
        outer.addLayout(cancel_row)

        # Status banner (transient messages)
        self.status_banner = StatusBanner()
        outer.addWidget(self.status_banner)

        # Status bar
        self.setStatusBar(QStatusBar())
        self.conn_label = QLabel("⚙  подключение…")
        self.conn_label.setStyleSheet("padding: 0 8px; font-size: 12px;")
        self.statusBar().addWidget(self.conn_label, 1)
        credit = QLabel("Разработчик: Салимов Булат Р.")
        credit.setStyleSheet("color: #9CA3AF; font-size: 11px; padding-right: 6px;")
        self.statusBar().addPermanentWidget(credit)
        # Periodic re-ping so the indicator stays accurate if server goes down
        self._ping_timer = QTimer(self)
        self._ping_timer.timeout.connect(self._ping_async)
        self._ping_timer.start(30_000)

        # Menu
        menu = self.menuBar(); file_menu = menu.addMenu("Файл")
        a_settings = QAction("Настройки…", self); a_settings.triggered.connect(self._open_settings)
        file_menu.addAction(a_settings)
        a_quit = QAction("Выход", self); a_quit.triggered.connect(self.close)
        file_menu.addAction(a_quit)

        # restore auto-mode if was on at last quit
        self._update_auto_panel()
        if self.settings.auto_mode and self.settings.watch_folder:
            self._start_watch()

        QTimer.singleShot(300, self._ping_async)
        # offer to resume unfinished jobs from last session
        QTimer.singleShot(600, self._restore_pending)

        # System tray (для фонового режима)
        self._setup_tray()

    # ── UI helpers ──
    def _mk_section_label(self, text: str) -> QLabel:
        lbl = QLabel(text); lbl.setObjectName("sectionTitle")
        return lbl

    def _refresh_statusbar(self):
        self.conn_label.setText(
            f"<span style='color:#6B7280;'>⚙ проверяю</span> "
            f"<span style='color:#374151;'>{self.settings.server_url}</span>"
        )

    def _ping_async(self):
        url = self.settings.server_url
        try:
            r = requests.get(f"{url}/ping", timeout=5)
            if r.status_code == 200:
                self.conn_label.setText(
                    f"<span style='color:#1B7C70;'><b>✓ Подключено:</b></span> "
                    f"<span style='color:#065F46;'>{url}</span>"
                )
                return
        except Exception:
            pass
        self.conn_label.setText(
            f"<span style='color:#DC2626;'><b>⚠ Сервер недоступен:</b></span> "
            f"<span style='color:#991B1B;'>{url}</span>"
        )

    def _open_settings(self):
        dlg = SettingsDialog(self, self.settings)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            new = dlg.get_settings(self.settings)
            old_watch, old_auto = self.settings.watch_folder, self.settings.auto_mode
            old_autostart = self.settings.autostart_windows
            self.settings = new
            self.settings.save()
            self.runner.settings = new
            self._refresh_statusbar(); self._ping_async()
            self._update_auto_panel()
            if old_auto and (old_watch != new.watch_folder):
                self.watcher.stop()
                if new.watch_folder:
                    self._start_watch()
            # Apply Windows autostart change (only on actual change)
            if new.autostart_windows != old_autostart:
                ok, msg = set_windows_autostart(new.autostart_windows)
                if not ok:
                    # revert checkbox state in settings — autostart was rejected
                    self.settings.autostart_windows = old_autostart
                    self.settings.save()
                    QMessageBox.warning(self, "Автозапуск", msg)
                else:
                    self.status_banner.show_message(f"⚙ {msg}", "info")

    def _pick_file(self):
        last_dir = QSettings(ORG_NAME, APP_NAME).value("last_open_dir", "", type=str)
        path_str, _ = QFileDialog.getOpenFileName(
            self, "Выбери аудиофайл", last_dir, ALLOWED_FILTERS
        )
        if path_str:
            self._add_manual(Path(path_str))

    # ── Job queueing ──
    def _add_manual(self, path: Path):
        if path.suffix.lower() not in ALLOWED_EXTS:
            QMessageBox.warning(self, "Неподдерживаемый формат",
                f"Расширение {path.suffix} не поддерживается.")
            return
        try:
            size_mb = path.stat().st_size / (1024 * 1024)
        except OSError:
            return
        if size_mb > 500:
            QMessageBox.warning(self, "Файл слишком большой",
                f"Размер {size_mb:.1f} МБ > лимит 500 МБ."); return
        job = Job(file_path=path, source="manual")
        idx = self.runner.add(job)
        self._manual_jobs_pending_save.add(idx)
        QSettings(ORG_NAME, APP_NAME).setValue("last_open_dir", str(path.parent))
        self._add_activity(idx, job)
        self.counter_queue += 1
        self._refresh_counters()

    def _on_watch_new_file(self, path: Path):
        job = Job(file_path=path, source="watch")
        idx = self.runner.add(job)
        self.watcher.mark_queued(path.name)
        self._add_activity(idx, job)
        self.counter_queue += 1
        self._refresh_counters()

    # ── Activity list ──
    def _add_activity(self, idx: int, job: Job):
        item = QListWidgetItem()
        emoji = {"manual": "🎯", "watch": "📁", "restored": "↩"}.get(job.source, "•")
        item.setText(f"{emoji}  {job.display_name}   ·   ожидание…")
        item.setData(Qt.ItemDataRole.UserRole, idx)
        self.activity.insertItem(0, item)
        while self.activity.count() > 30:
            self.activity.takeItem(self.activity.count() - 1)

    def _update_activity_row(self, idx: int, text: str):
        for i in range(self.activity.count()):
            it = self.activity.item(i)
            if it.data(Qt.ItemDataRole.UserRole) == idx:
                it.setText(text); return

    def _activity_text(self, job: Job, state: str, queue_ahead: int) -> str:
        emoji_src = {"manual": "🎯", "watch": "📁", "restored": "↩"}.get(job.source, "•")
        emoji_state = {
            "preparing": "⏳", "uploading": "📤", "queued": "🕐",
            "stt": "📝", "llm": "🤖", "downloading": "⬇",
            "done": "✅", "failed": "❌", "cancelled": "⛔",
            "restoring": "🔄",
        }.get(state, "•")
        label = STATE_LABELS.get(state, state)
        if state == "queued" and queue_ahead > 0:
            label = f"В очереди (впереди {queue_ahead})"
        return f"{emoji_src}  {job.display_name}   ·   {emoji_state} {label}"

    # ── Runner callbacks ──
    def _on_state(self, idx: int, state: str, queue_ahead: int):
        if idx >= len(self.runner.queue): return
        job = self.runner.queue[idx]
        if state == "cancelled":
            self._on_cancelled(idx)
            return
        job.state = state; job.queue_ahead = queue_ahead
        self._update_activity_row(idx, self._activity_text(job, state, queue_ahead))
        self._update_cancel_btn()

    def _on_done(self, idx: int, docx_path: str):
        if idx >= len(self.runner.queue): return
        job = self.runner.queue[idx]
        job.state = "done"; job.docx_saved_to = docx_path
        self.counter_done += 1
        self.counter_queue = max(0, self.counter_queue - 1)
        self._refresh_counters()
        self._update_cancel_btn()

        if job.source == "watch":
            try:
                if self.tracker and job.file_path:
                    self.tracker.mark_success(job.file_path.name, docx_path)
                if job.file_path:
                    job.file_path.unlink(missing_ok=True)
            except Exception:
                pass
            saved_msg = Path(docx_path).name
            if getattr(job, "transcript_path", None):
                saved_msg = f"{Path(docx_path).name} + {Path(job.transcript_path).name}"
            self._update_activity_row(
                idx,
                f"📁  {job.display_name}   ·   ✅ сохранено: {saved_msg}"
            )
        else:
            # manual or restored: prompt user for save location
            emoji = "🎯" if job.source == "manual" else "↩"
            self._update_activity_row(
                idx,
                f"{emoji}  {job.display_name}   ·   ✅ получен docx, сохранение…"
            )
            self._manual_save(idx, docx_path)

    def _manual_save(self, idx: int, tmp_docx: str):
        job = self.runner.queue[idx]
        # Suggest filename = audio stem + .docx (if we have audio path)
        if job.file_path:
            suggest = job.file_path.stem + ".docx"
        else:
            # restored — no audio path, fall back to server-suggested
            suggest = Path(tmp_docx).name
            if "_" in suggest:
                suggest = suggest.split("_", 1)[1]
        last_save = QSettings(ORG_NAME, APP_NAME).value("last_save_dir", "", type=str)
        default_path = str(Path(last_save) / suggest) if last_save else suggest
        save_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить расшифровку", default_path,
            "Word документ (*.docx)"
        )
        if not save_path:
            self.status_banner.show_message(
                f"⚠ Сохранение отменено. Файл лежит во временной папке: {tmp_docx}", "warn"
            )
            return
        try:
            data = Path(tmp_docx).read_bytes()
            Path(save_path).write_bytes(data)
            # Companion: сырая расшифровка рядом со сводкой ({имя}_расшифровка.docx)
            saved_msg = Path(save_path).name
            tp = getattr(job, "transcript_path", None)
            if tp and Path(tp).exists():
                tdest = Path(save_path).with_name(Path(save_path).stem + "_расшифровка.docx")
                try:
                    tdest.write_bytes(Path(tp).read_bytes())
                    Path(tp).unlink(missing_ok=True)
                    saved_msg = f"{Path(save_path).name} + {tdest.name}"
                except Exception:
                    pass
            QSettings(ORG_NAME, APP_NAME).setValue("last_save_dir", str(Path(save_path).parent))
            emoji = "🎯" if job.source == "manual" else "↩"
            self._update_activity_row(
                idx,
                f"{emoji}  {job.display_name}   ·   ✅ сохранено: {saved_msg}"
            )
            try:
                Path(tmp_docx).unlink(missing_ok=True)
            except Exception:
                pass
        except Exception as e:
            self.status_banner.show_message(f"❌ Ошибка сохранения: {e}", "error")

    def _on_failed(self, idx: int, error: str):
        if idx >= len(self.runner.queue): return
        job = self.runner.queue[idx]
        job.state = "failed"; job.error = error
        self.counter_failed += 1
        self.counter_queue = max(0, self.counter_queue - 1)
        self._refresh_counters()
        self._update_cancel_btn()
        name = job.display_name
        self._update_activity_row(idx, f"❌  {name}   ·   {error[:60]}")
        if job.source == "watch" and self.tracker and job.file_path:
            self.tracker.mark_failed(job.file_path.name, error)

    # ── Cancel ──
    def _update_cancel_btn(self):
        items = self.activity.selectedItems()
        if not items:
            self.cancel_btn.setEnabled(False); return
        idx = items[0].data(Qt.ItemDataRole.UserRole)
        if idx is None or idx >= len(self.runner.queue):
            self.cancel_btn.setEnabled(False); return
        job = self.runner.queue[idx]
        self.cancel_btn.setEnabled(job.state in CANCELLABLE_STATES and not job.cancelled)

    def _cancel_selected(self):
        items = self.activity.selectedItems()
        if not items: return
        idx = items[0].data(Qt.ItemDataRole.UserRole)
        if idx is None or idx >= len(self.runner.queue): return
        job = self.runner.queue[idx]
        if job.state in TERMINAL_STATES or job.cancelled:
            return
        # Confirm if mid-processing
        reply = QMessageBox.question(
            self, "Отменить задачу?",
            f"Отменить обработку файла «{job.display_name}»?\n\n"
            "Загруженный файл будет удалён с сервера, результат не будет сформирован.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        self.runner.cancel_job(idx)
        # Optimistically update UI; runner will emit cancelled state when it checks
        job.cancelled = True
        self._update_activity_row(idx, f"⛔  {job.display_name}   ·   Отмена…")
        self.cancel_btn.setEnabled(False)

    def _on_cancelled(self, idx: int):
        """Called when runner confirms cancellation."""
        if idx >= len(self.runner.queue): return
        job = self.runner.queue[idx]
        job.state = "cancelled"
        self.counter_queue = max(0, self.counter_queue - 1)
        self._refresh_counters()
        self._update_activity_row(idx, f"⛔  {job.display_name}   ·   Отменено")
        self._update_cancel_btn()

    # ── Auto-mode controls ──
    def _update_auto_panel(self):
        active = self.settings.auto_mode and bool(self.settings.watch_folder)
        self.auto_panel.set_state(
            active=active,
            watch_folder=self.settings.watch_folder,
            counters=f"обработано: {self.counter_done}   •   ошибок: {self.counter_failed}   •   в очереди: {self.counter_queue}",
        )

    def _refresh_counters(self):
        self._update_auto_panel()

    def _toggle_auto(self):
        if self.settings.auto_mode:
            self.settings.auto_mode = False; self.settings.save()
            self.watcher.stop()
            self.tracker = None
            self._update_auto_panel()
            self.status_banner.show_message("🤖 Авто-режим выключен", "info")
            return
        if not self.settings.watch_folder:
            QMessageBox.information(
                self, "Папка не задана",
                "Сначала укажи watch-папку в Настройках (меню «Файл → Настройки»)."
            )
            return
        if not self.settings.output_folder:
            QMessageBox.information(
                self, "Output-папка не задана",
                "Укажи папку для сохранения docx в Настройках."
            )
            return
        self.settings.auto_mode = True; self.settings.save()
        self._start_watch()
        self._update_auto_panel()
        self.status_banner.show_message(
            f"🤖 Авто-режим включён. Слежу за: {self.settings.watch_folder}", "success"
        )

    def _start_watch(self):
        folder = Path(self.settings.watch_folder)
        if not folder.exists():
            try:
                folder.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                QMessageBox.warning(self, "Watch-папка",
                    f"Не удалось создать папку {folder}: {e}")
                self.settings.auto_mode = False; self.settings.save()
                self._update_auto_panel(); return
        self.tracker = ProcessedTracker(folder)
        self.watcher.start(folder, self.tracker)

    # ── Session save/restore ──
    def _save_pending(self):
        """On close: snapshot jobs that may still have a server-side result.
        Сохраняем всё кроме отменённых и успешно скачанных — даже failed,
        потому что на сервере docx мог родиться после фиксов парсера/промта."""
        pending = []
        for job in self.runner.queue:
            if job.cancelled:
                continue
            if not job.job_id:
                continue
            # Уже успешно сохранён локально → нечего восстанавливать
            if job.state == "done" and job.docx_saved_to:
                continue
            pending.append({
                "job_id": job.job_id,
                "original_name": job.file_path.name if job.file_path else "",
                "source": job.source if job.source in ("manual", "watch") else "manual",
                "added_at": job.added_at,
                "last_state": job.state,
                "last_error": job.error,
            })
        QSettings(ORG_NAME, APP_NAME).setValue("pending_jobs_json", json.dumps(pending))

    def _restore_pending(self):
        """On startup: offer to re-check all saved jobs against server.
        Клиент опросит /status для каждого — если сервер скажет done → скачает,
        если failed на сервере → покажет в Активности как failed, и т.д."""
        raw = QSettings(ORG_NAME, APP_NAME).value("pending_jobs_json", "[]", type=str)
        try:
            pending = json.loads(raw)
        except Exception:
            pending = []
        if not isinstance(pending, list) or not pending:
            return
        # clear immediately so we don't loop on a crash
        QSettings(ORG_NAME, APP_NAME).setValue("pending_jobs_json", "[]")
        # Build readable names list
        lines = []
        for p in pending[:5]:
            name = p.get("original_name", "?") or p.get("job_id", "?")
            last = p.get("last_state", "?")
            last_label = STATE_LABELS.get(last, last)
            lines.append(f"  • {name} (было: {last_label})")
        if len(pending) > 5:
            lines.append(f"  и ещё {len(pending) - 5}…")
        names = "\n".join(lines)
        reply = QMessageBox.question(
            self, "Незавершённые задачи",
            f"С прошлого раза остались {len(pending)} задач:\n\n"
            f"{names}\n\nПроверить их статус на сервере?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        for p in pending:
            job = Job(file_path=None, source="restored", job_id=p.get("job_id"))
            idx = self.runner.add(job)
            self._add_activity(idx, job)
            self.counter_queue += 1
        self._refresh_counters()

    # ── System tray (background mode) ──
    def _setup_tray(self):
        """Создаёт иконку в трее с меню."""
        if not QSystemTrayIcon.isSystemTrayAvailable():
            self.tray = None
            return
        icon = self.style().standardIcon(QStyle.StandardPixmap.SP_MediaPlay)
        self.tray = QSystemTrayIcon(icon, self)
        self.tray.setToolTip("Transcriber Client — фоновый режим")
        menu = QMenu()
        act_show = QAction("Открыть окно", self)
        act_show.triggered.connect(self._show_from_tray)
        menu.addAction(act_show)
        menu.addSeparator()
        act_quit = QAction("Выйти полностью", self)
        act_quit.triggered.connect(self._quit_app)
        menu.addAction(act_quit)
        self.tray.setContextMenu(menu)
        self.tray.activated.connect(self._tray_activated)
        self.tray.show()
        self._tray_hint_shown = False

    def _tray_activated(self, reason):
        # Двойной клик или левый клик — показать окно
        if reason in (QSystemTrayIcon.ActivationReason.Trigger,
                       QSystemTrayIcon.ActivationReason.DoubleClick):
            self._show_from_tray()

    def _show_from_tray(self):
        self.show()
        self.raise_()
        self.activateWindow()

    def _quit_app(self):
        """Принудительный выход через меню трея."""
        self._quit_requested = True
        self.close()
        QApplication.quit()

    # ── Cleanup ──
    def closeEvent(self, e):
        # Свернуть в трей при закрытии окна, если включено и трей доступен
        if (self.settings.minimize_to_tray
                and getattr(self, "tray", None) is not None
                and not getattr(self, "_quit_requested", False)):
            self.hide()
            if not getattr(self, "_tray_hint_shown", False):
                self.tray.showMessage(
                    "Transcriber Client",
                    "Свёрнут в трей. Двойной клик чтобы открыть окно.",
                    QSystemTrayIcon.MessageIcon.Information,
                    3000,
                )
                self._tray_hint_shown = True
            e.ignore()
            return
        # Полный выход
        self._save_pending()
        try:
            self.runner.stop()
            self.runner.wait(2000)
        except Exception:
            pass
        if getattr(self, "tray", None) is not None:
            self.tray.hide()
        super().closeEvent(e)


def main():
    start_minimized = "--minimized" in sys.argv
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setOrganizationName(ORG_NAME)
    app.setStyleSheet(QSS)
    app.setStyle("Fusion")
    # Не выходить когда последнее окно скрыто (мы живём в трее)
    app.setQuitOnLastWindowClosed(False)
    win = MainWindow()
    if not start_minimized:
        win.show()
    # иначе окно скрыто, бэкграунд: watch-папка работает, tray-икон в системном лотке
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
