import json
import os
import sys
import threading
from dataclasses import dataclass
from typing import Optional

from PyQt5.QtCore import Qt, QThread, QUrl, pyqtSignal
from PyQt5.QtGui import QDesktopServices, QFont, QFontDatabase
from PyQt5.QtWidgets import (
    QApplication,
    QBoxLayout,
    QFileDialog,
    QFormLayout,
    QHBoxLayout,
    QDialog,
    QPlainTextEdit,
    QSpinBox,
    QMessageBox,
    QVBoxLayout,
    QWidget,
)
from qfluentwidgets import (
    BodyLabel,
    CardWidget,
    CaptionLabel,
    ComboBox,
    FluentIcon as FIF,
    FluentWindow,
    InfoBar,
    InfoBarPosition,
    NavigationItemPosition,
    IconWidget,
    PasswordLineEdit,
    PrimaryPushButton,
    PushButton,
    SubtitleLabel,
    SwitchButton,
    Theme,
    LargeTitleLabel,
    StrongBodyLabel,
    TitleLabel,
    LineEdit,
    setTheme,
    setThemeColor,
)

from dirgc.cli import run_dirgc, validate_row_range
from dirgc.logging_utils import set_log_handler
from dirgc.settings import (
    DEFAULT_EXCEL_FILE,
    DEFAULT_IDLE_TIMEOUT_MS,
    DEFAULT_WEB_TIMEOUT_S,
)

GUI_SETTINGS_PATH = os.path.join("config", "gui_settings.json")
MAX_RECENT_EXCEL = 8
RESPONSIVE_BREAKPOINT = 980


def load_gui_settings():
    try:
        with open(GUI_SETTINGS_PATH, "r", encoding="utf-8") as handle:
            return json.load(handle)
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return {}


def save_gui_settings(data):
    os.makedirs(os.path.dirname(GUI_SETTINGS_PATH), exist_ok=True)
    with open(GUI_SETTINGS_PATH, "w", encoding="utf-8") as handle:
        json.dump(data, handle, ensure_ascii=False, indent=2)


def build_footer_label():
    footer = CaptionLabel(
        "Made with ❤️and ☕ - Novanni Indi Pradana - IPDS BPS 6502"
    )
    footer.setAlignment(Qt.AlignCenter)
    footer.setStyleSheet("color: #6B6B6B;")
    return footer


def apply_app_font(app):
    font_path = os.path.join("assets", "fonts", "Poppins-Regular.ttf")
    if os.path.exists(font_path):
        QFontDatabase.addApplicationFont(font_path)
    families = QFontDatabase().families()
    if "Poppins" in families:
        app.setFont(QFont("Poppins", 10))
    else:
        app.setFont(QFont("Segoe UI Variable", 10))


@dataclass
class RunConfig:
    headless: bool
    manual_only: bool
    excel_file: Optional[str]
    start_row: Optional[int]
    end_row: Optional[int]
    idle_timeout_ms: int
    keep_open: bool
    use_sso: bool
    sso_username: Optional[str]
    sso_password: Optional[str]
    web_timeout_s: int


class RunWorker(QThread):
    log_line = pyqtSignal(str)
    request_close = pyqtSignal()
    finished_ok = pyqtSignal()
    failed = pyqtSignal(str)
    progress = pyqtSignal(int, int, int)

    def __init__(self, config: RunConfig):
        super().__init__()
        self._config = config
        self._close_event = threading.Event()
        self._stop_event = threading.Event()

    def _handle_log(self, line, spacer=False, divider=False):
        if spacer:
            self.log_line.emit("")
        if divider:
            self.log_line.emit("-" * 72)
        self.log_line.emit(line)

    def run(self):
        set_log_handler(self._handle_log)
        try:
            credentials = None
            if self._config.use_sso:
                credentials = (
                    self._config.sso_username,
                    self._config.sso_password,
                )
            run_dirgc(
                headless=self._config.headless,
                manual_only=self._config.manual_only,
                excel_file=self._config.excel_file,
                start_row=self._config.start_row,
                end_row=self._config.end_row,
                idle_timeout_ms=self._config.idle_timeout_ms,
                web_timeout_s=self._config.web_timeout_s,
                keep_open=self._config.keep_open,
                credentials=credentials,
                stop_event=self._stop_event,
                progress_callback=self._emit_progress,
                wait_for_close=self._wait_for_close
                if self._config.keep_open
                else None,
            )
            self.finished_ok.emit()
        except Exception as exc:
            self.failed.emit(str(exc))
        finally:
            set_log_handler(None)

    def _wait_for_close(self):
        self.request_close.emit()
        self._close_event.wait()

    def release_close(self):
        self._close_event.set()

    def request_stop(self):
        self._stop_event.set()

    def _emit_progress(self, processed, total, excel_row):
        self.progress.emit(int(processed), int(total), int(excel_row))


class RunPage(QWidget):
    def __init__(self, sso_page=None, parent=None):
        super().__init__(parent)
        self._worker = None
        self._sso_page = sso_page
        self._recent_excels = []

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(16)

        title = TitleLabel("Run DIRGC")
        subtitle = BodyLabel("Atur file, opsi, lalu jalankan proses.")
        layout.addWidget(title)
        layout.addWidget(subtitle)

        content = QWidget()
        self._content_layout = QBoxLayout(QBoxLayout.LeftToRight, content)
        self._content_layout.setContentsMargins(0, 0, 0, 0)
        self._content_layout.setSpacing(16)

        left_col = QWidget()
        left_layout = QVBoxLayout(left_col)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(16)
        left_layout.addWidget(self._build_files_card())
        left_layout.addWidget(self._build_options_card())
        left_layout.addStretch()

        right_col = QWidget()
        right_layout = QVBoxLayout(right_col)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(16)
        right_layout.addWidget(self._build_run_card())
        right_layout.addWidget(self._build_log_card(), stretch=1)

        self._content_layout.addWidget(left_col, stretch=2)
        self._content_layout.addWidget(right_col, stretch=3)

        layout.addWidget(content, stretch=1)
        layout.addWidget(build_footer_label())
        self._connect_sso()
        self._is_stacked = False
        self._update_layout_mode(self.width())
        self._load_settings()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._update_layout_mode(event.size().width())

    def _update_layout_mode(self, width):
        stacked = width < RESPONSIVE_BREAKPOINT
        if stacked == self._is_stacked:
            return
        direction = QBoxLayout.TopToBottom if stacked else QBoxLayout.LeftToRight
        self._content_layout.setDirection(direction)
        self._is_stacked = stacked

    def _build_files_card(self):
        card = CardWidget()
        card_layout = QVBoxLayout(card)
        card_layout.addWidget(SubtitleLabel("Files"))

        form = QFormLayout()
        form.setHorizontalSpacing(12)
        form.setVerticalSpacing(8)
        form.setLabelAlignment(Qt.AlignLeft | Qt.AlignVCenter)

        self.excel_input = LineEdit()
        self.excel_input.setPlaceholderText("data/Direktori_SBR_20260114.xlsx")
        self.excel_input.editingFinished.connect(self._on_excel_edit_finished)
        self.excel_browse = PushButton("Browse")
        self.excel_browse.clicked.connect(
            lambda: self._browse_file(
                self.excel_input, "Excel (*.xlsx *.xls)"
            )
        )

        excel_row = QWidget()
        excel_layout = QHBoxLayout(excel_row)
        excel_layout.setContentsMargins(0, 0, 0, 0)
        excel_layout.setSpacing(8)
        excel_layout.addWidget(self.excel_input, stretch=1)
        excel_layout.addWidget(self.excel_browse)

        self.recent_combo = ComboBox()
        self.recent_combo.setPlaceholderText("File terakhir")
        self.recent_combo.currentIndexChanged.connect(
            self._on_recent_selected
        )

        form.addRow(BodyLabel("Excel file"), excel_row)
        form.addRow(BodyLabel("Recent"), self.recent_combo)

        card_layout.addLayout(form)

        self._apply_default_paths()

        return card

    def _build_options_card(self):
        card = CardWidget()
        card_layout = QVBoxLayout(card)
        card_layout.addWidget(SubtitleLabel("Options"))

        self.manual_switch = SwitchButton()
        self.manual_switch.setChecked(False)
        card_layout.addWidget(
            self._make_option_row(
                "Login manual (tanpa auto-login)",
                "ON: isi OTP langsung di browser. OFF: gunakan Akun SSO.",
                self.manual_switch,
            )
        )

        self.headless_switch = SwitchButton()
        self.headless_switch.setChecked(False)
        card_layout.addWidget(
            self._make_option_row(
                "Browser tanpa tampilan (headless)",
                "ON: browser tidak terlihat. Tidak disarankan untuk SSO/OTP.",
                self.headless_switch,
            )
        )

        self.keep_open_switch = SwitchButton()
        self.keep_open_switch.setChecked(False)
        card_layout.addWidget(
            self._make_option_row(
                "Biarkan browser tetap terbuka",
                "ON: browser tetap terbuka sampai kamu menutupnya.",
                self.keep_open_switch,
            )
        )

        idle_row = QWidget()
        idle_layout = QHBoxLayout(idle_row)
        idle_layout.setContentsMargins(0, 0, 0, 0)
        idle_layout.setSpacing(12)

        idle_label = StrongBodyLabel("Batas idle (detik)")
        idle_hint = CaptionLabel(
            "Jika tidak ada aktivitas, proses dihentikan otomatis."
        )
        idle_hint.setStyleSheet("color: #6B6B6B;")
        idle_hint.setWordWrap(True)
        self.idle_spin = QSpinBox()
        self.idle_spin.setRange(30, 3600 * 6)
        self.idle_spin.setValue(DEFAULT_IDLE_TIMEOUT_MS // 1000)
        self.idle_spin.setSuffix(" s")
        idle_text = QWidget()
        idle_text_layout = QVBoxLayout(idle_text)
        idle_text_layout.setContentsMargins(0, 0, 0, 0)
        idle_text_layout.setSpacing(2)
        idle_text_layout.addWidget(idle_label)
        idle_text_layout.addWidget(idle_hint)
        idle_layout.addWidget(idle_text, stretch=1)
        idle_layout.addStretch()
        idle_layout.addWidget(self.idle_spin)
        card_layout.addWidget(idle_row)

        web_row = QWidget()
        web_layout = QHBoxLayout(web_row)
        web_layout.setContentsMargins(0, 0, 0, 0)
        web_layout.setSpacing(12)

        web_label = StrongBodyLabel("Timeout loading web (detik)")
        web_hint = CaptionLabel(
            "Naikkan jika koneksi lambat atau halaman sering timeout."
        )
        web_hint.setStyleSheet("color: #6B6B6B;")
        web_hint.setWordWrap(True)

        web_text = QWidget()
        web_text_layout = QVBoxLayout(web_text)
        web_text_layout.setContentsMargins(0, 0, 0, 0)
        web_text_layout.setSpacing(2)
        web_text_layout.addWidget(web_label)
        web_text_layout.addWidget(web_hint)

        self.web_timeout_spin = QSpinBox()
        self.web_timeout_spin.setRange(10, 600)
        self.web_timeout_spin.setValue(DEFAULT_WEB_TIMEOUT_S)
        self.web_timeout_spin.setSuffix(" s")

        web_layout.addWidget(web_text, stretch=1)
        web_layout.addStretch()
        web_layout.addWidget(self.web_timeout_spin)
        card_layout.addWidget(web_row)

        range_row = QWidget()
        range_layout = QHBoxLayout(range_row)
        range_layout.setContentsMargins(0, 0, 0, 0)
        range_layout.setSpacing(8)

        self.range_switch = SwitchButton()
        self.range_switch.setChecked(False)
        self.range_switch.checkedChanged.connect(self._toggle_range)
        range_layout.addWidget(StrongBodyLabel("Batasi baris Excel"))
        range_layout.addStretch()
        range_layout.addWidget(self.range_switch)
        card_layout.addWidget(range_row)

        range_inputs = QWidget()
        range_inputs_layout = QHBoxLayout(range_inputs)
        range_inputs_layout.setContentsMargins(0, 0, 0, 0)
        range_inputs_layout.setSpacing(12)

        self.start_spin = QSpinBox()
        self.start_spin.setRange(1, 1000000)
        self.start_spin.setValue(1)
        self.end_spin = QSpinBox()
        self.end_spin.setRange(1, 1000000)
        self.end_spin.setValue(1)

        range_inputs_layout.addWidget(BodyLabel("Start row"))
        range_inputs_layout.addWidget(self.start_spin)
        range_inputs_layout.addSpacing(12)
        range_inputs_layout.addWidget(BodyLabel("End row"))
        range_inputs_layout.addWidget(self.end_spin)
        range_inputs_layout.addStretch()
        card_layout.addWidget(range_inputs)

        self._toggle_range()

        return card

    def _build_run_card(self):
        card = CardWidget()
        card_layout = QVBoxLayout(card)
        card_layout.addWidget(SubtitleLabel("Run"))

        button_row = QWidget()
        button_layout = QHBoxLayout(button_row)
        button_layout.setContentsMargins(0, 0, 0, 0)
        button_layout.setSpacing(8)

        self.run_button = PrimaryPushButton("Mulai")
        self.run_button.clicked.connect(self._confirm_start)
        self.stop_button = PushButton("Stop")
        self.stop_button.clicked.connect(self._confirm_stop)
        self.stop_button.setEnabled(False)
        button_layout.addWidget(self.run_button)
        button_layout.addWidget(self.stop_button)
        button_layout.addStretch()

        self.status_label = BodyLabel("Status: idle")
        self.progress_label = CaptionLabel("Progress: -")
        self.progress_label.setStyleSheet("color: #6B6B6B;")
        card_layout.addWidget(self.status_label)
        card_layout.addWidget(self.progress_label)
        card_layout.addWidget(button_row)

        return card

    def _build_log_card(self):
        card = CardWidget()
        card_layout = QVBoxLayout(card)
        header = QWidget()
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(0, 0, 0, 0)
        header_layout.setSpacing(8)
        header_layout.addWidget(SubtitleLabel("Log"))
        header_layout.addStretch()
        self.open_log_button = PushButton("Buka folder log")
        self.open_log_button.clicked.connect(self._open_log_folder)
        header_layout.addWidget(self.open_log_button)
        self.clear_log_button = PushButton("Bersihkan log")
        self.clear_log_button.clicked.connect(self._confirm_clear_log)
        header_layout.addWidget(self.clear_log_button)
        card_layout.addWidget(header)

        self.log_output = QPlainTextEdit()
        self.log_output.setReadOnly(True)
        self.log_output.setPlaceholderText("Log proses akan muncul di sini.")
        card_layout.addWidget(self.log_output)

        return card

    def _make_option_row(self, title, description, switch):
        row = QWidget()
        layout = QHBoxLayout(row)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        text_block = QWidget()
        text_layout = QVBoxLayout(text_block)
        text_layout.setContentsMargins(0, 0, 0, 0)
        text_layout.setSpacing(2)

        title_label = StrongBodyLabel(title)
        desc_label = CaptionLabel(description)
        desc_label.setWordWrap(True)
        desc_label.setStyleSheet("color: #6B6B6B;")

        text_layout.addWidget(title_label)
        text_layout.addWidget(desc_label)

        layout.addWidget(text_block, stretch=1)
        layout.addWidget(switch)
        return row

    def _apply_default_paths(self):
        excel_path = self._resolve_default_path(DEFAULT_EXCEL_FILE)
        if excel_path:
            self._set_excel_path(excel_path, push_recent=False, save=False)

    def _resolve_default_path(self, relative_path):
        candidate = os.path.join(os.getcwd(), relative_path)
        if os.path.exists(candidate):
            return candidate
        return ""

    def _browse_file(self, input_widget, file_filter):
        start_dir = os.getcwd()
        if input_widget.text():
            start_dir = os.path.dirname(input_widget.text())
        path, _ = QFileDialog.getOpenFileName(
            self, "Select file", start_dir, file_filter
        )
        if path:
            self._set_excel_path(path)

    def _set_excel_path(self, path, push_recent=True, save=True):
        self.excel_input.setText(path)
        if push_recent:
            self._push_recent_excel(path)
        if save:
            self._save_settings()

    def _on_recent_selected(self, index):
        if index < 0:
            return
        path = self.recent_combo.currentText()
        if path:
            self._set_excel_path(path, push_recent=False)

    def _on_excel_edit_finished(self):
        path = self.excel_input.text().strip()
        if path:
            self._push_recent_excel(path)
            self._save_settings()

    def _push_recent_excel(self, path):
        normalized = os.path.normpath(path)
        if not normalized:
            return
        updated = [
            item
            for item in self._recent_excels
            if os.path.normpath(item) != normalized
        ]
        updated.insert(0, normalized)
        self._recent_excels = updated[:MAX_RECENT_EXCEL]
        self._refresh_recent_combo()

    def _refresh_recent_combo(self):
        self.recent_combo.blockSignals(True)
        self.recent_combo.clear()
        if self._recent_excels:
            self.recent_combo.addItems(self._recent_excels)
            self.recent_combo.setCurrentIndex(-1)
        self.recent_combo.blockSignals(False)

    def _load_settings(self):
        data = load_gui_settings()
        options = data.get("options", {})
        excel_path = data.get("excel_path")
        recents = data.get("recent_excels", [])

        if isinstance(recents, list):
            self._recent_excels = [
                item for item in recents if isinstance(item, str)
            ]
            self._refresh_recent_combo()

        if isinstance(excel_path, str) and excel_path:
            self._set_excel_path(excel_path, push_recent=False, save=False)

        if "manual_only" in options:
            self.manual_switch.setChecked(bool(options["manual_only"]))
        if "headless" in options:
            self.headless_switch.setChecked(bool(options["headless"]))
        if "keep_open" in options:
            self.keep_open_switch.setChecked(bool(options["keep_open"]))
        if "idle_timeout_s" in options:
            self.idle_spin.setValue(int(options["idle_timeout_s"]))
        if "web_timeout_s" in options:
            self.web_timeout_spin.setValue(int(options["web_timeout_s"]))
        if "range_enabled" in options:
            self.range_switch.setChecked(bool(options["range_enabled"]))
        if "start_row" in options:
            self.start_spin.setValue(int(options["start_row"]))
        if "end_row" in options:
            self.end_spin.setValue(int(options["end_row"]))

        self._toggle_range()

    def _save_settings(self):
        data = load_gui_settings()
        data["excel_path"] = self.excel_input.text().strip()
        data["recent_excels"] = self._recent_excels
        data["options"] = {
            "manual_only": self.manual_switch.isChecked(),
            "headless": self.headless_switch.isChecked(),
            "keep_open": self.keep_open_switch.isChecked(),
            "idle_timeout_s": self.idle_spin.value(),
            "web_timeout_s": self.web_timeout_spin.value(),
            "range_enabled": self.range_switch.isChecked(),
            "start_row": self.start_spin.value(),
            "end_row": self.end_spin.value(),
        }
        save_gui_settings(data)

    def _toggle_range(self):
        enabled = self.range_switch.isChecked()
        self.start_spin.setEnabled(enabled)
        self.end_spin.setEnabled(enabled)


    def _clear_log(self):
        self.log_output.clear()

    def _confirm_clear_log(self):
        if self.log_output.toPlainText().strip() == "":
            return
        if self._confirm_dialog(
            "Bersihkan log",
            "Hapus semua log yang tampil di layar?",
        ):
            self._clear_log()

    def _open_log_folder(self):
        log_dir = os.path.join(os.getcwd(), "logs")
        os.makedirs(log_dir, exist_ok=True)
        if not QDesktopServices.openUrl(QUrl.fromLocalFile(log_dir)):
            InfoBar.error(
                title="Gagal membuka folder",
                content="Tidak bisa membuka folder logs.",
                duration=3000,
                parent=self,
                position=InfoBarPosition.TOP_RIGHT,
            )

    def _append_log(self, line):
        if line is None:
            return
        self.log_output.appendPlainText(line)
        self.log_output.verticalScrollBar().setValue(
            self.log_output.verticalScrollBar().maximum()
        )

    def _connect_sso(self):
        if not self._sso_page:
            return
        self._sso_page.use_switch.checkedChanged.connect(
            self._sync_sso_state
        )
        self._sync_sso_state()

    def _sync_sso_state(self):
        if not self._sso_page:
            return
        sso_enabled = self._sso_page.is_enabled()
        if sso_enabled:
            self.manual_switch.setChecked(False)
        self.manual_switch.setEnabled(not sso_enabled)

    def _build_config(self):
        excel_text = self.excel_input.text().strip()
        excel_file = excel_text if excel_text else None
        use_sso, sso_username, sso_password = self._get_sso_values()

        start_row = None
        end_row = None
        if self.range_switch.isChecked():
            start_row = self.start_spin.value()
            end_row = self.end_spin.value()

        return RunConfig(
            headless=self.headless_switch.isChecked(),
            manual_only=self.manual_switch.isChecked(),
            excel_file=excel_file,
            start_row=start_row,
            end_row=end_row,
            idle_timeout_ms=self.idle_spin.value() * 1000,
            web_timeout_s=self.web_timeout_spin.value(),
            keep_open=self.keep_open_switch.isChecked(),
            use_sso=use_sso,
            sso_username=sso_username,
            sso_password=sso_password,
        )

    def _validate_inputs(self, config: RunConfig):
        try:
            validate_row_range(config.start_row, config.end_row)
        except ValueError as exc:
            self._show_error(str(exc))
            return False

        if not config.excel_file:
            self._show_error("Excel file belum dipilih.")
            return False
        if not os.path.exists(config.excel_file):
            self._show_error("Excel file tidak ditemukan.")
            return False

        if config.use_sso:
            if config.manual_only:
                self._show_error(
                    "Matikan manual login untuk menggunakan Akun SSO."
                )
                return False
            if not config.sso_username or not config.sso_password:
                self._show_error("Akun SSO belum lengkap.")
                return False

        return True

    def _get_sso_values(self):
        if not self._sso_page:
            return False, None, None
        use_sso = self._sso_page.is_enabled()
        username, password = self._sso_page.get_credentials()
        return use_sso, username, password

    def _start_run(self):
        if self._worker and self._worker.isRunning():
            return

        config = self._build_config()
        if not self._validate_inputs(config):
            return

        self._save_settings()
        self.status_label.setText("Status: running")
        self._set_running_state(True)
        self.progress_label.setText("Progress: memuat data...")
        self._append_log("=== START RUN ===")

        self._worker = RunWorker(config)
        self._worker.log_line.connect(self._append_log)
        self._worker.finished_ok.connect(self._run_finished)
        self._worker.failed.connect(self._run_failed)
        self._worker.request_close.connect(self._show_keep_open_dialog)
        self._worker.progress.connect(self._update_progress)
        self._worker.start()

    def _confirm_start(self):
        if self._confirm_dialog(
            "Mulai proses",
            "Mulai proses sekarang?",
        ):
            self._start_run()

    def _confirm_stop(self):
        if self._confirm_dialog(
            "Hentikan proses",
            "Proses akan dihentikan. Lanjutkan?",
        ):
            self._stop_run()

    def _stop_run(self):
        if not self._worker or not self._worker.isRunning():
            return
        self.status_label.setText("Status: stopping")
        self._append_log("=== STOP REQUESTED ===")
        self.stop_button.setEnabled(False)
        self._worker.request_stop()

    def _run_finished(self):
        self._append_log("=== RUN FINISHED ===")
        self.status_label.setText("Status: idle")
        self._set_running_state(False)
        InfoBar.success(
            title="Run selesai",
            content="Proses selesai tanpa error.",
            duration=3000,
            parent=self,
            position=InfoBarPosition.TOP_RIGHT,
        )

    def _run_failed(self, message):
        if "Run stopped by user." in message:
            self._run_stopped()
            return
        self._append_log(f"ERROR: {message}")
        self.status_label.setText("Status: error")
        self._set_running_state(False)
        InfoBar.error(
            title="Run gagal",
            content=message,
            duration=5000,
            parent=self,
            position=InfoBarPosition.TOP_RIGHT,
        )

    def _run_stopped(self):
        self._append_log("=== RUN STOPPED ===")
        self.status_label.setText("Status: idle")
        self._set_running_state(False)
        InfoBar.warning(
            title="Run dihentikan",
            content="Proses dihentikan oleh pengguna.",
            duration=3000,
            parent=self,
            position=InfoBarPosition.TOP_RIGHT,
        )

    def _update_progress(self, processed, total, excel_row):
        if total <= 0:
            self.progress_label.setText("Progress: -")
            return
        text = f"Progress: {processed}/{total}"
        if excel_row and excel_row > 0:
            text = f"{text} | Baris Excel {excel_row}"
        self.progress_label.setText(text)

    def _show_error(self, message):
        InfoBar.error(
            title="Input tidak valid",
            content=message,
            duration=4000,
            parent=self,
            position=InfoBarPosition.TOP_RIGHT,
        )

    def _confirm_dialog(self, title, message):
        result = QMessageBox.question(
            self,
            title,
            message,
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        return result == QMessageBox.Yes

    def _set_running_state(self, running):
        self.run_button.setEnabled(not running)
        self.stop_button.setEnabled(running)
        self._set_controls_enabled(not running)

    def _set_controls_enabled(self, enabled):
        for widget in [
            self.excel_input,
            self.excel_browse,
            self.recent_combo,
            self.manual_switch,
            self.headless_switch,
            self.keep_open_switch,
            self.idle_spin,
            self.web_timeout_spin,
            self.range_switch,
            self.start_spin,
            self.end_spin,
        ]:
            widget.setEnabled(enabled)

        if enabled:
            self._toggle_range()
            self._sync_sso_state()
        else:
            self.start_spin.setEnabled(False)
            self.end_spin.setEnabled(False)

        if self._sso_page:
            self._sso_page.set_controls_enabled(enabled)
    def _show_keep_open_dialog(self):
        self.status_label.setText("Status: waiting for browser close")
        dialog = QDialog(self)
        dialog.setWindowTitle("Browser terbuka")
        layout = QVBoxLayout(dialog)
        layout.addWidget(
            BodyLabel(
                "Browser masih terbuka. Klik tombol di bawah untuk menutup."
            )
        )
        close_button = PrimaryPushButton("Close browser")
        close_button.clicked.connect(lambda: self._close_browser(dialog))
        layout.addWidget(close_button)
        dialog.rejected.connect(lambda: self._close_browser(dialog))
        dialog.exec()

    def _close_browser(self, dialog):
        if self._worker:
            self._worker.release_close()
        if dialog.isVisible():
            dialog.accept()


class HomePage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(16)

        hero_card = CardWidget()
        hero_layout = QHBoxLayout(hero_card)
        hero_layout.setContentsMargins(16, 16, 16, 16)
        hero_layout.setSpacing(16)

        hero_icon = IconWidget()
        hero_icon.setIcon(FIF.INFO)
        hero_icon.setFixedSize(36, 36)
        hero_layout.addWidget(hero_icon, alignment=Qt.AlignTop)

        hero_text = QWidget()
        hero_text_layout = QVBoxLayout(hero_text)
        hero_text_layout.setContentsMargins(0, 0, 0, 0)
        hero_text_layout.setSpacing(6)

        title = LargeTitleLabel("DIRGC Automation")
        subtitle = BodyLabel(
            "Alat bantu untuk mempercepat proses Ground Check "
            "berdasarkan data Excel di portal DIRGC."
        )
        subtitle.setWordWrap(True)
        hero_text_layout.addWidget(title)
        hero_text_layout.addWidget(subtitle)
        hero_layout.addWidget(hero_text, stretch=1)

        layout.addWidget(hero_card)

        summary_card = CardWidget()
        summary_layout = QVBoxLayout(summary_card)
        summary_layout.addWidget(SubtitleLabel("Ringkasan"))

        summary_text = BodyLabel(
            "Fokus utama: mempercepat input GC, mengurangi kesalahan manual, "
            "dan menghasilkan log hasil run untuk monitoring."
        )
        summary_text.setWordWrap(True)
        summary_layout.addWidget(summary_text)
        layout.addWidget(summary_card)

        steps_card = CardWidget()
        steps_layout = QVBoxLayout(steps_card)
        steps_layout.addWidget(SubtitleLabel("Cara Pakai Singkat"))

        steps = [
            (
                "1. Siapkan Excel",
                "Pastikan file Excel mengikuti format kolom yang disarankan.",
            ),
            (
                "2. Isi Akun SSO",
                "Buka menu Akun SSO, aktifkan switch, lalu isi username "
                "dan password.",
            ),
            (
                "3. Jalankan",
                "Buka menu Run, pilih file Excel, atur opsi, lalu klik Mulai.",
            ),
            (
                "4. Pantau hasil",
                "Log tampil di aplikasi dan file output tersimpan di folder logs.",
            ),
        ]
        for title_text, desc_text in steps:
            title_label = StrongBodyLabel(title_text)
            desc_label = CaptionLabel(desc_text)
            desc_label.setWordWrap(True)
            desc_label.setStyleSheet("color: #6B6B6B;")
            steps_layout.addWidget(title_label)
            steps_layout.addWidget(desc_label)
        layout.addWidget(steps_card)

        notes_card = CardWidget()
        notes_layout = QVBoxLayout(notes_card)
        notes_layout.addWidget(SubtitleLabel("Keterangan Opsi"))

        notes = [
            (
                "Login manual (tanpa auto-login)",
                "ON: login dilakukan manual di browser. OFF: gunakan Akun SSO.",
            ),
            (
                "Browser tanpa tampilan (headless)",
                "ON: browser tidak terlihat. Tidak disarankan untuk SSO/OTP.",
            ),
            (
                "Biarkan browser tetap terbuka",
                "ON: browser tetap terbuka setelah proses selesai.",
            ),
            (
                "Batas idle (detik)",
                "Jika tidak ada aktivitas, proses dihentikan otomatis.",
            ),
            (
                "Timeout loading web (detik)",
                "Naikkan jika halaman sering lambat saat login atau load data.",
            ),
            (
                "Batasi baris Excel",
                "ON: hanya memproses baris Start-End dari Excel.",
            ),
        ]
        for title_text, desc_text in notes:
            title_label = StrongBodyLabel(title_text)
            desc_label = CaptionLabel(desc_text)
            desc_label.setWordWrap(True)
            desc_label.setStyleSheet("color: #6B6B6B;")
            notes_layout.addWidget(title_label)
            notes_layout.addWidget(desc_label)
        layout.addWidget(notes_card)
        layout.addStretch()
        layout.addWidget(build_footer_label())


class SsoPage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(16)

        title = TitleLabel("Akun SSO")
        subtitle = BodyLabel(
            "Isi kredensial untuk auto-login. Data hanya dipakai saat proses berjalan."
        )
        layout.addWidget(title)
        layout.addWidget(subtitle)

        card = CardWidget()
        card_layout = QVBoxLayout(card)
        card_layout.addWidget(SubtitleLabel("Kredensial"))

        self.use_switch = SwitchButton()
        self.use_switch.setChecked(False)
        card_layout.addWidget(
            self._make_toggle_row(
                "Gunakan kredensial SSO",
                "Aktifkan jika ingin auto-login di halaman SSO.",
                self.use_switch,
            )
        )

        form = QFormLayout()
        form.setHorizontalSpacing(12)
        form.setVerticalSpacing(8)
        form.setLabelAlignment(Qt.AlignLeft | Qt.AlignVCenter)

        self.username_input = LineEdit()
        self.username_input.setPlaceholderText("username.sso")
        self.password_input = PasswordLineEdit()
        self.password_input.setPlaceholderText("password")

        form.addRow(BodyLabel("SSO Username"), self.username_input)
        form.addRow(BodyLabel("SSO Password"), self.password_input)
        card_layout.addLayout(form)

        layout.addWidget(card)
        layout.addStretch()
        layout.addWidget(build_footer_label())

        self.use_switch.checkedChanged.connect(self._toggle_fields)
        self._toggle_fields()

    def _make_toggle_row(self, text, description, switch):
        row = QWidget()
        row_layout = QHBoxLayout(row)
        row_layout.setContentsMargins(0, 0, 0, 0)
        row_layout.setSpacing(12)

        text_block = QWidget()
        text_layout = QVBoxLayout(text_block)
        text_layout.setContentsMargins(0, 0, 0, 0)
        text_layout.setSpacing(2)
        text_layout.addWidget(StrongBodyLabel(text))
        hint = CaptionLabel(description)
        hint.setWordWrap(True)
        hint.setStyleSheet("color: #6B6B6B;")
        text_layout.addWidget(hint)

        row_layout.addWidget(text_block, stretch=1)
        row_layout.addWidget(switch)
        return row

    def _toggle_fields(self):
        enabled = self.use_switch.isChecked() and self.use_switch.isEnabled()
        self.username_input.setEnabled(enabled)
        self.password_input.setEnabled(enabled)

    def is_enabled(self):
        return self.use_switch.isChecked()

    def get_credentials(self):
        if not self.is_enabled():
            return None, None
        username = self.username_input.text().strip()
        password = self.password_input.text()
        return username, password

    def set_controls_enabled(self, enabled):
        self.use_switch.setEnabled(enabled)
        self._toggle_fields()


class PlaceholderPage(QWidget):
    def __init__(self, title, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.addWidget(TitleLabel(title))
        layout.addWidget(BodyLabel("Halaman ini akan diisi di iterasi berikutnya."))
        layout.addStretch()
        layout.addWidget(build_footer_label())


class MainWindow(FluentWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("DIRGC Automation")
        self.resize(1100, 720)

        self.home_page = HomePage(self)
        self.sso_page = SsoPage(self)
        self.run_page = RunPage(self.sso_page, self)
        self.settings_page = PlaceholderPage("Settings", self)
        self.home_page.setObjectName("home_page")
        self.run_page.setObjectName("run_page")
        self.sso_page.setObjectName("sso_page")
        self.settings_page.setObjectName("settings_page")

        self.addSubInterface(self.home_page, FIF.HOME, "Beranda")
        self.addSubInterface(self.sso_page, FIF.PEOPLE, "Akun SSO")
        self.addSubInterface(self.run_page, FIF.PLAY, "Run")
        self.addSubInterface(
            self.settings_page,
            FIF.SETTING,
            "Settings",
            NavigationItemPosition.BOTTOM,
        )

    def closeEvent(self, event):
        if self.run_page:
            self.run_page._save_settings()
        super().closeEvent(event)


def main():
    app = QApplication(sys.argv)
    apply_app_font(app)
    setTheme(Theme.LIGHT)
    setThemeColor("#0078D4")

    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
