import json
import os
import re
import sys
import threading
import time
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
    QFrame,
    QPlainTextEdit,
    QProgressBar,
    QScrollArea,
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
    setFontFamilies,
    setTheme,
    setThemeColor,
)

from dirgc.cli import run_dirgc, validate_row_range
from dirgc.logging_utils import set_log_handler
from dirgc.resume_state import load_resume_state
from dirgc.settings import (
    DEFAULT_EXCEL_FILE,
    DEFAULT_IDLE_TIMEOUT_MS,
    DEFAULT_RECAP_LENGTH,
    DEFAULT_RATE_LIMIT_PROFILE,
    DEFAULT_SESSION_REFRESH_EVERY,
    DEFAULT_WEB_TIMEOUT_S,
    RATE_LIMIT_PROFILES,
    RECAP_LENGTH_MAX,
    RECAP_LENGTH_WARN_THRESHOLD,
)

GUI_SETTINGS_PATH = os.path.join("config", "gui_settings.json")
MAX_RECENT_EXCEL = 8
RESPONSIVE_BREAKPOINT = 980
BASE_FONT_SIZE = 11
DEFAULT_FONT_SCALE = 100
FONT_SCALE_OPTIONS = (100, 110, 120, 125)
FONT_BASE_PX_PROPERTY = "font_base_px"
FONT_BASE_PT_PROPERTY = "font_base_pt"
MUTED_TEXT_COLOR = "#4A4A4A"
RATE_LIMIT_SETTINGS_KEY = "rate_limit"
MIN_FONT_SIZE_PT = 10.0
ADVANCED_SETTINGS_KEY = "advanced"
IDLE_TIMEOUT_MIN_S = 30
IDLE_TIMEOUT_MAX_S = 3600 * 6
WEB_TIMEOUT_MIN_S = 10
WEB_TIMEOUT_MAX_S = 600


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
        'Made with ❤️ and ☕ - <a href="https://www.linkedin.com/in/novanniindipradana">'
        "Novanni Indi Pradana</a> - IPDS BPS 6502"
    )
    footer.setAlignment(Qt.AlignCenter)
    footer.setTextFormat(Qt.RichText)
    footer.setTextInteractionFlags(Qt.TextBrowserInteraction)
    footer.setOpenExternalLinks(True)
    footer.setStyleSheet(
        f"color: {MUTED_TEXT_COLOR};"
        f"QLabel a {{ color: {MUTED_TEXT_COLOR}; text-decoration: none; }}"
        "QLabel a:hover { text-decoration: underline; }"
    )
    return footer


def build_scroll_area(parent):
    scroll = QScrollArea(parent)
    scroll.setWidgetResizable(True)
    scroll.setFrameShape(QFrame.NoFrame)
    scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
    container = QWidget()
    scroll.setWidget(container)
    layout = QVBoxLayout(container)
    layout.setContentsMargins(24, 24, 24, 24)
    layout.setSpacing(16)
    return scroll, layout


def _normalize_font_scale(value):
    try:
        value = int(value)
    except (TypeError, ValueError):
        return DEFAULT_FONT_SCALE
    if value not in FONT_SCALE_OPTIONS:
        return DEFAULT_FONT_SCALE
    return value


def load_font_scale():
    data = load_gui_settings()
    ui_settings = data.get("ui", {})
    return _normalize_font_scale(ui_settings.get("font_scale"))


def save_font_scale(value):
    data = load_gui_settings()
    ui_settings = data.get("ui")
    if not isinstance(ui_settings, dict):
        ui_settings = {}
        data["ui"] = ui_settings
    ui_settings["font_scale"] = _normalize_font_scale(value)
    save_gui_settings(data)


def _normalize_rate_limit_profile(value):
    key = (value or "").strip().lower()
    if key in RATE_LIMIT_PROFILES:
        return key
    return DEFAULT_RATE_LIMIT_PROFILE


def load_rate_limit_profile():
    data = load_gui_settings()
    options = data.get(RATE_LIMIT_SETTINGS_KEY, {})
    profile = None
    if isinstance(options, dict):
        profile = options.get("profile")
    return _normalize_rate_limit_profile(profile)


def save_rate_limit_profile(value):
    data = load_gui_settings()
    options = data.get(RATE_LIMIT_SETTINGS_KEY)
    if not isinstance(options, dict):
        options = {}
        data[RATE_LIMIT_SETTINGS_KEY] = options
    options["profile"] = _normalize_rate_limit_profile(value)
    save_gui_settings(data)


def _normalize_int_setting(value, default, minimum, maximum):
    try:
        value = int(value)
    except (TypeError, ValueError):
        return default
    if value < minimum:
        return minimum
    if value > maximum:
        return maximum
    return value


def _load_legacy_timeout(data, key):
    for options_key in ("options", "options_update"):
        options = data.get(options_key)
        if isinstance(options, dict) and key in options:
            return options.get(key)
    return None


def load_idle_timeout_s():
    data = load_gui_settings()
    advanced = data.get(ADVANCED_SETTINGS_KEY, {})
    value = None
    if isinstance(advanced, dict):
        value = advanced.get("idle_timeout_s")
    if value is None:
        value = _load_legacy_timeout(data, "idle_timeout_s")
    return _normalize_int_setting(
        value,
        DEFAULT_IDLE_TIMEOUT_MS // 1000,
        IDLE_TIMEOUT_MIN_S,
        IDLE_TIMEOUT_MAX_S,
    )


def load_web_timeout_s():
    data = load_gui_settings()
    advanced = data.get(ADVANCED_SETTINGS_KEY, {})
    value = None
    if isinstance(advanced, dict):
        value = advanced.get("web_timeout_s")
    if value is None:
        value = _load_legacy_timeout(data, "web_timeout_s")
    return _normalize_int_setting(
        value,
        DEFAULT_WEB_TIMEOUT_S,
        WEB_TIMEOUT_MIN_S,
        WEB_TIMEOUT_MAX_S,
    )


def save_advanced_timeouts(idle_timeout_s, web_timeout_s):
    data = load_gui_settings()
    advanced = data.get(ADVANCED_SETTINGS_KEY)
    if not isinstance(advanced, dict):
        advanced = {}
        data[ADVANCED_SETTINGS_KEY] = advanced
    advanced["idle_timeout_s"] = _normalize_int_setting(
        idle_timeout_s,
        DEFAULT_IDLE_TIMEOUT_MS // 1000,
        IDLE_TIMEOUT_MIN_S,
        IDLE_TIMEOUT_MAX_S,
    )
    advanced["web_timeout_s"] = _normalize_int_setting(
        web_timeout_s,
        DEFAULT_WEB_TIMEOUT_S,
        WEB_TIMEOUT_MIN_S,
        WEB_TIMEOUT_MAX_S,
    )
    save_gui_settings(data)


def _font_point_size_for_widget(widget, font):
    point_size = font.pointSizeF()
    if point_size > 0:
        return point_size
    pixel_size = font.pixelSize()
    if pixel_size <= 0:
        return None
    dpi = widget.logicalDpiY() if hasattr(widget, "logicalDpiY") else 96
    if dpi <= 0:
        dpi = 96
    return pixel_size * 72.0 / dpi


def _apply_font_scale_to_widget(widget, font_scale):
    base_pt = widget.property(FONT_BASE_PT_PROPERTY)
    base_px = widget.property(FONT_BASE_PX_PROPERTY)
    scale_factor = font_scale / 100.0
    if base_pt is None:
        if base_px is not None:
            dpi = widget.logicalDpiY() if hasattr(widget, "logicalDpiY") else 96
            if dpi <= 0:
                dpi = 96
            base_pt = float(base_px) * 72.0 / dpi
            widget.setProperty(FONT_BASE_PT_PROPERTY, base_pt)
        else:
            font = widget.font()
            base_pt = _font_point_size_for_widget(widget, font)
            if base_pt is None:
                return
            widget.setProperty(FONT_BASE_PT_PROPERTY, base_pt)

    font = widget.font()
    target_pt = max(MIN_FONT_SIZE_PT, base_pt * scale_factor)
    font.setPointSizeF(target_pt)
    widget.setFont(font)


def apply_font_scale(app, font_scale):
    font_scale = _normalize_font_scale(font_scale)
    for widget in QApplication.allWidgets():
        _apply_font_scale_to_widget(widget, font_scale)

    base_size = max(MIN_FONT_SIZE_PT, BASE_FONT_SIZE * (font_scale / 100.0))
    font = app.font()
    font.setPointSizeF(base_size)
    app.setFont(font)


def apply_app_font(app):
    font_path = os.path.join("assets", "fonts", "Poppins-Regular.ttf")
    if os.path.exists(font_path):
        QFontDatabase.addApplicationFont(font_path)
    families = QFontDatabase().families()
    if "Poppins" in families:
        font_family = "Poppins"
    else:
        font_family = "Segoe UI Variable"
    app.setFont(QFont(font_family, BASE_FONT_SIZE))
    setFontFamilies([font_family, "Segoe UI Variable", "Segoe UI"])


class TogglePasswordLineEdit(PasswordLineEdit):
    def __init__(self, parent=None):
        super().__init__(parent)
        # Replace press-and-hold with click-to-toggle behavior.
        self.viewButton.removeEventFilter(self)
        self.viewButton.clicked.connect(self._toggle_password_visibility)

    def _toggle_password_visibility(self):
        self.setPasswordVisible(not self.isPasswordVisible())


@dataclass
class RunConfig:
    headless: bool
    manual_only: bool
    excel_file: Optional[str]
    start_row: Optional[int]
    end_row: Optional[int]
    idle_timeout_ms: int
    keep_open: bool
    dirgc_only: bool
    edit_nama_alamat: bool
    prefer_excel_coords: bool
    update_mode: bool
    update_fields: Optional[list]
    use_sso: bool
    sso_username: Optional[str]
    sso_password: Optional[str]
    web_timeout_s: int
    rate_limit_profile: str
    submit_mode: str
    session_refresh_every: int
    stop_on_cooldown: bool
    recap: bool
    recap_length: int
    recap_output_dir: Optional[str]
    recap_sleep_ms: int
    recap_max_retries: int
    recap_backup_every: int
    recap_resume: bool


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
                dirgc_only=self._config.dirgc_only,
                edit_nama_alamat=self._config.edit_nama_alamat,
                prefer_excel_coords=self._config.prefer_excel_coords,
                update_mode=self._config.update_mode,
                update_fields=self._config.update_fields,
                credentials=credentials,
                rate_limit_profile=self._config.rate_limit_profile,
                submit_mode=self._config.submit_mode,
                session_refresh_every=self._config.session_refresh_every,
                stop_on_cooldown=self._config.stop_on_cooldown,
                recap=self._config.recap,
                recap_length=self._config.recap_length,
                recap_output_dir=self._config.recap_output_dir,
                recap_sleep_ms=self._config.recap_sleep_ms,
                recap_max_retries=self._config.recap_max_retries,
                recap_backup_every=self._config.recap_backup_every,
                recap_resume=self._config.recap_resume,
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
    def __init__(
        self,
        sso_page=None,
        parent=None,
        *,
        update_mode_default=False,
        title_text="Run DIRGC",
        subtitle_text="Atur file, opsi, lalu jalankan proses.",
        run_label="Mulai",
        run_card_title="Run",
        confirm_title="Mulai proses",
        confirm_message="Mulai proses sekarang?",
        settings_key="options",
    ):
        super().__init__(parent)
        self._worker = None
        self._sso_page = sso_page
        self._recent_excels = []
        self._update_mode_default = bool(update_mode_default)
        self._confirm_start_title = confirm_title
        self._confirm_start_message = confirm_message
        self._run_label = run_label
        self._run_card_title = run_card_title
        self._settings_key = settings_key
        self._update_fields = {}
        self._cooldown_active = False
        self._show_dirgc_only = not self._update_mode_default
        self._show_edit_nama_alamat = not self._update_mode_default
        self._show_prefer_web_coords = not self._update_mode_default
        self._show_range = True
        self._show_keep_open = True

        outer_layout = QVBoxLayout(self)
        outer_layout.setContentsMargins(0, 0, 0, 0)
        outer_layout.setSpacing(0)
        scroll, layout = build_scroll_area(self)
        outer_layout.addWidget(scroll)

        title = TitleLabel(title_text)
        subtitle = BodyLabel(subtitle_text)
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

        self.resume_button = PushButton("Lanjutkan dari resume_state")
        self.resume_button.clicked.connect(self._apply_resume_state)

        form.addRow(BodyLabel("Excel file"), excel_row)
        form.addRow(BodyLabel("Recent"), self.recent_combo)
        form.addRow(BodyLabel("Resume"), self.resume_button)

        card_layout.addLayout(form)

        self._apply_default_paths()

        return card

    def _build_options_card(self):
        card = CardWidget()
        card_layout = QVBoxLayout(card)
        card_layout.addWidget(SubtitleLabel("Options"))

        self.dirgc_only_switch = SwitchButton()
        self.dirgc_only_switch.setChecked(False)
        self.dirgc_only_switch.checkedChanged.connect(
            self._toggle_dirgc_only
        )
        dirgc_row = self._make_option_row(
            "Hanya sampai halaman DIRGC",
            "ON: login lalu berhenti di halaman DIRGC tanpa filter/input.",
            self.dirgc_only_switch,
        )
        card_layout.addWidget(dirgc_row)
        dirgc_row.setVisible(self._show_dirgc_only)

        self.edit_nama_alamat_switch = SwitchButton()
        self.edit_nama_alamat_switch.setChecked(False)
        edit_row = self._make_option_row(
            "Edit Nama/Alamat Usaha dari Excel",
            "ON: aktifkan toggle edit di popup dan isi dari data Excel.",
            self.edit_nama_alamat_switch,
        )
        card_layout.addWidget(edit_row)
        edit_row.setVisible(self._show_edit_nama_alamat)

        self.prefer_web_coords_switch = SwitchButton()
        self.prefer_web_coords_switch.setChecked(False)
        coords_row = self._make_option_row(
            "Prioritaskan koordinat web",
            "ON: jika koordinat web sudah terisi, tidak dioverwrite. "
            "OFF: gunakan koordinat dari Excel.",
            self.prefer_web_coords_switch,
        )
        card_layout.addWidget(coords_row)
        coords_row.setVisible(self._show_prefer_web_coords)
        if self._update_mode_default:
            card_layout.addWidget(SubtitleLabel("Field update"))

            self._update_fields = {
                "hasil_gc": SwitchButton(),
                "nama_usaha": SwitchButton(),
                "alamat": SwitchButton(),
                "koordinat": SwitchButton(),
            }
            for switch in self._update_fields.values():
                switch.setChecked(True)

            card_layout.addWidget(
                self._make_option_row(
                    "Hasil GC",
                    "ON: perbarui nilai Hasil GC sesuai Excel.",
                    self._update_fields["hasil_gc"],
                )
            )
            card_layout.addWidget(
                self._make_option_row(
                    "Nama usaha",
                    "ON: perbarui nama usaha dari Excel.",
                    self._update_fields["nama_usaha"],
                )
            )
            card_layout.addWidget(
                self._make_option_row(
                    "Alamat usaha",
                    "ON: perbarui alamat usaha dari Excel.",
                    self._update_fields["alamat"],
                )
            )
            card_layout.addWidget(
                self._make_option_row(
                    "Koordinat (Lat/Long)",
                    "ON: perbarui latitude dan longitude dari Excel.",
                    self._update_fields["koordinat"],
                )
            )

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
        range_row.setVisible(self._show_range)

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
        range_inputs.setVisible(self._show_range)

        self.advanced_switch = SwitchButton()
        self.advanced_switch.setChecked(False)
        self.advanced_switch.checkedChanged.connect(self._toggle_advanced)
        card_layout.addWidget(
            self._make_option_row(
                "Opsi lanjutan",
                "Tampilkan opsi teknis seperti cooldown, refresh session, "
                "dan keep open.",
                self.advanced_switch,
            )
        )

        self.advanced_container = QWidget()
        advanced_layout = QVBoxLayout(self.advanced_container)
        advanced_layout.setContentsMargins(0, 0, 0, 0)
        advanced_layout.setSpacing(8)

        self.stop_on_cooldown_switch = SwitchButton()
        self.stop_on_cooldown_switch.setChecked(False)
        cooldown_row = self._make_option_row(
            "Stop saat cooldown (simpan posisi)",
            "ON: jika server meminta jeda, proses dihentikan dan baris "
            "terakhir disimpan untuk dilanjutkan.",
            self.stop_on_cooldown_switch,
        )
        advanced_layout.addWidget(cooldown_row)

        self.submit_request_switch = SwitchButton()
        self.submit_request_switch.setChecked(False)
        advanced_layout.addWidget(
            self._make_option_row(
                "Submit via request (API)",
                "ON: kirim POST langsung ke endpoint konfirmasi. "
                "Pastikan sesuai aturan akses yang berlaku.",
                self.submit_request_switch,
            )
        )

        refresh_row = QWidget()
        refresh_layout = QHBoxLayout(refresh_row)
        refresh_layout.setContentsMargins(0, 0, 0, 0)
        refresh_layout.setSpacing(12)

        refresh_text = QWidget()
        refresh_text_layout = QVBoxLayout(refresh_text)
        refresh_text_layout.setContentsMargins(0, 0, 0, 0)
        refresh_text_layout.setSpacing(4)

        refresh_title = StrongBodyLabel("Auto refresh session")
        refresh_desc = CaptionLabel(
            "Refresh session tiap N submit sukses. 0 = nonaktif."
        )
        refresh_desc.setWordWrap(True)
        refresh_desc.setStyleSheet(f"color: {MUTED_TEXT_COLOR};")

        refresh_text_layout.addWidget(refresh_title)
        refresh_text_layout.addWidget(refresh_desc)

        self.session_refresh_spin = QSpinBox()
        self.session_refresh_spin.setRange(0, 1000000)
        self.session_refresh_spin.setValue(DEFAULT_SESSION_REFRESH_EVERY)

        refresh_layout.addWidget(refresh_text, stretch=1)
        refresh_layout.addWidget(self.session_refresh_spin)
        advanced_layout.addWidget(refresh_row)

        self.keep_open_switch = SwitchButton()
        self.keep_open_switch.setChecked(True)
        keep_open_row = self._make_option_row(
            "Biarkan browser tetap terbuka",
            "ON: browser tetap terbuka sampai kamu menutupnya.",
            self.keep_open_switch,
        )
        advanced_layout.addWidget(keep_open_row)
        keep_open_row.setVisible(self._show_keep_open)

        card_layout.addWidget(self.advanced_container)
        self._toggle_advanced()

        self._toggle_dirgc_only()

        return card

    def _build_run_card(self):
        card = CardWidget()
        card_layout = QVBoxLayout(card)
        card_layout.addWidget(SubtitleLabel(self._run_card_title))

        button_row = QWidget()
        button_layout = QHBoxLayout(button_row)
        button_layout.setContentsMargins(0, 0, 0, 0)
        button_layout.setSpacing(8)

        self.run_button = PrimaryPushButton(self._run_label)
        self.run_button.clicked.connect(self._confirm_start)
        self.stop_button = PushButton("Stop")
        self.stop_button.clicked.connect(self._confirm_stop)
        self.stop_button.setEnabled(False)
        button_layout.addWidget(self.run_button)
        button_layout.addWidget(self.stop_button)
        button_layout.addStretch()

        self.status_label = BodyLabel("Status: idle")
        self.progress_label = CaptionLabel("Progress: -")
        self.progress_label.setStyleSheet(f"color: {MUTED_TEXT_COLOR};")
        self.cooldown_label = CaptionLabel("Cooldown: -")
        self.cooldown_label.setStyleSheet(f"color: {MUTED_TEXT_COLOR};")
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(False)
        card_layout.addWidget(self.status_label)
        card_layout.addWidget(self.progress_label)
        card_layout.addWidget(self.cooldown_label)
        card_layout.addWidget(self.progress_bar)
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
        text_layout.setSpacing(4)

        title_label = StrongBodyLabel(title)
        desc_label = CaptionLabel(description)
        desc_label.setWordWrap(True)
        desc_label.setStyleSheet(f"color: {MUTED_TEXT_COLOR};")

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

    def _normalize_path(self, path):
        if not path:
            return ""
        try:
            return os.path.abspath(path)
        except OSError:
            return path

    def _apply_resume_state(self):
        state = load_resume_state()
        next_row = state.get("next_row")
        excel_file = state.get("excel_file") or ""
        saved_at = state.get("saved_at") or ""

        if not next_row:
            InfoBar.warning(
                title="Resume tidak tersedia",
                content="Resume state belum ditemukan.",
                duration=4000,
                parent=self,
                position=InfoBarPosition.TOP_RIGHT,
            )
            return

        if self.dirgc_only_switch.isChecked():
            self.dirgc_only_switch.setChecked(False)

        if excel_file:
            if not os.path.exists(excel_file):
                InfoBar.error(
                    title="Resume gagal",
                    content="File Excel dari resume_state tidak ditemukan.",
                    duration=5000,
                    parent=self,
                    position=InfoBarPosition.TOP_RIGHT,
                )
            else:
                self._set_excel_path(
                    excel_file, push_recent=False, save=False
                )

        self.range_switch.setChecked(True)
        try:
            next_row = int(next_row)
        except (TypeError, ValueError):
            next_row = None
        if next_row and next_row > 0:
            self.start_spin.setValue(next_row)
            if self.end_spin.value() < next_row:
                self.end_spin.setValue(next_row)
        self._toggle_range()

        content = (
            f"Mulai dari baris {next_row}."
            if next_row
            else "Resume diterapkan."
        )
        if saved_at:
            content = f"{content} (tersimpan {saved_at})"
        InfoBar.success(
            title="Resume diterapkan",
            content=content,
            duration=4000,
            parent=self,
            position=InfoBarPosition.TOP_RIGHT,
        )

    def _auto_apply_resume_state(self, config: RunConfig):
        if config.dirgc_only:
            return config
        state = load_resume_state()
        next_row = state.get("next_row")
        if not next_row:
            return config
        if config.start_row is not None:
            return config

        excel_file = config.excel_file or ""
        state_file = state.get("excel_file") or ""

        if excel_file and state_file:
            if self._normalize_path(excel_file) != self._normalize_path(
                state_file
            ):
                InfoBar.warning(
                    title="Auto-resume diabaikan",
                    content=(
                        "Resume state berasal dari file Excel berbeda."
                    ),
                    duration=4000,
                    parent=self,
                    position=InfoBarPosition.TOP_RIGHT,
                )
                return config
        elif not excel_file:
            if state_file:
                if not os.path.exists(state_file):
                    InfoBar.warning(
                        title="Auto-resume diabaikan",
                        content="File Excel dari resume_state tidak ditemukan.",
                        duration=4000,
                        parent=self,
                        position=InfoBarPosition.TOP_RIGHT,
                    )
                    return config
                config.excel_file = state_file
                self._set_excel_path(
                    state_file, push_recent=False, save=False
                )
            else:
                return config

        try:
            next_row = int(next_row)
        except (TypeError, ValueError):
            return config
        if next_row <= 0:
            return config

        config.start_row = next_row
        if config.end_row is not None and config.end_row < next_row:
            config.end_row = next_row

        InfoBar.success(
            title="Auto-resume",
            content=f"Mulai dari baris {next_row}.",
            duration=4000,
            parent=self,
            position=InfoBarPosition.TOP_RIGHT,
        )
        return config

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
        options = data.get(self._settings_key, {})
        excel_path = data.get("excel_path")
        recents = data.get("recent_excels", [])

        if isinstance(recents, list):
            self._recent_excels = [
                item for item in recents if isinstance(item, str)
            ]
            self._refresh_recent_combo()

        if isinstance(excel_path, str) and excel_path:
            self._set_excel_path(excel_path, push_recent=False, save=False)

        if self._show_keep_open and "keep_open" in options:
            self.keep_open_switch.setChecked(bool(options["keep_open"]))
        elif self._show_keep_open:
            self.keep_open_switch.setChecked(True)
        else:
            self.keep_open_switch.setChecked(False)
        if self._show_dirgc_only and "dirgc_only" in options:
            self.dirgc_only_switch.setChecked(bool(options["dirgc_only"]))
        else:
            self.dirgc_only_switch.setChecked(False)
        if self._show_edit_nama_alamat and "edit_nama_alamat" in options:
            self.edit_nama_alamat_switch.setChecked(
                bool(options["edit_nama_alamat"])
            )
        else:
            self.edit_nama_alamat_switch.setChecked(False)
        if self._show_prefer_web_coords and "prefer_web_coords" in options:
            self.prefer_web_coords_switch.setChecked(
                bool(options["prefer_web_coords"])
            )
        else:
            self.prefer_web_coords_switch.setChecked(False)
        if "stop_on_cooldown" in options:
            self.stop_on_cooldown_switch.setChecked(
                bool(options["stop_on_cooldown"])
            )
        else:
            self.stop_on_cooldown_switch.setChecked(False)
        if "submit_via_request" in options:
            self.submit_request_switch.setChecked(
                bool(options["submit_via_request"])
            )
        else:
            self.submit_request_switch.setChecked(False)
        if "session_refresh_every" in options:
            try:
                self.session_refresh_spin.setValue(
                    int(options["session_refresh_every"])
                )
            except (TypeError, ValueError):
                self.session_refresh_spin.setValue(
                    DEFAULT_SESSION_REFRESH_EVERY
                )
        else:
            self.session_refresh_spin.setValue(DEFAULT_SESSION_REFRESH_EVERY)
        if hasattr(self, "advanced_switch"):
            advanced_active = any(
                [
                    self.keep_open_switch.isChecked(),
                    self.stop_on_cooldown_switch.isChecked(),
                    self.submit_request_switch.isChecked(),
                    self.session_refresh_spin.value() > 0,
                ]
            )
            self.advanced_switch.setChecked(advanced_active)
        if self._update_fields:
            if "update_fields" in options and isinstance(
                options.get("update_fields"), list
            ):
                fields = {
                    item
                    for item in options.get("update_fields", [])
                    if isinstance(item, str)
                }
                for key, switch in self._update_fields.items():
                    switch.setChecked(key in fields)
            else:
                for switch in self._update_fields.values():
                    switch.setChecked(True)
        if self._show_range and "range_enabled" in options:
            self.range_switch.setChecked(bool(options["range_enabled"]))
        else:
            self.range_switch.setChecked(False)
        if self._show_range and "start_row" in options:
            self.start_spin.setValue(int(options["start_row"]))
        if self._show_range and "end_row" in options:
            self.end_spin.setValue(int(options["end_row"]))

        self._toggle_dirgc_only()

    def _save_settings(self):
        data = load_gui_settings()
        data["excel_path"] = self.excel_input.text().strip()
        data["recent_excels"] = self._recent_excels
        options = {
            "keep_open": self.keep_open_switch.isChecked(),
            "dirgc_only": self.dirgc_only_switch.isChecked(),
            "edit_nama_alamat": self.edit_nama_alamat_switch.isChecked(),
            "prefer_web_coords": self.prefer_web_coords_switch.isChecked(),
            "stop_on_cooldown": self.stop_on_cooldown_switch.isChecked(),
            "submit_via_request": self.submit_request_switch.isChecked(),
            "session_refresh_every": self.session_refresh_spin.value(),
            "range_enabled": self.range_switch.isChecked(),
            "start_row": self.start_spin.value(),
            "end_row": self.end_spin.value(),
        }
        if self._update_fields:
            selected_fields = [
                key
                for key, switch in self._update_fields.items()
                if switch.isChecked()
            ]
            options["update_fields"] = selected_fields
        data[self._settings_key] = options
        save_gui_settings(data)

    def _toggle_range(self):
        if hasattr(self, "dirgc_only_switch") and self.dirgc_only_switch.isChecked():
            enabled = False
        else:
            enabled = self.range_switch.isChecked()
        self.start_spin.setEnabled(enabled)
        self.end_spin.setEnabled(enabled)

    def _toggle_dirgc_only(self):
        enabled = not self.dirgc_only_switch.isChecked()
        for widget in [
            self.excel_input,
            self.excel_browse,
            self.recent_combo,
            self.resume_button,
            self.range_switch,
            self.edit_nama_alamat_switch,
            self.prefer_web_coords_switch,
            self.stop_on_cooldown_switch,
            self.submit_request_switch,
            self.session_refresh_spin,
        ]:
            widget.setEnabled(enabled)
        for switch in self._update_fields.values():
            switch.setEnabled(enabled)
        if enabled:
            self._toggle_range()
        else:
            self.start_spin.setEnabled(False)
            self.end_spin.setEnabled(False)

    def _toggle_advanced(self):
        if hasattr(self, "advanced_container"):
            self.advanced_container.setVisible(
                self.advanced_switch.isChecked()
            )


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
        self._handle_cooldown_notifications(line)

    def _extract_log_field(self, line, key):
        if not line or not key:
            return ""
        pattern = rf"(?:^|\|\s){re.escape(key)}=([^|]+)"
        match = re.search(pattern, line)
        if not match:
            return ""
        value = match.group(1).strip()
        if value.startswith('"') and value.endswith('"'):
            value = value[1:-1]
        return value.strip()

    def _handle_cooldown_notifications(self, line):
        if not line:
            return
        if "Cooldown aktif; menunggu sebelum lanjut." in line:
            sisa = self._extract_log_field(line, "sisa")
            if sisa:
                self.cooldown_label.setText(f"Cooldown: {sisa}")
            self._cooldown_active = True
            return
        if "Cooldown aktif; menghentikan proses." in line:
            resume_at = self._extract_log_field(line, "resume_at")
            content = (
                "Server meminta jeda; proses dihentikan."
                if not resume_at
                else f"Server meminta jeda; proses dihentikan. Resume: {resume_at}"
            )
            self.cooldown_label.setText("Cooldown: -")
            InfoBar.warning(
                title="Cooldown aktif",
                content=content,
                duration=6000,
                parent=self,
                position=InfoBarPosition.TOP_RIGHT,
            )
            self._cooldown_active = False
            return

        if "Server meminta jeda sebelum " in line:
            if self._cooldown_active:
                return
            self._cooldown_active = True
            resume_at = self._extract_log_field(line, "resume_at")
            wait_s = self._extract_log_field(line, "wait_s")
            reason = self._extract_log_field(line, "reason")
            details = []
            if wait_s:
                details.append(f"Jeda {wait_s} detik")
                self.cooldown_label.setText(f"Cooldown: {wait_s}s")
            if resume_at:
                details.append(f"Lanjut {resume_at}")
            if reason and reason != "-":
                details.append(reason)
            content = " | ".join(details) if details else "Menunggu sesuai instruksi server."
            InfoBar.warning(
                title="Auto-pause aktif",
                content=content,
                duration=5000,
                parent=self,
                position=InfoBarPosition.TOP_RIGHT,
            )
            return

        if "Cooldown selesai; melanjutkan proses." in line:
            if not self._cooldown_active:
                return
            self._cooldown_active = False
            self.cooldown_label.setText("Cooldown: -")
            InfoBar.success(
                title="Auto-resume",
                content="Cooldown selesai. Proses dilanjutkan otomatis.",
                duration=3000,
                parent=self,
                position=InfoBarPosition.TOP_RIGHT,
            )

    def _build_config(self):
        excel_text = self.excel_input.text().strip()
        excel_file = excel_text if excel_text else None
        use_sso, sso_username, sso_password = self._get_sso_values()

        start_row = None
        end_row = None
        if self.range_switch.isChecked():
            start_row = self.start_spin.value()
            end_row = self.end_spin.value()

        update_fields = None
        if self._update_fields:
            selected = [
                key
                for key, switch in self._update_fields.items()
                if switch.isChecked()
            ]
            update_fields = selected

        idle_timeout_s = load_idle_timeout_s()
        web_timeout_s = load_web_timeout_s()
        submit_mode = "request" if self.submit_request_switch.isChecked() else "ui"
        session_refresh_every = self.session_refresh_spin.value()

        return RunConfig(
            headless=False,
            manual_only=not use_sso,
            excel_file=excel_file,
            start_row=start_row,
            end_row=end_row,
            idle_timeout_ms=idle_timeout_s * 1000,
            web_timeout_s=web_timeout_s,
            keep_open=self.keep_open_switch.isChecked(),
            dirgc_only=self.dirgc_only_switch.isChecked(),
            edit_nama_alamat=self.edit_nama_alamat_switch.isChecked(),
            prefer_excel_coords=not self.prefer_web_coords_switch.isChecked(),
            update_mode=self._update_mode_default,
            update_fields=update_fields,
            use_sso=use_sso,
            sso_username=sso_username,
            sso_password=sso_password,
            rate_limit_profile=load_rate_limit_profile(),
            submit_mode=submit_mode,
            session_refresh_every=session_refresh_every,
            stop_on_cooldown=self.stop_on_cooldown_switch.isChecked(),
            recap=False,
            recap_length=DEFAULT_RECAP_LENGTH,
            recap_output_dir=None,
            recap_sleep_ms=800,
            recap_max_retries=3,
            recap_backup_every=10,
            recap_resume=True,
        )

    def _validate_inputs(self, config: RunConfig):
        try:
            validate_row_range(config.start_row, config.end_row)
        except ValueError as exc:
            self._show_error(str(exc))
            return False

        if not config.dirgc_only:
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
        config = self._auto_apply_resume_state(config)
        if not self._validate_inputs(config):
            return

        self._save_settings()
        self.status_label.setText("Status: running")
        self._set_running_state(True)
        self.progress_label.setText("Progress: memuat data...")
        self.cooldown_label.setText("Cooldown: -")
        self._set_progress_loading()
        self._append_log("=== START RUN ===")

        if isinstance(self, RecapPage):
            self._recap_start_ts = None

        self._worker = RunWorker(config)
        self._worker.log_line.connect(self._append_log)
        self._worker.finished_ok.connect(self._run_finished)
        self._worker.failed.connect(self._run_failed)
        self._worker.request_close.connect(self._show_keep_open_dialog)
        self._worker.progress.connect(self._update_progress)
        self._worker.start()

    def _confirm_start(self):
        if self._confirm_dialog(
            self._confirm_start_title,
            self._confirm_start_message,
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
        self.progress_label.setText("Progress: -")
        self._reset_progress()
        self.cooldown_label.setText("Cooldown: -")
        self._cooldown_active = False
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
        self.progress_label.setText("Progress: -")
        self._reset_progress()
        self.cooldown_label.setText("Cooldown: -")
        self._cooldown_active = False
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
        self.progress_label.setText("Progress: -")
        self._reset_progress()
        self.cooldown_label.setText("Cooldown: -")
        self._cooldown_active = False
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
            self._set_progress_loading()
            return
        text = f"Progress: {processed}/{total}"
        if excel_row and excel_row > 0:
            text = f"{text} | Baris Excel {excel_row}"
        self.progress_label.setText(text)
        self.progress_bar.setRange(0, int(total))
        value = min(max(int(processed), 0), int(total))
        self.progress_bar.setValue(value)

    def _set_progress_loading(self):
        if hasattr(self, "progress_bar"):
            self.progress_bar.setRange(0, 0)
            self.progress_bar.setValue(0)

    def _reset_progress(self):
        if hasattr(self, "progress_bar"):
            self.progress_bar.setRange(0, 100)
            self.progress_bar.setValue(0)

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
        widgets = [
            self.excel_input,
            self.excel_browse,
            self.recent_combo,
            self.resume_button,
            self.advanced_switch,
            self.keep_open_switch,
            self.stop_on_cooldown_switch,
            self.dirgc_only_switch,
            self.edit_nama_alamat_switch,
            self.prefer_web_coords_switch,
            self.range_switch,
            self.start_spin,
            self.end_spin,
        ]
        for widget in widgets:
            widget.setEnabled(enabled)
        for switch in self._update_fields.values():
            switch.setEnabled(enabled)

        if enabled:
            self._toggle_dirgc_only()
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


class RecapPage(RunPage):
    def __init__(self, sso_page=None, parent=None):
        super().__init__(
            sso_page,
            parent,
            update_mode_default=False,
            title_text="Recap DIRGC",
            subtitle_text=(
                "Tarik rekap data DIRGC via API dan simpan ke Excel."
            ),
            run_label="Mulai Rekap",
            run_card_title="Rekap",
            confirm_title="Mulai rekap",
            confirm_message="Mulai rekap sekarang?",
            settings_key="options_recap",
        )
        self._recap_start_ts = None
        self._warned_large_page_size = False

    def _build_files_card(self):
        card = CardWidget()
        card_layout = QVBoxLayout(card)
        card_layout.addWidget(SubtitleLabel("Output"))

        form = QFormLayout()
        form.setHorizontalSpacing(12)
        form.setVerticalSpacing(8)
        form.setLabelAlignment(Qt.AlignLeft | Qt.AlignVCenter)

        self.output_dir_input = LineEdit()
        self.output_dir_input.setPlaceholderText("logs/recap")
        self.output_dir_input.editingFinished.connect(self._save_settings)
        self.output_dir_browse = PushButton("Browse")
        self.output_dir_browse.clicked.connect(self._browse_output_dir)

        output_row = QWidget()
        output_layout = QHBoxLayout(output_row)
        output_layout.setContentsMargins(0, 0, 0, 0)
        output_layout.setSpacing(8)
        output_layout.addWidget(self.output_dir_input, stretch=1)
        output_layout.addWidget(self.output_dir_browse)

        form.addRow(BodyLabel("Folder output"), output_row)
        card_layout.addLayout(form)

        hint = CaptionLabel("Kosong = default logs/recap.")
        hint.setStyleSheet(f"color: {MUTED_TEXT_COLOR};")
        card_layout.addWidget(hint)

        return card

    def _build_options_card(self):
        card = CardWidget()
        card_layout = QVBoxLayout(card)
        card_layout.addWidget(SubtitleLabel("Rekap Options"))

        form = QFormLayout()
        form.setHorizontalSpacing(12)
        form.setVerticalSpacing(8)
        form.setLabelAlignment(Qt.AlignLeft | Qt.AlignVCenter)

        self.length_spin = QSpinBox()
        self.length_spin.setRange(10, RECAP_LENGTH_MAX)
        self.length_spin.setValue(DEFAULT_RECAP_LENGTH)
        self.length_spin.valueChanged.connect(self._handle_page_size_change)

        self.sleep_spin = QSpinBox()
        self.sleep_spin.setRange(0, 10000)
        self.sleep_spin.setValue(800)
        self.sleep_spin.setSuffix(" ms")

        self.retry_spin = QSpinBox()
        self.retry_spin.setRange(0, 10)
        self.retry_spin.setValue(3)

        self.backup_spin = QSpinBox()
        self.backup_spin.setRange(0, 9999)
        self.backup_spin.setValue(10)
        self.backup_spin.setSuffix(" batch")

        form.addRow(BodyLabel("Page size"), self.length_spin)
        form.addRow(BodyLabel("Delay"), self.sleep_spin)
        form.addRow(BodyLabel("Max retries"), self.retry_spin)
        form.addRow(BodyLabel("Backup tiap N batch"), self.backup_spin)

        card_layout.addLayout(form)
        page_size_hint = CaptionLabel(
            "Page size = jumlah baris per request. Lebih besar lebih cepat, "
            "tapi lebih berat; server bisa membatasi dan otomatis turun."
        )
        page_size_hint.setStyleSheet(f"color: {MUTED_TEXT_COLOR};")
        card_layout.addWidget(page_size_hint)
        status_hint = CaptionLabel("Status filter dikunci: Semua.")
        status_hint.setStyleSheet(f"color: {MUTED_TEXT_COLOR};")
        card_layout.addWidget(status_hint)

        self.resume_switch = SwitchButton()
        self.resume_switch.setChecked(True)
        card_layout.addWidget(
            self._make_option_row(
                "Resume dari checkpoint",
                "ON: lanjutkan dari proses terakhir jika checkpoint tersedia.",
                self.resume_switch,
            )
        )

        self.keep_open_switch = SwitchButton()
        self.keep_open_switch.setChecked(True)
        card_layout.addWidget(
            self._make_option_row(
                "Biarkan browser tetap terbuka",
                "ON: browser tetap terbuka setelah proses selesai.",
                self.keep_open_switch,
            )
        )

        return card

    def _handle_page_size_change(self, value):
        if value > RECAP_LENGTH_WARN_THRESHOLD:
            if self._warned_large_page_size:
                return
            self._warned_large_page_size = True
            InfoBar.warning(
                title="Page size besar",
                content=(
                    f"Di atas {RECAP_LENGTH_WARN_THRESHOLD} akan dicoba, "
                    f"tetapi jika server membatasi akan otomatis turun "
                    f"ke {RECAP_LENGTH_WARN_THRESHOLD}."
                ),
                duration=4500,
                position=InfoBarPosition.TOP_RIGHT,
                parent=self,
            )
        else:
            self._warned_large_page_size = False

    def _browse_output_dir(self):
        start_dir = os.getcwd()
        if self.output_dir_input.text():
            start_dir = self.output_dir_input.text()
        path = QFileDialog.getExistingDirectory(
            self, "Select folder", start_dir
        )
        if path:
            self.output_dir_input.setText(path)
            self._save_settings()

    def _load_settings(self):
        data = load_gui_settings()
        options = data.get(self._settings_key, {})

        if isinstance(options, dict):
            output_dir = options.get("output_dir")
            if isinstance(output_dir, str):
                self.output_dir_input.setText(output_dir)

            for key, widget in (
                ("length", self.length_spin),
                ("sleep_ms", self.sleep_spin),
                ("max_retries", self.retry_spin),
                ("backup_every", self.backup_spin),
            ):
                if key in options:
                    try:
                        widget.setValue(int(options[key]))
                    except (TypeError, ValueError):
                        pass

            if "resume" in options:
                self.resume_switch.setChecked(bool(options["resume"]))
            if "keep_open" in options:
                self.keep_open_switch.setChecked(bool(options["keep_open"]))
            else:
                self.keep_open_switch.setChecked(True)

    def _save_settings(self):
        data = load_gui_settings()
        options = {
            "output_dir": self.output_dir_input.text().strip(),
            "length": self.length_spin.value(),
            "sleep_ms": self.sleep_spin.value(),
            "max_retries": self.retry_spin.value(),
            "backup_every": self.backup_spin.value(),
            "resume": self.resume_switch.isChecked(),
            "keep_open": self.keep_open_switch.isChecked(),
        }
        data[self._settings_key] = options
        save_gui_settings(data)

    def _build_config(self):
        use_sso, sso_username, sso_password = self._get_sso_values()
        idle_timeout_s = load_idle_timeout_s()
        web_timeout_s = load_web_timeout_s()

        output_dir = self.output_dir_input.text().strip()
        output_dir = output_dir if output_dir else None

        return RunConfig(
            headless=False,
            manual_only=not use_sso,
            excel_file=None,
            start_row=None,
            end_row=None,
            idle_timeout_ms=idle_timeout_s * 1000,
            web_timeout_s=web_timeout_s,
            keep_open=self.keep_open_switch.isChecked(),
            dirgc_only=False,
            edit_nama_alamat=False,
            prefer_excel_coords=True,
            update_mode=False,
            update_fields=None,
            use_sso=use_sso,
            sso_username=sso_username,
            sso_password=sso_password,
            rate_limit_profile=load_rate_limit_profile(),
            submit_mode="ui",
            session_refresh_every=0,
            stop_on_cooldown=False,
            recap=True,
            recap_length=self.length_spin.value(),
            recap_output_dir=output_dir,
            recap_sleep_ms=self.sleep_spin.value(),
            recap_max_retries=self.retry_spin.value(),
            recap_backup_every=self.backup_spin.value(),
            recap_resume=self.resume_switch.isChecked(),
        )

    def _validate_inputs(self, config: RunConfig):
        if config.use_sso:
            if config.manual_only:
                self._show_error(
                    "Matikan manual login untuk menggunakan Akun SSO."
                )
                return False
        return True

    def _update_progress(self, processed, total, _excel_row):
        if self._recap_start_ts is None and processed > 0:
            self._recap_start_ts = time.time()
        if total <= 0:
            self.progress_label.setText("Progress: memuat data...")
            self._set_progress_loading()
            return
        percent = int(round((processed / total) * 100))
        percent = min(max(percent, 0), 100)
        eta_text = ""
        speed_text = " | Speed (rows/sec): -"
        if self._recap_start_ts and processed > 0:
            elapsed = max(1.0, time.time() - self._recap_start_ts)
            rate = processed / elapsed
            remaining = max(0, total - processed)
            eta_s = int(remaining / rate) if rate > 0 else 0
            hours = eta_s // 3600
            minutes = (eta_s % 3600) // 60
            secs = eta_s % 60
            speed_text = f" | Speed (rows/sec): {rate:.1f}"
            eta_text = (
                f" | Estimated Time: {hours} Jam {minutes} Menit {secs} Detik"
            )
        self.progress_label.setText(
            f"Progress: {processed}/{total} ({percent}%){speed_text}{eta_text}"
        )
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(percent)

    def _set_controls_enabled(self, enabled):
        for widget in [
            self.output_dir_input,
            self.output_dir_browse,
            self.length_spin,
            self.sleep_spin,
            self.retry_spin,
            self.backup_spin,
            self.resume_switch,
            self.keep_open_switch,
        ]:
            widget.setEnabled(enabled)
        if self._sso_page:
            self._sso_page.set_controls_enabled(enabled)

    def _auto_apply_resume_state(self, config: RunConfig):
        return config

    def _open_log_folder(self):
        log_dir = os.path.join(os.getcwd(), "logs", "recap")
        os.makedirs(log_dir, exist_ok=True)
        if not QDesktopServices.openUrl(QUrl.fromLocalFile(log_dir)):
            InfoBar.error(
                title="Gagal membuka folder",
                content="Tidak bisa membuka folder recap.",
                duration=3000,
                parent=self,
                position=InfoBarPosition.TOP_RIGHT,
            )


class HomePage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        outer_layout = QVBoxLayout(self)
        outer_layout.setContentsMargins(0, 0, 0, 0)
        outer_layout.setSpacing(0)
        scroll, layout = build_scroll_area(self)
        outer_layout.addWidget(scroll)

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

        self._content_widget = QWidget()
        self._content_layout = QBoxLayout(QBoxLayout.LeftToRight)
        self._content_layout.setContentsMargins(0, 0, 0, 0)
        self._content_layout.setSpacing(16)
        self._content_widget.setLayout(self._content_layout)
        layout.addWidget(self._content_widget)

        left_col = QWidget()
        left_layout = QVBoxLayout(left_col)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(16)

        summary_card = CardWidget()
        summary_layout = QVBoxLayout(summary_card)
        summary_layout.addWidget(SubtitleLabel("Ringkasan"))

        summary_text = BodyLabel(
            "Fokus utama: mempercepat input GC, mengurangi kesalahan manual, "
            "dan menghasilkan log hasil run untuk monitoring."
        )
        summary_text.setWordWrap(True)
        summary_layout.addWidget(summary_text)
        left_layout.addWidget(summary_card)

        steps_card = CardWidget()
        steps_layout = QVBoxLayout(steps_card)
        steps_layout.setSpacing(6)
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
            desc_label.setStyleSheet(f"color: {MUTED_TEXT_COLOR};")
            steps_layout.addWidget(title_label)
            steps_layout.addWidget(desc_label)
        left_layout.addWidget(steps_card)
        left_layout.addStretch()

        right_col = QWidget()
        right_layout = QVBoxLayout(right_col)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(16)

        notes_card = CardWidget()
        notes_layout = QVBoxLayout(notes_card)
        notes_layout.setSpacing(6)
        notes_layout.addWidget(SubtitleLabel("Keterangan Opsi"))

        notes = [
            (
                "Biarkan browser tetap terbuka - Opsi lanjutan",
                "ON: browser tetap terbuka setelah proses selesai.",
            ),
            (
                "Stop saat cooldown (simpan posisi) - Opsi lanjutan",
                "ON: jika server meminta jeda, proses dihentikan dan baris "
                "terakhir disimpan untuk dilanjutkan.",
            ),
            (
                "Hanya sampai halaman DIRGC",
                "ON: berhenti di halaman DIRGC tanpa filter/input dari Excel.",
            ),
            (
                "Edit Nama/Alamat Usaha dari Excel",
                "ON: aktifkan toggle edit di popup dan isi dari data Excel.",
            ),
            (
                "Prioritaskan koordinat web",
                "ON: koordinat web dipertahankan; OFF: koordinat dari Excel.",
            ),
            (
                "Menu Update Data",
                "Gunakan menu Update untuk klik Edit Hasil dan memperbarui data.",
            ),
            (
                "Batas idle (detik) - Settings > Advanced",
                "Jika tidak ada aktivitas, proses dihentikan otomatis.",
            ),
            (
                "Timeout loading web (detik) - Settings > Advanced",
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
            desc_label.setStyleSheet(f"color: {MUTED_TEXT_COLOR};")
            notes_layout.addWidget(title_label)
            notes_layout.addWidget(desc_label)
        right_layout.addWidget(notes_card)

        appreciation_card = CardWidget()
        appreciation_layout = QVBoxLayout(appreciation_card)
        appreciation_layout.setSpacing(6)
        appreciation_layout.addWidget(SubtitleLabel("Dukungan"))
        appreciation_message = BodyLabel(
            'Kalau merasa terbantu, saya senang sekali jika Anda berkenan '
            'memberi ulasan di <a href="https://www.linkedin.com/in/novanniindipradana">'
            "LinkedIn</a>."
        )
        appreciation_message.setWordWrap(True)
        appreciation_message.setTextFormat(Qt.RichText)
        appreciation_message.setOpenExternalLinks(True)
        appreciation_message.setStyleSheet(f"color: {MUTED_TEXT_COLOR};")
        appreciation_layout.addWidget(appreciation_message)
        right_layout.addWidget(appreciation_card)
        right_layout.addStretch()

        self._content_layout.addWidget(left_col, stretch=1)
        self._content_layout.addWidget(right_col, stretch=1)

        layout.addStretch()
        layout.addWidget(build_footer_label())
        self._is_stacked = None
        self._update_layout_mode(self.width())

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._update_layout_mode(event.size().width())

    def _update_layout_mode(self, width):
        stacked = width < RESPONSIVE_BREAKPOINT
        if stacked == self._is_stacked:
            return
        direction = (
            QBoxLayout.TopToBottom if stacked else QBoxLayout.LeftToRight
        )
        self._content_layout.setDirection(direction)
        self._is_stacked = stacked


class SsoPage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        outer_layout = QVBoxLayout(self)
        outer_layout.setContentsMargins(0, 0, 0, 0)
        outer_layout.setSpacing(0)
        scroll, layout = build_scroll_area(self)
        outer_layout.addWidget(scroll)

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
        self.use_switch.setChecked(True)
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
        self.password_input = TogglePasswordLineEdit()
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
        text_layout.setSpacing(4)
        text_layout.addWidget(StrongBodyLabel(text))
        hint = CaptionLabel(description)
        hint.setWordWrap(True)
        hint.setStyleSheet(f"color: {MUTED_TEXT_COLOR};")
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


class SettingsPage(QWidget):
    def __init__(self, app, parent=None):
        super().__init__(parent)
        self._app = app
        outer_layout = QVBoxLayout(self)
        outer_layout.setContentsMargins(0, 0, 0, 0)
        outer_layout.setSpacing(0)
        scroll, layout = build_scroll_area(self)
        outer_layout.addWidget(scroll)

        title = TitleLabel("Settings")
        subtitle = BodyLabel("Atur tampilan aplikasi sesuai kebutuhan.")
        subtitle.setWordWrap(True)
        layout.addWidget(title)
        layout.addWidget(subtitle)

        appearance_card = CardWidget()
        appearance_layout = QVBoxLayout(appearance_card)
        appearance_layout.addWidget(SubtitleLabel("Tampilan"))

        row = QWidget()
        row_layout = QHBoxLayout(row)
        row_layout.setContentsMargins(0, 0, 0, 0)
        row_layout.setSpacing(12)

        text_block = QWidget()
        text_layout = QVBoxLayout(text_block)
        text_layout.setContentsMargins(0, 0, 0, 0)
        text_layout.setSpacing(4)
        text_layout.addWidget(StrongBodyLabel("Ukuran font"))
        desc_label = CaptionLabel(
            "Pilih 100/110/120/125% untuk memperbesar teks."
        )
        desc_label.setWordWrap(True)
        desc_label.setStyleSheet(f"color: {MUTED_TEXT_COLOR};")
        text_layout.addWidget(desc_label)

        self.font_combo = ComboBox()
        for scale in FONT_SCALE_OPTIONS:
            self.font_combo.addItem(f"{scale}%")
        current_scale = load_font_scale()
        index = self.font_combo.findText(f"{current_scale}%")
        if index >= 0:
            self.font_combo.setCurrentIndex(index)

        row_layout.addWidget(text_block, stretch=1)
        row_layout.addWidget(self.font_combo)
        appearance_layout.addWidget(row)
        layout.addWidget(appearance_card)

        advanced_card = CardWidget()
        advanced_layout = QVBoxLayout(advanced_card)
        advanced_layout.addWidget(SubtitleLabel("Advanced"))

        idle_row = QWidget()
        idle_layout = QHBoxLayout(idle_row)
        idle_layout.setContentsMargins(0, 0, 0, 0)
        idle_layout.setSpacing(12)

        idle_label = StrongBodyLabel("Batas idle (detik)")
        idle_hint = CaptionLabel(
            "Jika tidak ada aktivitas, proses dihentikan otomatis."
        )
        idle_hint.setStyleSheet(f"color: {MUTED_TEXT_COLOR};")
        idle_hint.setWordWrap(True)
        idle_text = QWidget()
        idle_text_layout = QVBoxLayout(idle_text)
        idle_text_layout.setContentsMargins(0, 0, 0, 0)
        idle_text_layout.setSpacing(4)
        idle_text_layout.addWidget(idle_label)
        idle_text_layout.addWidget(idle_hint)
        self.idle_timeout_spin = QSpinBox()
        self.idle_timeout_spin.setRange(IDLE_TIMEOUT_MIN_S, IDLE_TIMEOUT_MAX_S)
        self.idle_timeout_spin.setValue(load_idle_timeout_s())
        self.idle_timeout_spin.setSuffix(" s")
        idle_layout.addWidget(idle_text, stretch=1)
        idle_layout.addStretch()
        idle_layout.addWidget(self.idle_timeout_spin)
        advanced_layout.addWidget(idle_row)

        web_row = QWidget()
        web_layout = QHBoxLayout(web_row)
        web_layout.setContentsMargins(0, 0, 0, 0)
        web_layout.setSpacing(12)

        web_label = StrongBodyLabel("Timeout loading web (detik)")
        web_hint = CaptionLabel(
            "Naikkan jika koneksi lambat atau halaman sering timeout."
        )
        web_hint.setStyleSheet(f"color: {MUTED_TEXT_COLOR};")
        web_hint.setWordWrap(True)
        web_text = QWidget()
        web_text_layout = QVBoxLayout(web_text)
        web_text_layout.setContentsMargins(0, 0, 0, 0)
        web_text_layout.setSpacing(4)
        web_text_layout.addWidget(web_label)
        web_text_layout.addWidget(web_hint)
        self.web_timeout_spin = QSpinBox()
        self.web_timeout_spin.setRange(WEB_TIMEOUT_MIN_S, WEB_TIMEOUT_MAX_S)
        self.web_timeout_spin.setValue(load_web_timeout_s())
        self.web_timeout_spin.setSuffix(" s")
        web_layout.addWidget(web_text, stretch=1)
        web_layout.addStretch()
        web_layout.addWidget(self.web_timeout_spin)
        advanced_layout.addWidget(web_row)

        layout.addWidget(advanced_card)
        layout.addStretch()
        layout.addWidget(build_footer_label())

        self.font_combo.currentIndexChanged.connect(
            self._apply_font_scale
        )
        self.idle_timeout_spin.valueChanged.connect(
            self._save_advanced_timeouts
        )
        self.web_timeout_spin.valueChanged.connect(
            self._save_advanced_timeouts
        )

    def _apply_font_scale(self):
        text = self.font_combo.currentText().strip()
        text = text.replace("%", "")
        scale = _normalize_font_scale(text)
        save_font_scale(scale)
        apply_font_scale(self._app, scale)

    def _save_advanced_timeouts(self):
        save_advanced_timeouts(
            self.idle_timeout_spin.value(),
            self.web_timeout_spin.value(),
        )


class RateLimitPage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        outer_layout = QVBoxLayout(self)
        outer_layout.setContentsMargins(0, 0, 0, 0)
        outer_layout.setSpacing(0)
        scroll, layout = build_scroll_area(self)
        layout.setSpacing(12)
        outer_layout.addWidget(scroll)

        title = TitleLabel("Mode Stabilitas")
        subtitle = BodyLabel(
            "Mode ini mengatur jeda otomatis saat submit agar server "
            "tidak sering menolak permintaan."
        )
        subtitle.setWordWrap(True)
        layout.addWidget(title)
        layout.addWidget(subtitle)

        highlight_card = CardWidget()
        highlight_card.setStyleSheet(
            "background-color: #FFF7E6; border: 1px solid #FFD666;"
        )
        highlight_layout = QHBoxLayout(highlight_card)
        highlight_layout.setContentsMargins(14, 10, 14, 10)
        highlight_layout.setSpacing(10)

        highlight_icon = IconWidget()
        highlight_icon.setIcon(FIF.INFO)
        highlight_icon.setFixedSize(26, 26)
        highlight_layout.addWidget(highlight_icon, alignment=Qt.AlignTop)

        highlight_text = QWidget()
        highlight_text_layout = QVBoxLayout(highlight_text)
        highlight_text_layout.setContentsMargins(0, 0, 0, 0)
        highlight_text_layout.setSpacing(4)
        highlight_title = StrongBodyLabel("Sering gagal submit?")
        highlight_desc = BodyLabel(
            "Jika muncul pesan 'Something Went Wrong' saat submit, "
            "pilih mode Safe atau Ultra agar jeda lebih panjang."
        )
        highlight_desc.setWordWrap(True)
        highlight_text_layout.addWidget(highlight_title)
        highlight_text_layout.addWidget(highlight_desc)
        highlight_layout.addWidget(highlight_text, stretch=1)
        layout.addWidget(highlight_card)

        select_card = CardWidget()
        select_layout = QVBoxLayout(select_card)
        select_layout.setSpacing(8)
        select_layout.addWidget(SubtitleLabel("Pilih Mode"))

        row = QWidget()
        row_layout = QVBoxLayout(row)
        row_layout.setContentsMargins(0, 0, 0, 0)
        row_layout.setSpacing(4)

        title_label = StrongBodyLabel("Profil kecepatan")
        hint = CaptionLabel(
            "Semakin aman, proses makin lama tetapi 429 lebih jarang."
        )
        hint.setWordWrap(True)
        hint.setStyleSheet(f"color: {MUTED_TEXT_COLOR};")

        self.profile_combo = ComboBox()
        self._profile_keys = ["normal", "safe", "ultra"]
        self._profile_labels = {
            "normal": "Normal (cepat)",
            "safe": "Safe",
            "ultra": "Ultra",
        }
        for key in self._profile_keys:
            self.profile_combo.addItem(
                self._profile_labels.get(key, key)
            )

        row_layout.addWidget(title_label)
        row_layout.addWidget(hint)
        row_layout.addWidget(self.profile_combo, alignment=Qt.AlignLeft)
        badge_row = QWidget()
        badge_layout = QHBoxLayout(badge_row)
        badge_layout.setContentsMargins(0, 0, 0, 0)
        badge_layout.setSpacing(8)
        badge_label = CaptionLabel("Mode aktif")
        badge_label.setStyleSheet(f"color: {MUTED_TEXT_COLOR};")
        self.active_badge = CaptionLabel("-")
        self.active_badge.setStyleSheet(
            "padding: 2px 8px; border-radius: 10px;"
        )
        badge_layout.addWidget(badge_label)
        badge_layout.addWidget(self.active_badge)
        badge_layout.addStretch()
        row_layout.addWidget(badge_row)

        self.estimate_label = CaptionLabel("Estimasi waktu: -")
        self.estimate_label.setStyleSheet(f"color: {MUTED_TEXT_COLOR};")
        row_layout.addWidget(self.estimate_label)
        select_layout.addWidget(row)
        layout.addWidget(select_card)

        guide_card = CardWidget()
        guide_layout = QVBoxLayout(guide_card)
        guide_layout.setSpacing(4)
        guide_layout.addWidget(SubtitleLabel("Keterangan"))
        guide_items = [
            (
                "Normal (cepat)",
                "Pakai saat server relatif stabil. Estimasi waktu: ~1x.",
            ),
            (
                "Safe",
                "Pakai saat jam sibuk atau 429 sering muncul. Estimasi waktu: ~1.4–1.8x.",
            ),
            (
                "Ultra",
                "Pakai jika Safe belum cukup. Estimasi waktu: ~2–3x.",
            ),
        ]
        for title_text, desc_text in guide_items:
            title_label = StrongBodyLabel(title_text)
            desc_label = CaptionLabel(desc_text)
            desc_label.setWordWrap(True)
            desc_label.setStyleSheet(f"color: {MUTED_TEXT_COLOR};")
            guide_layout.addWidget(title_label)
            guide_layout.addWidget(desc_label)
        layout.addWidget(guide_card)

        note_card = CardWidget()
        note_layout = QVBoxLayout(note_card)
        note_layout.setSpacing(4)
        note_layout.addWidget(SubtitleLabel("Catatan Penting"))
        note_text = BodyLabel(
            "Mode ini tidak mengubah data Excel maupun hasil GC. "
            "Hanya mempengaruhi kecepatan dan jeda submit. "
            "Jika sering muncul pesan 'Something Went Wrong', "
            "pilih mode Safe atau Ultra."
        )
        note_text.setWordWrap(True)
        note_layout.addWidget(note_text)
        layout.addWidget(note_card)

        layout.addStretch()
        layout.addWidget(build_footer_label())

        self.profile_combo.currentIndexChanged.connect(
            self._apply_profile_selection
        )
        self._load_profile()

    def _load_profile(self):
        selected = load_rate_limit_profile()
        if selected in self._profile_keys:
            self.profile_combo.setCurrentIndex(
                self._profile_keys.index(selected)
            )
        self._update_detail(selected)

    def _apply_profile_selection(self):
        index = self.profile_combo.currentIndex()
        if index < 0 or index >= len(self._profile_keys):
            return
        key = self._profile_keys[index]
        save_rate_limit_profile(key)
        self._update_detail(key)

    def _update_detail(self, key):
        if key == "safe":
            estimate = "Estimasi waktu: ~1.4–1.8x (lebih lama dari Normal)"
            self._set_badge_style(
                "#FA8C16",
                "#FFF7E6",
                self._profile_labels.get("safe", "Safe"),
            )
        elif key == "ultra":
            estimate = "Estimasi waktu: ~2–3x (paling lama)"
            self._set_badge_style(
                "#CF1322",
                "#FFF1F0",
                self._profile_labels.get("ultra", "Ultra"),
            )
        else:
            estimate = "Estimasi waktu: ~1x (paling cepat)"
            self._set_badge_style(
                "#096DD9",
                "#E6F4FF",
                self._profile_labels.get("normal", "Normal"),
            )
        self.estimate_label.setText(estimate)

    def _set_badge_style(self, text_color, bg_color, label):
        self.active_badge.setText(label)
        self.active_badge.setStyleSheet(
            "padding: 2px 8px; border-radius: 10px; "
            f"background-color: {bg_color}; color: {text_color};"
        )


class PlaceholderPage(QWidget):
    def __init__(self, title, parent=None):
        super().__init__(parent)
        outer_layout = QVBoxLayout(self)
        outer_layout.setContentsMargins(0, 0, 0, 0)
        outer_layout.setSpacing(0)
        scroll, layout = build_scroll_area(self)
        outer_layout.addWidget(scroll)
        layout.addWidget(TitleLabel(title))
        layout.addWidget(BodyLabel("Halaman ini akan diisi di iterasi berikutnya."))
        layout.addStretch()
        layout.addWidget(build_footer_label())


class MainWindow(FluentWindow):
    def __init__(self, app):
        super().__init__()
        self._app = app

        self.setWindowTitle("DIRGC Automation")
        self.resize(1100, 720)

        self.home_page = HomePage(self)
        self.sso_page = SsoPage(self)
        self.run_page = RunPage(self.sso_page, self)
        self.update_page = RunPage(
            self.sso_page,
            self,
            update_mode_default=True,
            title_text="Update Data",
            subtitle_text="Perbarui data di DIRGC berdasarkan Excel.",
            run_label="Update",
            run_card_title="Update",
            confirm_title="Mulai update",
            confirm_message="Mulai update sekarang?",
            settings_key="options_update",
        )
        self.recap_page = RecapPage(self.sso_page, self)
        self.settings_page = SettingsPage(self._app, self)
        self.rate_limit_page = RateLimitPage(self)
        self.home_page.setObjectName("home_page")
        self.run_page.setObjectName("run_page")
        self.sso_page.setObjectName("sso_page")
        self.update_page.setObjectName("update_page")
        self.recap_page.setObjectName("recap_page")
        self.settings_page.setObjectName("settings_page")
        self.rate_limit_page.setObjectName("rate_limit_page")

        self.addSubInterface(self.home_page, FIF.HOME, "Beranda")
        self.addSubInterface(self.sso_page, FIF.PEOPLE, "Akun SSO")
        self.addSubInterface(self.run_page, FIF.PLAY, "Run")
        self.addSubInterface(self.update_page, FIF.EDIT, "Update")
        self.addSubInterface(self.recap_page, FIF.DOCUMENT, "Recap")
        self.addSubInterface(
            self.rate_limit_page, FIF.INFO, "Mode Stabilitas"
        )
        self.addSubInterface(
            self.settings_page,
            FIF.SETTING,
            "Settings",
            NavigationItemPosition.BOTTOM,
        )

    def closeEvent(self, event):
        if self.run_page:
            self.run_page._save_settings()
        if self.update_page:
            self.update_page._save_settings()
        if self.recap_page:
            self.recap_page._save_settings()
        super().closeEvent(event)


def main():
    app = QApplication(sys.argv)
    apply_app_font(app)
    setTheme(Theme.LIGHT)
    setThemeColor("#0078D4")

    window = MainWindow(app)
    apply_font_scale(app, load_font_scale())
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
