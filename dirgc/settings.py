import os


TARGET_URL = "https://matchapro.web.bps.go.id/dirgc"
LOGIN_PATH = "/login"
MATCHAPRO_HOST = "matchapro.web.bps.go.id"
SSO_HOST = "sso.bps.go.id"
AUTO_LOGIN_RESULT_TIMEOUT_S = 15
DEFAULT_IDLE_TIMEOUT_MS = 1800000
DEFAULT_WEB_TIMEOUT_S = 300
DEFAULT_RATE_LIMIT_PROFILE = "safe"
DEFAULT_SUBMIT_MODE = "ui"
SUBMIT_MODES = ("ui", "request")
DEFAULT_SESSION_REFRESH_EVERY = 0
DEFAULT_VPN_PREFIXES = ("10.",)
DEFAULT_RECAP_LENGTH = 5000
RECAP_LENGTH_WARN_THRESHOLD = 5000
RECAP_LENGTH_MAX = 10000

DEFAULT_CREDENTIALS_FILE = os.path.join("config", "credentials.json")
LEGACY_CREDENTIALS_FILE = "credentials.json"

DEFAULT_EXCEL_FILE = os.path.join("data", "Direktori_SBR_20260114.xlsx")
LEGACY_EXCEL_FILE = "Direktori_SBR_20260114.xlsx"

HASIL_GC_LABELS = {
    0: "Tidak Ditemukan",
    99: "Tidak Ditemukan",
    1: "Ditemukan",
    3: "Tutup",
    4: "Ganda",
}
VALID_HASIL_GC_CODES = set(HASIL_GC_LABELS.keys())
MAX_MATCH_LOGS = 3

BLOCK_UI_SELECTOR = ".blockUI.blockOverlay"

# Resource blocking for lighter pages (avoid blocking CSS/JS).
ENABLE_RESOURCE_BLOCKING = True
BLOCK_RESOURCE_TYPES = {"image", "media", "font"}
BLOCK_RESOURCE_DOMAINS = {"fonts.googleapis.com", "fonts.gstatic.com"}

# Write run log checkpoints periodically to reduce data loss on long runs.
RUN_LOG_CHECKPOINT_EVERY = 50

RATE_LIMIT_PROFILES = {
    "normal": {
        "label": "Normal",
        "submit_min_interval_s": 6.0,
        "submit_jitter_s": 1.2,
        "submit_penalty_initial_s": 20.0,
        "submit_penalty_max_s": 180.0,
        "submit_cooldown_after": 3,
        "submit_cooldown_s": 120.0,
        "submit_success_delay_s": 4.0,
        "filter_min_interval_s": 2.5,
        "filter_jitter_s": 0.6,
        "filter_penalty_initial_s": 8.0,
        "filter_penalty_max_s": 60.0,
    },
    "safe": {
        "label": "Safe",
        "submit_min_interval_s": 8.0,
        "submit_jitter_s": 1.5,
        "submit_penalty_initial_s": 30.0,
        "submit_penalty_max_s": 240.0,
        "submit_cooldown_after": 3,
        "submit_cooldown_s": 180.0,
        "submit_success_delay_s": 6.0,
        "filter_min_interval_s": 3.0,
        "filter_jitter_s": 0.8,
        "filter_penalty_initial_s": 10.0,
        "filter_penalty_max_s": 90.0,
    },
    "ultra": {
        "label": "Ultra",
        "submit_min_interval_s": 10.0,
        "submit_jitter_s": 2.0,
        "submit_penalty_initial_s": 45.0,
        "submit_penalty_max_s": 300.0,
        "submit_cooldown_after": 2,
        "submit_cooldown_s": 240.0,
        "submit_success_delay_s": 8.0,
        "filter_min_interval_s": 3.5,
        "filter_jitter_s": 1.0,
        "filter_penalty_initial_s": 12.0,
        "filter_penalty_max_s": 120.0,
    },
}
