import os


TARGET_URL = "https://matchapro.web.bps.go.id/dirgc"
LOGIN_PATH = "/login"
MATCHAPRO_HOST = "matchapro.web.bps.go.id"
SSO_HOST = "sso.bps.go.id"
AUTO_LOGIN_RESULT_TIMEOUT_S = 15
DEFAULT_IDLE_TIMEOUT_MS = 300000
DEFAULT_WEB_TIMEOUT_S = 30

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
