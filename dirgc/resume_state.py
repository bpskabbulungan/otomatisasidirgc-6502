import json
import os
import time

from .logging_utils import log_info


RESUME_STATE_PATH = os.path.join("config", "resume_state.json")


def save_resume_state(
    *,
    excel_file,
    next_row,
    reason="",
    run_log_path=None,
):
    payload = {
        "excel_file": excel_file or "",
        "next_row": int(next_row) if next_row is not None else None,
        "reason": reason or "",
        "run_log_path": str(run_log_path) if run_log_path else "",
        "saved_at": time.strftime("%Y-%m-%d %H:%M:%S"),
    }
    os.makedirs(os.path.dirname(RESUME_STATE_PATH), exist_ok=True)
    with open(RESUME_STATE_PATH, "w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2)
    log_info(
        "Resume state saved.",
        next_row=payload.get("next_row") or "-",
        path=RESUME_STATE_PATH,
    )
    return RESUME_STATE_PATH


def load_resume_state():
    try:
        with open(RESUME_STATE_PATH, "r", encoding="utf-8") as handle:
            return json.load(handle)
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return {}
