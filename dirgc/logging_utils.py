import os
import sys
import time


FIELD_ORDER = (
    "row",
    "total",
    "row_excel",
    "idsbr",
    "nama_usaha",
    "alamat",
    "hasil_gc",
    "count",
    "status",
    "note",
    "error",
    "url",
    "path",
)
DIVIDER_LEN = 72
LEVEL_COLORS = {
    "INFO": "\x1b[32m",
    "WARN": "\x1b[33m",
    "ERROR": "\x1b[31m",
}
RESET_COLOR = "\x1b[0m"
_LOG_HANDLER = None


def normalize_log_value(value):
    if value is None:
        return ""
    text = str(value)
    if not text:
        return ""
    return " ".join(text.split())


def format_log_fields(fields):
    parts = []
    keys = []
    for key in FIELD_ORDER:
        if key in fields:
            keys.append(key)
    for key in sorted(fields.keys()):
        if key not in keys:
            keys.append(key)
    for key in keys:
        value = fields.get(key)
        text = normalize_log_value(value)
        if text == "":
            continue
        if " " in text or "=" in text:
            text = f'"{text}"'
        parts.append(f"{key}={text}")
    return " | ".join(parts)


def colorize_level(level):
    if not sys.stdout.isatty():
        return level
    if os.getenv("NO_COLOR"):
        return level
    color = LEVEL_COLORS.get(level)
    if not color:
        return level
    return f"{color}{level}{RESET_COLOR}"


def log(level, message, **fields):
    spacer = bool(fields.pop("_spacer", False))
    divider = bool(fields.pop("_divider", False))
    timestamp = time.strftime("%H:%M:%S")
    suffix = format_log_fields(fields)
    level_text = colorize_level(level)
    if _LOG_HANDLER:
        line = format_log_line(level, message, fields, timestamp)
        _LOG_HANDLER(line, spacer=spacer, divider=divider)
    if spacer:
        print()
    if divider:
        print("-" * DIVIDER_LEN)
    if suffix:
        print(f"[{timestamp}] {level_text}: {message} | {suffix}")
    else:
        print(f"[{timestamp}] {level_text}: {message}")


def format_log_line(level, message, fields, timestamp=None):
    if timestamp is None:
        timestamp = time.strftime("%H:%M:%S")
    suffix = format_log_fields(fields)
    if suffix:
        return f"[{timestamp}] {level}: {message} | {suffix}"
    return f"[{timestamp}] {level}: {message}"


def set_log_handler(handler):
    global _LOG_HANDLER
    _LOG_HANDLER = handler


def log_info(message, **fields):
    log("INFO", message, **fields)


def log_warn(message, **fields):
    log("WARN", message, **fields)


def log_error(message, **fields):
    log("ERROR", message, **fields)
