import json
import math
import re
import time

from .browser import (
    SUBMIT_POST_SUCCESS_DELAY_S,
    SUBMIT_RATE_LIMITER,
    apply_filter,
    ensure_on_dirgc,
    get_server_cooldown_remaining,
    hasil_gc_select,
    is_visible,
    set_server_cooldown,
    wait_with_keepalive,
    wait_for_block_ui_clear,
)
from .excel import load_excel_rows
from .logging_utils import log_error, log_info, log_warn
from .matching import select_matching_card
from .resume_state import save_resume_state
from .run_logs import build_run_log_path, write_run_log, append_run_log_rows
from .settings import HASIL_GC_LABELS, RUN_LOG_CHECKPOINT_EVERY, TARGET_URL


RATE_LIMIT_POPUP_MARKERS = (
    "something wrong",
    "something went wrong",
    "too many request",
    "too many requests",
    "terlalu banyak permintaan",
    "429",
    "limit",
    "batas permintaan",
)


def detect_rate_limit_popup_text(page):
    try:
        return (
            page.evaluate(
                """
                (markers) => {
                  const lowered = markers.map((item) => (item || "").toLowerCase());
                  const popups = Array.from(document.querySelectorAll(".swal2-popup"));
                  for (const popup of popups) {
                    const style = window.getComputedStyle(popup);
                    const hidden = style.display === "none" || style.visibility === "hidden";
                    const rect = popup.getBoundingClientRect();
                    const invisible = rect.width === 0 && rect.height === 0;
                    if (hidden || invisible) continue;
                    const text = (popup.innerText || "").toLowerCase();
                    if (lowered.some((marker) => marker && text.includes(marker))) {
                      return popup.innerText || "";
                    }
                  }
                  return "";
                }
                """,
                list(RATE_LIMIT_POPUP_MARKERS),
            )
            or ""
        )
    except Exception:
        return ""


def _is_swal_overlay_visible(page):
    try:
        container = page.locator(".swal2-container")
        total = container.count()
    except Exception:
        return False
    if total == 0:
        return False
    check_total = min(total, 5)
    for idx in range(check_total):
        try:
            current = container.nth(idx)
            if current.is_visible():
                aria_hidden = (current.get_attribute("aria-hidden") or "").lower()
                if aria_hidden != "true":
                    return True
        except Exception:
            continue
    return False


def dismiss_swal_overlays(page, monitor, context="", timeout_s=12):
    deadline = time.monotonic() + timeout_s
    attempts = 0
    while True:
        if not _is_swal_overlay_visible(page):
            return True
        attempts += 1
        popup_locator = page.locator(".swal2-container").locator(".swal2-popup")
        acted = False
        for selector in (".swal2-confirm", ".swal2-close", ".swal2-cancel"):
            button = popup_locator.locator(selector)
            if button.count() > 0 and button.first.is_enabled():
                try:
                    monitor.bot_click(button.first)
                    acted = True
                    break
                except Exception:
                    continue
        if not acted:
            try:
                page.keyboard.press("Escape")
                acted = True
            except Exception:
                pass
        monitor.wait_for_condition(lambda: False, timeout_s=0.3)
        if time.monotonic() >= deadline:
            log_warn(
                "Popup SweetAlert tetap terbuka.",
                context=context or "-",
                attempts=attempts,
            )
            return False


def _is_bootstrap_modal_visible(page):
    try:
        modal = page.locator(".modal.show")
        total = modal.count()
    except Exception:
        return False
    if total == 0:
        return False
    check_total = min(total, 3)
    for idx in range(check_total):
        try:
            if modal.nth(idx).is_visible():
                return True
        except Exception:
            continue
    return False


def dismiss_bootstrap_modals(page, monitor, context="", timeout_s=10):
    deadline = time.monotonic() + timeout_s
    attempts = 0
    while True:
        if not _is_bootstrap_modal_visible(page):
            return True
        attempts += 1
        modal = page.locator(".modal.show")
        current = modal.last if modal.count() > 0 else modal
        acted = False
        for selector in ("[data-bs-dismiss='modal']", ".btn-close"):
            button = current.locator(selector)
            if button.count() > 0 and button.first.is_enabled():
                try:
                    monitor.bot_click(button.first)
                    acted = True
                    break
                except Exception:
                    continue
        if not acted:
            for label in ("Batal", "Tutup", "Close", "Cancel", "OK"):
                button = current.locator("button", has_text=label)
                if button.count() > 0 and button.first.is_enabled():
                    try:
                        monitor.bot_click(button.first)
                        acted = True
                        break
                    except Exception:
                        continue
        if not acted:
            try:
                page.keyboard.press("Escape")
                acted = True
            except Exception:
                pass
        monitor.wait_for_condition(lambda: False, timeout_s=0.3)
        if time.monotonic() >= deadline:
            log_warn(
                "Modal bootstrap tetap terbuka.",
                context=context or "-",
                attempts=attempts,
            )
            return False


def _parse_int(value):
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return None


_COORD_EMPTY_MARKERS = {
    "",
    "-",
    "--",
    "n/a",
    "na",
    "n.a.",
    "nan",
    "none",
    "null",
    "undefined",
}


def _normalize_coord_text(value):
    if value is None:
        return ""
    text = str(value).strip()
    if not text:
        return ""
    if text.lower() in _COORD_EMPTY_MARKERS:
        return ""
    return text


def _parse_coord_value(value, min_value, max_value):
    text = _normalize_coord_text(value)
    if not text:
        return None
    text = text.replace(",", ".")
    try:
        number = float(text)
    except (TypeError, ValueError):
        return None
    if math.isnan(number):
        return None
    if number < min_value or number > max_value:
        return None
    return number


def _is_indonesia_coord(lat_value, lon_value):
    lat = _parse_coord_value(lat_value, -90, 90)
    lon = _parse_coord_value(lon_value, -180, 180)
    if lat is None or lon is None:
        return False
    return -11.5 <= lat <= 6.5 and 95 <= lon <= 141.5


def _extract_cooldown_seconds(body_text, headers):
    seconds = 0
    reason = ""
    retry_after = None
    if headers:
        retry_after = headers.get("retry-after") or headers.get("Retry-After")
    retry_seconds = _parse_int(retry_after)
    if retry_seconds is not None:
        seconds = max(seconds, retry_seconds)

    text = (body_text or "").strip()
    if text:
        parsed = None
        if text.startswith("{") or text.startswith("["):
            try:
                parsed = json.loads(text)
            except json.JSONDecodeError:
                parsed = None
        if isinstance(parsed, dict):
            message = parsed.get("message") or ""
            reason = message or reason
            retry_after_body = _parse_int(parsed.get("retry_after"))
            if retry_after_body is not None:
                seconds = max(seconds, retry_after_body)
            text = message or text

        match = re.search(
            r"(?:tunggu|wait)\s+(\d+)\s*(menit|minute|detik|second|jam|hour)",
            text.lower(),
        )
        if match:
            value = _parse_int(match.group(1)) or 0
            unit = match.group(2)
            if unit in ("menit", "minute"):
                seconds = max(seconds, value * 60)
            elif unit in ("jam", "hour"):
                seconds = max(seconds, value * 3600)
            else:
                seconds = max(seconds, value)
    return seconds, reason


def _collect_form_payload(page):
    try:
        payload = page.evaluate(
            """
            () => {
              const button = document.querySelector("#save-tandai-usaha-btn");
              const form = button ? button.closest("form") : document.querySelector("form");
              if (!form) return null;
              const data = new FormData(form);
              const out = {};
              for (const [key, value] of data.entries()) {
                if (out[key] === undefined) {
                  out[key] = value;
                } else if (Array.isArray(out[key])) {
                  out[key].push(value);
                } else {
                  out[key] = [out[key], value];
                }
              }
              return out;
            }
            """
        )
    except Exception:
        return {}
    return payload if isinstance(payload, dict) else {}


def _extract_csrf_token(page):
    try:
        locator = page.locator('meta[name="csrf-token"]')
        if locator.count() > 0:
            return locator.first.get_attribute("content") or ""
    except Exception:
        return ""
    return ""


def _extract_gc_token(page):
    try:
        value = page.evaluate(
            "() => window.gcSubmitToken || window.gc_submit_token || null"
        )
        if value:
            return str(value)
    except Exception:
        pass
    for selector in (
        "input[name='gc_token']",
        "input#gc_token",
        "input[name*='gc_token']",
        "input[id*='gc_token']",
    ):
        try:
            locator = page.locator(selector)
            if locator.count() > 0:
                value = locator.first.get_attribute("value")
                if value:
                    return str(value)
                value = locator.first.input_value()
                if value:
                    return str(value)
        except Exception:
            continue
    try:
        content = page.content()
    except Exception:
        return ""
    match = re.search(
        r"gcSubmitToken\\s*=\\s*['\\\"]([^'\\\"]+)['\\\"]", content
    )
    if match:
        return match.group(1)
    return ""


def _extract_perusahaan_id(page):
    selectors = ["input[name='perusahaan_id']", "input#perusahaan_id"]
    for selector in selectors:
        try:
            locator = page.locator(selector)
            if locator.count() > 0:
                value = locator.first.get_attribute("value")
                if value:
                    return str(value)
                value = locator.first.input_value()
                if value:
                    return str(value)
        except Exception:
            continue
    return ""


def _set_payload_value(payload, key, value):
    if value is None:
        return
    payload[key] = str(value)


def _build_request_payload(
    page,
    *,
    update_hasil_gc,
    hasil_gc_value,
    update_lat,
    latitude_value,
    update_lon,
    longitude_value,
    update_nama,
    nama_value,
    update_alamat,
    alamat_value,
    perusahaan_id_override=None,
    gc_token_override=None,
):
    payload = _collect_form_payload(page)

    if update_hasil_gc and hasil_gc_value is not None:
        _set_payload_value(payload, "hasilgc", hasil_gc_value)
    if update_lat and latitude_value is not None:
        _set_payload_value(payload, "latitude", latitude_value)
    if update_lon and longitude_value is not None:
        _set_payload_value(payload, "longitude", longitude_value)
    if not update_lat:
        payload.pop("latitude", None)
    if not update_lon:
        payload.pop("longitude", None)

    if update_nama:
        if "edit_nama" not in payload:
            payload["edit_nama"] = "1" if str(nama_value or "").strip() else "0"
        if "nama_usaha" not in payload:
            payload["nama_usaha"] = str(nama_value or "")
    if update_alamat:
        if "edit_alamat" not in payload:
            payload["edit_alamat"] = "1" if str(alamat_value or "").strip() else "0"
        if "alamat_usaha" not in payload:
            payload["alamat_usaha"] = str(alamat_value or "")

    csrf_token = _extract_csrf_token(page)
    if csrf_token:
        payload["_token"] = csrf_token

    gc_token = gc_token_override or payload.get("gc_token") or _extract_gc_token(page)
    if gc_token:
        payload["gc_token"] = gc_token

    if not payload.get("perusahaan_id"):
        perusahaan_id = (
            perusahaan_id_override or _extract_perusahaan_id(page)
        )
        if perusahaan_id:
            payload["perusahaan_id"] = perusahaan_id

    return payload, gc_token


def _submit_via_request(page, payload, url, headers, timeout_ms=30000):
    try:
        response = page.request.post(
            url, form=payload, headers=headers, timeout=timeout_ms
        )
    except Exception as exc:
        return {
            "ok": False,
            "status": "error",
            "message": str(exc),
        }

    status_code = response.status
    response_text = ""
    response_json = None
    try:
        response_text = response.text() or ""
    except Exception:
        response_text = ""
    try:
        response_json = response.json()
    except Exception:
        response_json = None

    message = ""
    if isinstance(response_json, dict):
        message = response_json.get("message") or ""
    if not message:
        message = response_text.strip()

    if status_code == 200:
        if isinstance(response_json, dict) and response_json.get("status") == "error":
            return {
                "ok": False,
                "status": "error",
                "message": message or "Server mengembalikan status error.",
                "json": response_json,
            }
        return {
            "ok": True,
            "status": "success",
            "message": message or "Submit sukses (request).",
            "json": response_json,
        }

    if status_code == 429:
        resp_headers = {}
        try:
            resp_headers = response.headers or {}
        except Exception:
            resp_headers = {}
        cooldown_seconds, reason = _extract_cooldown_seconds(
            response_text, resp_headers
        )
        return {
            "ok": False,
            "status": "rate_limit",
            "message": message or "HTTP 429",
            "cooldown_s": cooldown_seconds,
            "reason": reason,
        }

    if status_code == 419 or "csrf token mismatch" in (message or "").lower():
        return {
            "ok": False,
            "status": "csrf",
            "message": message or "CSRF token mismatch.",
        }

    return {
        "ok": False,
        "status": "error",
        "message": message or f"HTTP {status_code}",
        "status_code": status_code,
    }


def _normalize_hasil_label(text):
    if text is None:
        return ""
    cleaned = " ".join(str(text).split())
    cleaned = re.sub(r"^\d+\s*[.\)\-:]\s*", "", cleaned)
    return cleaned.strip().lower()


def _select_hasil_gc_value(
    page,
    monitor,
    selector,
    code,
    *,
    allow_first_non_empty=False,
):
    if code is None:
        return False
    monitor.wait_for_condition(
        lambda: page.locator(selector).count() > 0, timeout_s=15
    )
    select_locator = page.locator(selector).first
    try:
        options = select_locator.locator("option").evaluate_all(
            "elements => elements.map(e => ({ value: e.value, text: e.textContent || '' }))"
        )
    except Exception:
        options = []

    code_value = str(code)
    if options and any(opt.get("value") == code_value for opt in options):
        monitor.bot_select_option(selector, value=code_value)
        return True
    if not options:
        value_locator = select_locator.locator(f'option[value="{code_value}"]')
        if value_locator.count() > 0:
            monitor.bot_select_option(selector, value=code_value)
            return True

    code_key = _parse_int(code)
    label = HASIL_GC_LABELS.get(code_key) if code_key is not None else None
    if label and options:
        target_norm = _normalize_hasil_label(label)
        matched_value = None
        for opt in options:
            opt_text = opt.get("text", "")
            if _normalize_hasil_label(opt_text) == target_norm:
                matched_value = opt.get("value")
                break
        if matched_value is None:
            for opt in options:
                opt_text = opt.get("text", "")
                if target_norm and target_norm in _normalize_hasil_label(opt_text):
                    matched_value = opt.get("value")
                    break
        if matched_value is not None:
            monitor.bot_select_option(selector, value=matched_value)
            return True
    if allow_first_non_empty and options:
        fallback_value = ""
        for opt in options:
            value = str(opt.get("value") or "").strip()
            if value:
                fallback_value = value
                break
        if fallback_value:
            monitor.bot_select_option(selector, value=fallback_value)
            return True
    return False


def _collect_report_form_validation(page):
    try:
        payload = page.evaluate(
            """
            () => {
              const ids = [
                "report_hasil_gc",
                "report_nama_usaha_gc",
                "report_alamat_usaha_gc",
                "report_latitude",
                "report_longitude",
              ];
              const messages = [];
              const fields = [];
              ids.forEach((id) => {
                const el = document.getElementById(id);
                if (!el) return;
                const raw = ("value" in el ? el.value : "") || "";
                const value = String(raw).trim();
                let msg = "";
                try {
                  if (typeof el.checkValidity === "function" && !el.checkValidity()) {
                    msg = (el.validationMessage || "").trim();
                  }
                } catch (e) {}
                const isInvalid = el.classList && el.classList.contains("is-invalid");
                if (!msg && isInvalid) {
                  msg = "Nilai tidak valid";
                }
                const required = !!el.required;
                if (!msg && required && !value) {
                  msg = "Wajib diisi";
                }
                if (msg) {
                  messages.push(`${id}: ${msg}`);
                }
                fields.push({
                  id,
                  required,
                  value,
                  invalid: !!msg,
                  validation_message: msg,
                });
              });
              return { messages, fields };
            }
            """
        )
    except Exception:
        return {"messages": [], "fields": []}
    if not isinstance(payload, dict):
        return {"messages": [], "fields": []}
    messages = payload.get("messages")
    fields = payload.get("fields")
    if not isinstance(messages, list):
        messages = []
    if not isinstance(fields, list):
        fields = []
    clean_messages = []
    for item in messages:
        text = str(item or "").strip()
        if text:
            clean_messages.append(text)
    return {"messages": clean_messages, "fields": fields}


def _handle_report_submit_result(page, monitor):
    success_markers = (
        "berhasil",
        "success",
        "submitted",
        "laporan berhasil",
        "data submitted",
    )
    error_markers = (
        "wajib",
        "harus diisi",
        "harus terisi",
        "invalid",
        "gagal",
        "error",
        "periksa kembali",
    )
    confirm_markers = (
        "konfirmasi laporan",
        "apakah anda yakin",
        "ya, laporkan",
    )

    def read_swal_message():
        try:
            payload = page.evaluate(
                """
                () => {
                  const popup = document.querySelector(".swal2-popup");
                  if (!popup) return null;
                  const isVisible = () => {
                    const style = window.getComputedStyle(popup);
                    if (!style) return false;
                    if (style.display === "none" || style.visibility === "hidden") return false;
                    const rect = popup.getBoundingClientRect();
                    return rect.width > 0 && rect.height > 0;
                  };
                  if (!isVisible()) return null;
                  const title = (popup.querySelector("#swal2-title")?.textContent || "").trim();
                  const body = (popup.querySelector("#swal2-html-container")?.textContent || "").trim();
                  const text = (popup.innerText || "").trim();
                  return {
                    title,
                    body,
                    text,
                  };
                }
                """
            )
        except Exception:
            return {"title": "", "body": "", "text": ""}
        if not isinstance(payload, dict):
            return {"title": "", "body": "", "text": ""}
        return {
            "title": str(payload.get("title") or "").strip(),
            "body": str(payload.get("body") or "").strip(),
            "text": str(payload.get("text") or "").strip(),
        }

    def swal_visible():
        popup = page.locator(".swal2-popup")
        return popup.count() > 0 and popup.first.is_visible()

    def wait_signal():
        if detect_rate_limit_popup_text(page).strip():
            return True
        if swal_visible():
            return True
        report_validation = _collect_report_form_validation(page)
        return bool(report_validation.get("messages"))

    for _ in range(5):
        monitor.wait_for_condition(wait_signal, timeout_s=12)

        rate_text = detect_rate_limit_popup_text(page).strip()
        if rate_text:
            return {
                "ok": False,
                "status": "error",
                "rate_limit": True,
                "message": rate_text,
            }

        if swal_visible():
            popup = page.locator(".swal2-popup").first
            swal_msg = read_swal_message()
            popup_title = " ".join((swal_msg.get("title") or "").split())
            popup_body = " ".join((swal_msg.get("body") or "").split())
            popup_text = " ".join((swal_msg.get("text") or "").split())
            combined_text = " ".join(
                part for part in (popup_title, popup_body, popup_text) if part
            )
            lowered = combined_text.lower()

            if any(marker in lowered for marker in confirm_markers):
                clicked = False
                try:
                    confirm_btn = popup.locator(".swal2-confirm", has_text="Ya")
                    if confirm_btn.count() > 0 and confirm_btn.first.is_visible():
                        monitor.bot_click(confirm_btn.first)
                        clicked = True
                except Exception:
                    clicked = False
                if not clicked:
                    try:
                        confirm_btn = popup.locator(".swal2-confirm")
                        if confirm_btn.count() > 0 and confirm_btn.first.is_visible():
                            monitor.bot_click(confirm_btn.first)
                            clicked = True
                    except Exception:
                        clicked = False
                if not clicked:
                    return {
                        "ok": False,
                        "status": "gagal",
                        "message": "Dialog konfirmasi laporan tanpa tombol konfirmasi.",
                    }
                monitor.wait_for_condition(lambda: False, timeout_s=0.4)
                continue

            is_success = any(marker in lowered for marker in success_markers) and not any(
                marker in lowered for marker in error_markers
            )
            try:
                confirm_btn = popup.locator(".swal2-confirm")
                if confirm_btn.count() > 0 and confirm_btn.first.is_visible():
                    monitor.bot_click(confirm_btn.first)
            except Exception:
                pass
            if is_success:
                return {
                    "ok": True,
                    "status": "success",
                    "message": popup_body or popup_text or "Laporan validasi terkirim.",
                }
            return {
                "ok": False,
                "status": "gagal",
                "message": popup_body or popup_text or "Validasi web menolak data laporan.",
            }

        report_validation = _collect_report_form_validation(page)
        messages = report_validation.get("messages") or []
        if messages:
            return {
                "ok": False,
                "status": "gagal",
                "message": "Validasi web: " + "; ".join(messages[:3]),
            }

        form_visible = False
        try:
            locator = page.locator("#report_hasil_gc")
            form_visible = locator.count() > 0 and locator.first.is_visible()
        except Exception:
            form_visible = False
        if form_visible:
            return {
                "ok": False,
                "status": "gagal",
                "message": "Validasi web menolak data laporan (tanpa pesan detail).",
            }

    return {
        "ok": False,
        "status": "error",
        "message": "Submit laporan validasi tidak mendapatkan respons web yang jelas.",
    }

def process_excel_rows(
    page,
    monitor,
    excel_file,
    use_saved_credentials,
    credentials,
    edit_nama_alamat=False,
    prefer_excel_coords=True,
    update_mode=False,
    validate_report_mode=False,
    update_fields=None,
    start_row=None,
    end_row=None,
    progress_callback=None,
    submit_mode="ui",
    session_refresh_every=0,
    stop_on_cooldown=False,
):
    mode_with_update_fields = bool(update_mode or validate_report_mode)
    if validate_report_mode:
        run_log_type = "validasi"
    elif update_mode:
        run_log_type = "update"
    else:
        run_log_type = "run"
    run_log_path = build_run_log_path(
        log_type=run_log_type
    )
    run_log_rows = []
    try:
        rows = load_excel_rows(excel_file)
    except Exception as exc:
        log_error("Failed to load Excel file.")
        run_log_rows.append(
            {
                "no": 0,
                "idsbr": "",
                "nama_usaha": "",
                "alamat": "",
                "keberadaanusaha_gc": "",
                "latitude": "",
                "latitude_source": "",
                "longitude": "",
                "longitude_source": "",
                "status": "error",
                "catatan": str(exc),
            }
        )
        write_run_log(run_log_rows, run_log_path)
        log_info("Run log saved.", path=str(run_log_path))
        return
    if not rows:
        log_warn("No rows found in Excel file.")
        write_run_log(run_log_rows, run_log_path)
        log_info("Run log saved.", path=str(run_log_path))
        return

    total_rows = len(rows)
    start_row = 1 if start_row is None else start_row
    end_row = total_rows if end_row is None else end_row
    if start_row < 1 or end_row < 1:
        log_error(
            "Start/end row must be >= 1.",
            start_row=start_row,
            end_row=end_row,
        )
        write_run_log(run_log_rows, run_log_path)
        log_info("Run log saved.", path=str(run_log_path))
        return
    if start_row > end_row:
        log_warn(
            "Start row is greater than end row; nothing to process.",
            start_row=start_row,
            end_row=end_row,
        )
        write_run_log(run_log_rows, run_log_path)
        log_info("Run log saved.", path=str(run_log_path))
        return
    if start_row > total_rows:
        log_warn(
            "Start row exceeds total rows; nothing to process.",
            start_row=start_row,
            total=total_rows,
        )
        write_run_log(run_log_rows, run_log_path)
        log_info("Run log saved.", path=str(run_log_path))
        return
    if end_row > total_rows:
        log_warn(
            "End row exceeds total rows; clamping.",
            end_row=end_row,
            total=total_rows,
        )
        end_row = total_rows

    rows = rows[start_row - 1 : end_row]
    selected_rows = len(rows)
    stats = {
        "total": selected_rows,
        "processed": 0,
        "skipped_no_results": 0,
        "skipped_gc": 0,
        "skipped_duplikat": 0,
        "skipped_no_tandai": 0,
        "hasil_gc_set": 0,
        "hasil_gc_skipped": 0,
    }
    log_info(
        "Start processing rows.",
        total=selected_rows,
        start_row=start_row,
        end_row=end_row,
    )
    if progress_callback:
        try:
            progress_callback(0, selected_rows, 0)
        except Exception:
            pass

    stop_reason = None
    idle_reason = None
    cooldown_reason = None
    checkpoint_every = int(RUN_LOG_CHECKPOINT_EVERY or 0)
    appended_rows = 0
    session_submit_count = 0
    session_refresh_pending = False
    request_gc_token = None
    submit_mode = (submit_mode or "ui").strip().lower()
    if submit_mode not in ("ui", "request"):
        submit_mode = "ui"
    if validate_report_mode and submit_mode == "request":
        log_warn(
            "Submit via request tidak didukung di mode Validasi GC; fallback ke UI."
        )
        submit_mode = "ui"
    try:
        session_refresh_every = int(session_refresh_every or 0)
    except (TypeError, ValueError):
        session_refresh_every = 0

    def checkpoint_run_log(force=False):
        nonlocal appended_rows
        if not run_log_rows:
            return
        if checkpoint_every <= 0 and not force:
            return
        if (
            checkpoint_every > 0
            and not force
            and len(run_log_rows) < checkpoint_every
        ):
            return
        append_run_log_rows(run_log_rows, run_log_path)
        appended_rows += len(run_log_rows)
        run_log_rows.clear()
        if not force:
            log_info(
                "Run log checkpoint saved.",
                rows=appended_rows,
                path=str(run_log_path),
            )

    for offset, row in enumerate(rows):
        batch_index = offset + 1
        excel_row = start_row + offset
        stats["processed"] += 1
        status = None
        note = ""
        extra_notes = []
        stop_processing = False

        idsbr = row["idsbr"]
        nama_usaha = row["nama_usaha"]
        alamat = row["alamat"]
        latitude = row["latitude"]
        longitude = row["longitude"]
        hasil_gc = row["hasil_gc"]
        latitude_value = latitude or ""
        longitude_value = longitude or ""
        latitude_source = "unknown"
        longitude_source = "unknown"
        latitude_before = ""
        latitude_after = ""
        longitude_before = ""
        longitude_after = ""
        nama_before = ""
        nama_after = ""
        alamat_before = ""
        alamat_after = ""
        hasil_gc_before = ""
        hasil_gc_after = ""
        gc_owner_username = ""

        update_fields_set = None
        if mode_with_update_fields:
            if isinstance(update_fields, (list, tuple, set)):
                update_fields_set = {
                    str(item).strip().lower()
                    for item in update_fields
                    if str(item).strip()
                }
            elif update_fields is not None:
                update_fields_set = set()
        update_hasil_gc = (
            not mode_with_update_fields
            or update_fields_set is None
            or "hasil_gc" in update_fields_set
        )
        update_nama = (
            not mode_with_update_fields
            or update_fields_set is None
            or "nama_usaha" in update_fields_set
        )
        update_alamat = (
            not mode_with_update_fields
            or update_fields_set is None
            or "alamat" in update_fields_set
        )
        update_lat = (
            not mode_with_update_fields
            or update_fields_set is None
            or "latitude" in update_fields_set
            or "koordinat" in update_fields_set
        )
        update_lon = (
            not mode_with_update_fields
            or update_fields_set is None
            or "longitude" in update_fields_set
            or "koordinat" in update_fields_set
        )

        log_info(
            "Processing row.",
            _spacer=True,
            _divider=True,
            row=batch_index,
            total=selected_rows,
            row_excel=excel_row,
            idsbr=idsbr or "-",
        )
        update_field_labels = []
        if update_hasil_gc:
            update_field_labels.append("hasil_gc")
        if update_nama:
            update_field_labels.append("nama_usaha")
        if update_alamat:
            update_field_labels.append("alamat")
        if update_lat or update_lon:
            if update_lat and update_lon:
                update_field_labels.append("koordinat")
            else:
                if update_lat:
                    update_field_labels.append("latitude")
                if update_lon:
                    update_field_labels.append("longitude")
        update_fields_label = ", ".join(update_field_labels) or "-"
        log_info(
            "Field proses.",
            idsbr=idsbr or "-",
            fields=update_fields_label,
            mode=(
                "validasi_gc"
                if validate_report_mode
                else ("update" if update_mode else "run")
            ),
        )

        def build_success_note(server_message):
            if validate_report_mode:
                return f"Validasi GC sukses: {update_fields_label}."
            if update_mode or edit_nama_alamat:
                return f"Update sukses: {update_fields_label}."
            message = (server_message or "").strip()
            return message or "Submit sukses"

        try:
            monitor.idle_check()
            if update_lat or update_lon:
                lat_text = (latitude or "").strip()
                lon_text = (longitude or "").strip()
                if (
                    lat_text
                    and lon_text
                    and not _is_indonesia_coord(lat_text, lon_text)
                ):
                    log_warn(
                        "Koordinat Excel invalid; lewati baris.",
                        idsbr=idsbr or "-",
                        latitude=lat_text,
                        longitude=lon_text,
                    )
                    status = "error"
                    note = "Koordinat Excel invalid; baris dilewati."
                    continue
            did_refresh = False
            if session_refresh_pending or (
                session_refresh_every and session_submit_count >= session_refresh_every
            ):
                reason = (
                    "pending"
                    if session_refresh_pending
                    else f"interval={session_refresh_every}"
                )
                log_info("Auto session refresh.", reason=reason)
                try:
                    page.reload()
                except Exception:
                    pass
                ensure_on_dirgc(
                    page,
                    monitor=monitor,
                    use_saved_credentials=use_saved_credentials,
                    credentials=credentials,
                )
                session_submit_count = 0
                session_refresh_pending = False
                request_gc_token = None
                did_refresh = True
            if not did_refresh:
                ensure_on_dirgc(
                    page,
                    monitor=monitor,
                    use_saved_credentials=use_saved_credentials,
                    credentials=credentials,
                )
            remaining, reason = get_server_cooldown_remaining()
            if remaining > 0:
                resume_at = time.strftime(
                    "%H:%M:%S", time.localtime(time.time() + remaining)
                )
                if stop_on_cooldown:
                    log_warn(
                        "Cooldown aktif; menghentikan proses.",
                        wait_s=round(remaining, 1),
                        resume_at=resume_at,
                        reason=reason or "-",
                        context="pre-filter",
                    )
                    save_resume_state(
                        excel_file=excel_file,
                        next_row=excel_row,
                        reason=reason,
                        run_log_path=run_log_path,
                    )
                    status = "stopped"
                    note = (
                        f"Cooldown aktif; lanjutkan dari baris {excel_row}."
                    )
                    cooldown_reason = "Cooldown active; stopped for resume."
                    stop_processing = True
                    continue
                log_warn(
                    "Server meminta jeda sebelum lanjut.",
                    wait_s=round(remaining, 1),
                    resume_at=resume_at,
                    reason=reason or "-",
                    context="pre-filter",
                )
                wait_with_keepalive(
                    monitor, remaining, log_context="pre-filter"
                )
            if mode_with_update_fields:
                missing_fields = []
                lat_text = (latitude or "").strip()
                lon_text = (longitude or "").strip()
                if update_hasil_gc and hasil_gc is None:
                    missing_fields.append("hasil_gc")
                if update_nama and not (nama_usaha or "").strip():
                    missing_fields.append("nama_usaha")
                if update_alamat and not (alamat or "").strip():
                    missing_fields.append("alamat")
                if update_lat and update_lon:
                    if not lat_text and not lon_text:
                        pass
                    elif (lat_text and not lon_text) or (lon_text and not lat_text):
                        partial_label = (
                            "latitude"
                            if lat_text and not lon_text
                            else "longitude"
                        )
                        log_warn(
                            "Update koordinat parsial: hanya satu nilai terisi.",
                            idsbr=idsbr or "-",
                            filled=partial_label,
                        )
                        extra_notes.append(
                            f"Koordinat parsial (hanya {partial_label})"
                        )
                elif update_lat and not lat_text:
                    missing_fields.append("latitude")
                elif update_lon and not lon_text:
                    missing_fields.append("longitude")

                if missing_fields:
                    missing_label = ", ".join(missing_fields)
                    log_warn(
                        (
                            "Validasi ditolak: field kosong di Excel."
                            if validate_report_mode
                            else "Update ditolak: field kosong di Excel."
                        ),
                        idsbr=idsbr or "-",
                        fields=missing_label,
                    )
                    status = "gagal"
                    note = (
                        f"Validasi ditolak: field kosong ({missing_label})"
                        if validate_report_mode
                        else f"Update ditolak: field kosong ({missing_label})"
                    )
                    continue

            log_info(
                "Applying filter.",
                idsbr=idsbr or "-",
                nama_usaha=nama_usaha or "-",
                alamat=alamat or "-",
            )
            result_count = apply_filter(page, monitor, idsbr, nama_usaha, alamat)
            log_info("Filter results.", count=result_count)

            if not dismiss_swal_overlays(
                page, monitor, context="pre-select card", timeout_s=10
            ):
                log_warn(
                    "Popup tidak tertutup sebelum pilih kartu; skipping.",
                    idsbr=idsbr or "-",
                )
                status = "error"
                note = "Popup SweetAlert tertahan sebelum pilih kartu"
                monitor.bot_goto(TARGET_URL)
                continue
            if not dismiss_bootstrap_modals(
                page, monitor, context="pre-select card", timeout_s=8
            ):
                log_warn(
                    "Modal dialog tetap terbuka sebelum pilih kartu; skipping.",
                    idsbr=idsbr or "-",
                )
                status = "error"
                note = "Modal dialog tertahan sebelum pilih kartu"
                monitor.bot_goto(TARGET_URL)
                continue

            selection = select_matching_card(
                page, monitor, idsbr, nama_usaha, alamat
            )
            if not selection:
                log_warn("No results found; skipping.", idsbr=idsbr or "-")
                stats["skipped_no_results"] += 1
                status = "gagal"
                note = "No results found"
                continue

            header_locator, card_scope = selection
            if card_scope.count() == 0:
                card_scope = page.locator("body")
            card_data_id = ""
            try:
                card_data_id = card_scope.get_attribute("data-id") or ""
            except Exception:
                card_data_id = ""
            if not card_data_id:
                try:
                    card_data_id = (
                        header_locator.get_attribute("data-id") or ""
                    )
                except Exception:
                    card_data_id = ""
            if not card_data_id:
                try:
                    candidate = pick_visible(
                        card_scope.locator(
                            "[data-id].btn-gc-edit, [data-id].btn-gc-report, [data-id].btn-tandai, [data-id].btn-detail-link, [data-id].btn-gc-map, [data-id]"
                        )
                    )
                    if candidate is not None:
                        card_data_id = candidate.get_attribute("data-id") or ""
                except Exception:
                    card_data_id = ""

            def is_card_expanded():
                try:
                    classes = (card_scope.get_attribute("class") or "")
                    if "expanded" in classes.split():
                        return True
                except Exception:
                    pass
                try:
                    details = card_scope.locator(".usaha-card-details")
                    if details.count() > 0 and details.first.is_visible():
                        return True
                except Exception:
                    pass
                return False

            def pick_visible(locator, max_checks=6):
                try:
                    total = locator.count()
                except Exception:
                    return None
                for idx in range(min(total, max_checks)):
                    try:
                        current = locator.nth(idx)
                        if current.is_visible():
                            return current
                    except Exception:
                        continue
                return None

            def locate_visible(selector, scope=None):
                if scope is not None:
                    found = pick_visible(scope.locator(selector))
                    if found is not None:
                        return found
                return pick_visible(page.locator(selector))

            def extract_username_from_text(text):
                normalized = " ".join(str(text or "").split())
                if not normalized:
                    return ""
                patterns = (
                    r"(?i)\b(?:gc\s*username|username\s*gc|user\s*gc|gc\s*oleh|digc\s*oleh|oleh\s*gc)\b\s*[:\-]?\s*([A-Za-z0-9._@-]{2,64})",
                    r"(?i)\b(?:gc\s*by|by\s*gc)\b\s*[:\-]?\s*([A-Za-z0-9._@-]{2,64})",
                )
                for pattern in patterns:
                    match = re.search(pattern, normalized)
                    if match:
                        return (match.group(1) or "").strip(".,;:|()[]{}")
                if re.fullmatch(r"[A-Za-z0-9._@-]{2,64}", normalized):
                    return normalized
                return ""

            def find_gc_owner_username():
                attr_candidates = (
                    "data-gc-username",
                    "data-username-gc",
                    "data-gc-user",
                    "data-username",
                )
                for locator in (card_scope, header_locator):
                    for attr_name in attr_candidates:
                        try:
                            raw = locator.get_attribute(attr_name) or ""
                        except Exception:
                            raw = ""
                        username = extract_username_from_text(raw)
                        if username:
                            return username

                selector_candidates = (
                    "[data-gc-username]",
                    "[data-username-gc]",
                    ".gc-username",
                    ".username-gc",
                    ".gc-user",
                    ".user-gc",
                    ".gc-by",
                )
                for selector in selector_candidates:
                    try:
                        locator = card_scope.locator(selector)
                        total = locator.count()
                    except Exception:
                        continue
                    for idx in range(min(total, 6)):
                        try:
                            node = locator.nth(idx)
                            if not node.is_visible():
                                continue
                            raw = node.inner_text() or ""
                        except Exception:
                            continue
                        username = extract_username_from_text(raw)
                        if username:
                            return username

                try:
                    card_text = card_scope.inner_text() or ""
                except Exception:
                    card_text = ""
                return extract_username_from_text(card_text)

            try:
                header_locator.scroll_into_view_if_needed()
            except Exception:
                pass
            if not is_card_expanded():
                try:
                    monitor.bot_click(header_locator)
                except Exception as exc:
                    if _is_swal_overlay_visible(page):
                        dismissed = dismiss_swal_overlays(
                            page,
                            monitor,
                            context="select card retry",
                            timeout_s=8,
                        )
                        if dismissed:
                            monitor.bot_click(header_locator)
                        else:
                            raise
                    elif _is_bootstrap_modal_visible(page):
                        dismissed = dismiss_bootstrap_modals(
                            page,
                            monitor,
                            context="select card retry",
                            timeout_s=8,
                        )
                        if dismissed:
                            monitor.bot_click(header_locator)
                        else:
                            raise
                    else:
                        raise
                monitor.wait_for_condition(is_card_expanded, timeout_s=6)

            if update_mode:
                gc_owner_username = find_gc_owner_username() or ""

            if not (update_mode or validate_report_mode):
                if (
                    card_scope.locator(
                        ".gc-badge", has_text="Sudah GC"
                    ).count()
                    > 0
                ):
                    log_info("Skipped: Sudah GC.", idsbr=idsbr or "-")
                    stats["skipped_gc"] += 1
                    status = "skipped"
                    note = "Sudah GC"
                    continue

            if card_scope.locator(
                ".usaha-status.tidak-aktif", has_text="Duplikat"
            ).count() > 0:
                log_info("Skipped: Duplikat.", idsbr=idsbr or "-")
                stats["skipped_duplikat"] += 1
                status = "skipped"
                note = "Duplikat"
                continue

            if validate_report_mode:
                action_selector = ".btn-gc-report"
                action_label = "Laporkan Hasil GC - Tidak Valid"
                hasil_selector = "#report_hasil_gc"
                lat_selector = "#report_latitude"
                lon_selector = "#report_longitude"
                nama_toggle_selector = "#report_toggle_edit_nama"
                nama_input_selector = "#report_nama_usaha_gc"
                alamat_toggle_selector = "#report_toggle_edit_alamat"
                alamat_input_selector = "#report_alamat_usaha_gc"
                submit_selector = "#submit-report-gc-btn"
            elif update_mode:
                action_selector = ".btn-gc-edit"
                action_label = "Edit Hasil"
                hasil_selector = "#tt_hasil_gc"
                lat_selector = "#tt_latitude_cek_user"
                lon_selector = "#tt_longitude_cek_user"
                nama_toggle_selector = "#toggle_edit_nama"
                nama_input_selector = "#tt_nama_usaha_gc"
                alamat_toggle_selector = "#toggle_edit_alamat"
                alamat_input_selector = "#tt_alamat_usaha_gc"
                submit_selector = "#save-tandai-usaha-btn"
            else:
                action_selector = ".btn-tandai"
                action_label = "Tandai"
                hasil_selector = "#tt_hasil_gc"
                lat_selector = "#tt_latitude_cek_user"
                lon_selector = "#tt_longitude_cek_user"
                nama_toggle_selector = "#toggle_edit_nama"
                nama_input_selector = "#tt_nama_usaha_gc"
                alamat_toggle_selector = "#toggle_edit_alamat"
                alamat_input_selector = "#tt_alamat_usaha_gc"
                submit_selector = "#save-tandai-usaha-btn"
            action_locator = card_scope.locator(action_selector)
            if action_locator.count() == 0:
                action_locator = page.locator(action_selector)
            action_button = pick_visible(action_locator)
            if action_button is None:
                if validate_report_mode:
                    missing_note = (
                        "Tombol Laporkan Hasil GC - Tidak Valid tidak ditemukan."
                    )
                elif update_mode:
                    owner_fragment = (
                        f' gc_username="{gc_owner_username}".'
                        if gc_owner_username
                        else ""
                    )
                    missing_note = (
                        "Tombol Edit Hasil tidak ditemukan "
                        "(kemungkinan belum GC atau GC milik user lain;"
                        f"{owner_fragment} gunakan menu Run/Validasi GC)."
                    )
                else:
                    missing_note = "Tombol Tandai tidak ditemukan."
                if update_mode:
                    log_info(
                        f"Tombol {action_label} tidak ditemukan; lewati.",
                        idsbr=idsbr or "-",
                        gc_username=gc_owner_username or "-",
                    )
                else:
                    log_warn(
                        f"Tombol {action_label} tidak ditemukan; skipping.",
                        idsbr=idsbr or "-",
                    )
                stats["skipped_no_tandai"] += 1
                status = "skipped" if update_mode else "gagal"
                note = missing_note
                continue

            wait_for_block_ui_clear(page, monitor, timeout_s=15)
            try:
                action_button.scroll_into_view_if_needed()
            except Exception:
                pass
            try:
                monitor.bot_click(action_button)
            except Exception as exc:
                log_warn(
                    f"Tombol {action_label} gagal diklik; skipping.",
                    idsbr=idsbr or "-",
                    error=str(exc),
                )
                stats["skipped_no_tandai"] += 1
                status = "gagal"
                note = f"Tombol {action_label} gagal diklik"
                continue
            form_ready = monitor.wait_for_condition(
                lambda: locate_visible(hasil_selector) is not None,
                timeout_s=30,
            )
            if not form_ready:
                log_warn(
                    "Form tidak muncul setelah klik aksi; skipping.",
                    idsbr=idsbr or "-",
                )
                stats["skipped_no_tandai"] += 1
                status = "gagal"
                note = (
                    "Form Validasi GC tidak muncul"
                    if validate_report_mode
                    else "Form Hasil GC tidak muncul"
                )
                continue

            form_scope = page
            modal_visible = pick_visible(page.locator(".modal.show"))
            if modal_visible is not None:
                form_scope = modal_visible

            if update_hasil_gc:
                select_locator = (
                    locate_visible(hasil_selector, form_scope)
                    or page.locator(hasil_selector).first
                )
                try:
                    hasil_gc_before = select_locator.input_value() or ""
                except Exception:
                    hasil_gc_before = ""
                if (
                    _select_hasil_gc_value(
                        page,
                        monitor,
                        hasil_selector,
                        hasil_gc,
                        allow_first_non_empty=validate_report_mode,
                    )
                    if validate_report_mode
                    else hasil_gc_select(page, monitor, hasil_gc)
                ):
                    hasil_gc_after = str(hasil_gc) if hasil_gc is not None else ""
                    log_info(
                        "Hasil GC set.", hasil_gc=hasil_gc, idsbr=idsbr or "-"
                    )
                    stats["hasil_gc_set"] += 1
                else:
                    hasil_gc_after = hasil_gc_before
                    log_warn(
                        "Hasil GC tidak diisi (kode tidak valid/kosong).",
                        idsbr=idsbr or "-",
                    )
                    stats["hasil_gc_skipped"] += 1
                    status = "gagal"
                    note = (
                        "Hasil GC validasi tidak valid/kosong"
                        if validate_report_mode
                        else "Hasil GC tidak valid/kosong"
                    )
            else:
                select_locator = (
                    locate_visible(hasil_selector, form_scope)
                    or page.locator(hasil_selector).first
                )
                try:
                    hasil_gc_before = select_locator.input_value() or ""
                except Exception:
                    hasil_gc_before = ""
                hasil_gc_after = hasil_gc_before

            def safe_fill(
                selector, value, field_name, allow_overwrite, scope=None
            ):
                locator = locate_visible(selector, scope)
                if locator is None:
                    if update_mode:
                        log_info(
                            "Field update tidak ditemukan; lewati (indikasi hak edit terbatas).",
                            idsbr=idsbr or "-",
                            field=field_name,
                        )
                    else:
                        log_warn(
                            "Field tidak ditemukan; lewati.",
                            idsbr=idsbr or "-",
                            field=field_name,
                        )
                    return "", "missing", "", ""
                def read_raw():
                    try:
                        return locator.input_value() or ""
                    except Exception:
                        return ""

                def read_value():
                    try:
                        return _normalize_coord_text(locator.input_value())
                    except Exception:
                        return ""

                raw_value = read_raw()
                current_value = _normalize_coord_text(raw_value)
                if raw_value and not current_value:
                    monitor.mark_activity("bot")
                    try:
                        locator.fill("")
                    except Exception:
                        pass
                    current_value = ""
                desired_value = _normalize_coord_text(value)

                def resolve_source(actual_value, filled):
                    if not actual_value:
                        return "empty"
                    if desired_value and actual_value == desired_value:
                        if current_value == desired_value and not filled:
                            return "excel_same_as_web"
                        if filled:
                            return "excel"
                    return "web"

                if allow_overwrite:
                    if desired_value:
                        filled = False
                        if current_value != desired_value:
                            monitor.mark_activity("bot")
                            try:
                                locator.fill(desired_value)
                                filled = True
                            except Exception:
                                filled = False
                        actual_value = read_value()
                        if actual_value:
                            source = resolve_source(actual_value, filled)
                            return actual_value, source, current_value, actual_value
                        if current_value:
                            source = resolve_source(current_value, False)
                            return (
                                current_value,
                                source,
                                current_value,
                                current_value,
                            )
                        return "", "empty", current_value, ""

                if current_value:
                    source = resolve_source(current_value, False)
                    return current_value, source, current_value, current_value
                if not desired_value:
                    return "", "empty", current_value, ""
                monitor.mark_activity("bot")
                filled = False
                try:
                    locator.fill(desired_value)
                    filled = True
                except Exception:
                    filled = False
                actual_value = read_value()
                if actual_value:
                    source = resolve_source(actual_value, filled)
                    return actual_value, source, current_value, actual_value
                return "", "empty", current_value, ""

            lat_value = latitude if update_lat else None
            lon_value = longitude if update_lon else None
            allow_overwrite_lat = prefer_excel_coords if update_lat else False
            allow_overwrite_lon = prefer_excel_coords if update_lon else False
            latitude_value, latitude_source, latitude_before, latitude_after = (
                safe_fill(
                    lat_selector,
                    lat_value,
                    "latitude",
                    allow_overwrite_lat,
                    form_scope,
                )
            )
            longitude_value, longitude_source, longitude_before, longitude_after = (
                safe_fill(
                    lon_selector,
                    lon_value,
                    "longitude",
                    allow_overwrite_lon,
                    form_scope,
                )
            )

            def ensure_edit_field(
                toggle_selector, input_selector, value, field_name, scope=None
            ):
                desired_value = str(value).strip() if value else ""
                toggle = locate_visible(toggle_selector, scope)
                if toggle is None:
                    if update_mode:
                        log_info(
                            "Toggle edit tidak ditemukan; lewati (indikasi hak edit terbatas).",
                            idsbr=idsbr or "-",
                            field=field_name,
                        )
                    else:
                        log_warn(
                            "Toggle edit tidak ditemukan; lewati.",
                            idsbr=idsbr or "-",
                            field=field_name,
                        )
                    return "", ""
                try:
                    toggle_checked = toggle.is_checked()
                except Exception:
                    toggle_checked = False
                if not toggle_checked:
                    try:
                        monitor.bot_click(toggle)
                    except Exception as exc:
                        log_warn(
                            "Toggle edit gagal diklik; lewati.",
                            idsbr=idsbr or "-",
                            field=field_name,
                            error=str(exc),
                        )
                        return "", ""

                input_locator = locate_visible(input_selector, scope)
                if input_locator is None:
                    if update_mode:
                        log_info(
                            "Field edit tidak ditemukan; lewati (indikasi hak edit terbatas).",
                            idsbr=idsbr or "-",
                            field=field_name,
                        )
                    else:
                        log_warn(
                            "Field edit tidak ditemukan; lewati.",
                            idsbr=idsbr or "-",
                            field=field_name,
                        )
                    return "", ""
                if not monitor.wait_for_condition(
                    lambda: input_locator.is_editable(),
                    timeout_s=5,
                ):
                    if update_mode:
                        log_info(
                            "Field edit tidak bisa diedit; lewati (indikasi hak edit terbatas).",
                            idsbr=idsbr or "-",
                            field=field_name,
                        )
                    else:
                        log_warn(
                            "Field edit tidak bisa diedit; lewati.",
                            idsbr=idsbr or "-",
                            field=field_name,
                        )
                    return "", ""
                try:
                    current_value = input_locator.input_value()
                except Exception:
                    current_value = ""
                current_value = (current_value or "").strip()
                if not desired_value:
                    return current_value, current_value
                if current_value == desired_value:
                    return current_value, desired_value
                monitor.mark_activity("bot")
                input_locator.fill(desired_value)
                return current_value, desired_value

            if update_mode or validate_report_mode:
                if update_nama:
                    nama_before, nama_after = ensure_edit_field(
                        nama_toggle_selector,
                        nama_input_selector,
                        nama_usaha,
                        "nama_usaha",
                        form_scope,
                    )
                if update_alamat:
                    alamat_before, alamat_after = ensure_edit_field(
                        alamat_toggle_selector,
                        alamat_input_selector,
                        alamat,
                        "alamat",
                        form_scope,
                    )
            elif edit_nama_alamat:
                ensure_edit_field(
                    "#toggle_edit_nama",
                    "#tt_nama_usaha_gc",
                    nama_usaha,
                    "nama_usaha",
                    form_scope,
                )
                ensure_edit_field(
                    "#toggle_edit_alamat",
                    "#tt_alamat_usaha_gc",
                    alamat,
                    "alamat",
                    form_scope,
                )

            if (
                (latitude_value or longitude_value)
                and not _is_indonesia_coord(latitude_value, longitude_value)
                and not (update_lat or update_lon)
            ):
                log_warn(
                    "Koordinat existing invalid; lewati baris.",
                    idsbr=idsbr or "-",
                    latitude=latitude_value or "-",
                    longitude=longitude_value or "-",
                )
                status = "error"
                note = "Koordinat existing invalid; baris dilewati."
                monitor.bot_goto(TARGET_URL)
                continue

            if status == "gagal" and note in {
                "Hasil GC tidak valid/kosong",
                "Hasil GC validasi tidak valid/kosong",
            }:
                monitor.bot_goto(TARGET_URL)
                continue
            if validate_report_mode and not _normalize_coord_text(hasil_gc_after):
                log_warn(
                    "Validasi ditolak sebelum submit: opsi keberadaan usaha belum terisi.",
                    idsbr=idsbr or "-",
                )
                status = "gagal"
                note = "Opsi keberadaan usaha hasil gc harus terisi!"
                monitor.bot_goto(TARGET_URL)
                continue

            submit_locator = locate_visible(submit_selector, form_scope)
            if submit_locator is None:
                submit_locator = locate_visible(submit_selector)
            if submit_locator is None:
                if update_mode:
                    log_info(
                        "Tombol submit update tidak terlihat; lewati (kemungkinan GC milik user lain).",
                        idsbr=idsbr or "-",
                        gc_username=gc_owner_username or "-",
                    )
                    status = "skipped"
                    owner_note = (
                        f' (gc_username: "{gc_owner_username}")'
                        if gc_owner_username
                        else ""
                    )
                    note = (
                        "Update dilewati: hasil GC milik user lain"
                        f"{owner_note}; gunakan menu Validasi GC."
                    )
                else:
                    log_warn(
                        "Tombol submit tidak ditemukan; skipping.",
                        idsbr=idsbr or "-",
                    )
                    status = "gagal"
                    note = "Tombol submit tidak ditemukan"
                monitor.bot_goto(TARGET_URL)
                continue

            remaining, reason = get_server_cooldown_remaining()
            if remaining > 0:
                resume_at = time.strftime(
                    "%H:%M:%S", time.localtime(time.time() + remaining)
                )
                if stop_on_cooldown:
                    log_warn(
                        "Cooldown aktif; menghentikan proses.",
                        wait_s=round(remaining, 1),
                        resume_at=resume_at,
                        reason=reason or "-",
                        context="pre-submit",
                    )
                    save_resume_state(
                        excel_file=excel_file,
                        next_row=excel_row,
                        reason=reason,
                        run_log_path=run_log_path,
                    )
                    status = "stopped"
                    note = (
                        f"Cooldown aktif; lanjutkan dari baris {excel_row}."
                    )
                    cooldown_reason = "Cooldown active; stopped for resume."
                    stop_processing = True
                    continue
                log_warn(
                    "Server meminta jeda sebelum submit.",
                    wait_s=round(remaining, 1),
                    resume_at=resume_at,
                    reason=reason or "-",
                    context="pre-submit",
                )
                wait_with_keepalive(
                    monitor, remaining, log_context="pre-submit"
                )
            SUBMIT_RATE_LIMITER.wait_for_slot(monitor)
            wait_for_block_ui_clear(page, monitor, timeout_s=15)
            try:
                submit_locator.scroll_into_view_if_needed()
            except Exception:
                pass
            if submit_mode == "request":
                request_payload, gc_token_value = _build_request_payload(
                    page,
                    update_hasil_gc=update_hasil_gc,
                    hasil_gc_value=hasil_gc_after or hasil_gc,
                    update_lat=update_lat,
                    latitude_value=latitude_value,
                    update_lon=update_lon,
                    longitude_value=longitude_value,
                    update_nama=update_nama or edit_nama_alamat,
                    nama_value=nama_after or nama_usaha,
                    update_alamat=update_alamat or edit_nama_alamat,
                    alamat_value=alamat_after or alamat,
                    perusahaan_id_override=card_data_id or None,
                    gc_token_override=request_gc_token,
                )
                if gc_token_value:
                    request_gc_token = gc_token_value

                missing_fields = []
                if not request_payload.get("perusahaan_id"):
                    missing_fields.append("perusahaan_id")
                if not request_payload.get("_token"):
                    missing_fields.append("_token")
                if not request_payload.get("gc_token"):
                    missing_fields.append("gc_token")

                ready_for_request = True
                if missing_fields:
                    if missing_fields == ["gc_token"]:
                        log_warn(
                            "GC token belum terdeteksi; coba submit request tanpa gc_token.",
                            idsbr=idsbr or "-",
                        )
                        ready_for_request = True
                    else:
                        log_warn(
                            "Submit via request tidak siap; fallback ke UI.",
                            idsbr=idsbr or "-",
                            missing=", ".join(missing_fields),
                        )
                        ready_for_request = False
                if ready_for_request:
                    post_headers = {
                        "origin": "https://matchapro.web.bps.go.id",
                        "referer": TARGET_URL,
                        "x-requested-with": "XMLHttpRequest",
                    }
                    result = _submit_via_request(
                        page,
                        request_payload,
                        url=f"{TARGET_URL}/konfirmasi-user",
                        headers=post_headers,
                    )
                    if result.get("status") == "rate_limit":
                        wait_penalty = SUBMIT_RATE_LIMITER.penalize()
                        cooldown_s = result.get("cooldown_s") or 0
                        if cooldown_s:
                            set_server_cooldown(
                                cooldown_s, reason=result.get("reason") or ""
                            )
                        status = "error"
                        note = (
                            "Submit ditolak server (HTTP 429/rate limit). "
                            f"{result.get('message') or ''}".strip()
                        )
                        monitor.bot_goto(TARGET_URL)
                        continue
                    if result.get("status") == "csrf":
                        session_refresh_pending = True
                        status = "error"
                        note = (
                            "Submit gagal: CSRF token mismatch. "
                            "Session akan di-refresh."
                        )
                        monitor.bot_goto(TARGET_URL)
                        continue
                    if result.get("ok"):
                        payload_json = result.get("json")
                        if isinstance(payload_json, dict):
                            new_token = payload_json.get("new_gc_token")
                            if new_token:
                                request_gc_token = str(new_token)
                        SUBMIT_RATE_LIMITER.reset_penalty()
                        if SUBMIT_POST_SUCCESS_DELAY_S > 0:
                            monitor.wait_for_condition(
                                lambda: False,
                                timeout_s=SUBMIT_POST_SUCCESS_DELAY_S,
                            )
                        status = "berhasil"
                        note = build_success_note(result.get("message"))
                        session_submit_count += 1
                        monitor.bot_goto(TARGET_URL)
                        continue

                    status = "error"
                    note = result.get("message") or "Submit via request gagal."
                    monitor.bot_goto(TARGET_URL)
                    continue
            try:
                monitor.bot_click(submit_locator)
            except Exception as exc:
                log_warn(
                    "Tombol submit gagal diklik; skipping.",
                    idsbr=idsbr or "-",
                    error=str(exc),
                )
                status = "gagal"
                note = "Tombol submit gagal diklik"
                monitor.bot_goto(TARGET_URL)
                continue

            if validate_report_mode:
                submit_result = _handle_report_submit_result(page, monitor)
                if submit_result.get("rate_limit"):
                    wait_penalty = SUBMIT_RATE_LIMITER.penalize()
                    detail = (submit_result.get("message") or "").strip()
                    log_warn(
                        "Server menolak submit laporan validasi; kemungkinan rate limit.",
                        idsbr=idsbr or "-",
                        wait_s=round(wait_penalty, 1),
                        detail=detail or "-",
                    )
                    monitor.wait_for_condition(
                        lambda: False, timeout_s=wait_penalty
                    )
                    status = "error"
                    note = (
                        "Submit laporan validasi ditolak server (HTTP 429/rate limit)."
                        if not detail
                        else (
                            "Submit laporan validasi ditolak server (HTTP 429/rate limit). "
                            + detail
                        )
                    )
                    monitor.bot_goto(TARGET_URL)
                    continue
                if submit_result.get("ok"):
                    SUBMIT_RATE_LIMITER.reset_penalty()
                    if SUBMIT_POST_SUCCESS_DELAY_S > 0:
                        monitor.wait_for_condition(
                            lambda: False,
                            timeout_s=SUBMIT_POST_SUCCESS_DELAY_S,
                        )
                    status = "berhasil"
                    note = build_success_note(submit_result.get("message"))
                    session_submit_count += 1
                    monitor.bot_goto(TARGET_URL)
                    continue
                status = submit_result.get("status") or "gagal"
                note = (
                    submit_result.get("message")
                    or "Submit laporan validasi gagal."
                )
                monitor.bot_goto(TARGET_URL)
                continue

            confirm_text = "tanpa melakukan geotag"
            success_text = "Data submitted successfully"
            swal_result = None
            rate_limit_info = {"text": ""}

            def capture_rate_limit():
                rate_text = detect_rate_limit_popup_text(page).strip()
                if rate_text:
                    rate_limit_info["text"] = rate_text
                    return True
                return False

            def handle_rate_limit_abort():
                wait_penalty = SUBMIT_RATE_LIMITER.penalize()
                log_warn(
                    "Server menolak submit; kemungkinan rate limit.",
                    idsbr=idsbr or "-",
                    wait_s=round(wait_penalty, 1),
                    detail=rate_limit_info["text"],
                )
                try:
                    popup_locator = page.locator(".swal2-popup")
                    if rate_limit_info["text"]:
                        popup_locator = popup_locator.filter(
                            has_text=rate_limit_info["text"]
                        )
                    confirm_btn = popup_locator.locator(".swal2-confirm")
                    if confirm_btn.count() > 0:
                        monitor.bot_click(confirm_btn.first)
                except Exception:
                    pass
                monitor.wait_for_condition(
                    lambda: False, timeout_s=wait_penalty
                )
                return wait_penalty

            def find_swal():
                nonlocal swal_result
                if capture_rate_limit():
                    swal_result = "rate_limit"
                    return True
                confirm_popup = page.locator(
                    ".swal2-popup", has_text=confirm_text
                )
                if (
                    confirm_popup.count() > 0
                    and confirm_popup.first.is_visible()
                ):
                    swal_result = "confirm"
                    return True
                success_popup = page.locator(
                    ".swal2-popup", has_text=success_text
                )
                if (
                    success_popup.count() > 0
                    and success_popup.first.is_visible()
                ):
                    swal_result = "success"
                    return True
                return False

            monitor.wait_for_condition(find_swal, timeout_s=15)

            def read_current_coord(selector):
                locator = locate_visible(selector, form_scope)
                if locator is None:
                    locator = locate_visible(selector)
                if locator is None:
                    return "", False
                try:
                    return _normalize_coord_text(locator.input_value()), True
                except Exception:
                    return "", True

            if swal_result == "rate_limit":
                handle_rate_limit_abort()
                status = "error"
                detail = rate_limit_info["text"].strip()
                note = (
                    "Submit ditolak server (HTTP 429/rate limit)."
                    if not detail
                    else f"Submit ditolak server (HTTP 429/rate limit). {detail}"
                )
                monitor.bot_goto(TARGET_URL)
                continue

            if swal_result == "confirm":
                lat_text, lat_found = read_current_coord("#tt_latitude_cek_user")
                lon_text, lon_found = read_current_coord("#tt_longitude_cek_user")
                if lat_found or lon_found:
                    has_valid_lat = (
                        _parse_coord_value(lat_text, -90, 90) is not None
                    )
                    has_valid_lon = (
                        _parse_coord_value(lon_text, -180, 180) is not None
                    )
                else:
                    has_valid_lat = (
                        _parse_coord_value(latitude_value, -90, 90) is not None
                    )
                    has_valid_lon = (
                        _parse_coord_value(longitude_value, -180, 180) is not None
                    )
                if not (has_valid_lat and has_valid_lon):
                    confirm_popup = page.locator(
                        ".swal2-popup", has_text=confirm_text
                    )
                    confirm_button = confirm_popup.locator(
                        ".swal2-confirm", has_text="Ya"
                    )
                    if confirm_button.count() > 0:
                        monitor.bot_click(confirm_button.first)
                    else:
                        log_warn(
                            "Tombol Ya pada dialog geotag tidak ditemukan.",
                            idsbr=idsbr or "-",
                        )
                        status = "gagal"
                        note = "Dialog geotag tanpa tombol Ya"
                        monitor.bot_goto(TARGET_URL)
                        continue
                else:
                    log_warn(
                        "Dialog geotag muncul padahal koordinat ada.",
                        idsbr=idsbr or "-",
                    )
                    status = "gagal"
                    note = "Anomali dialog geotag"
                    monitor.bot_goto(TARGET_URL)
                    continue

            if swal_result != "success":
                swal_result = None

                def find_success():
                    nonlocal swal_result
                    if capture_rate_limit():
                        swal_result = "rate_limit"
                        return True
                    success_popup = page.locator(
                        ".swal2-popup", has_text=success_text
                    )
                    if (
                        success_popup.count() > 0
                        and success_popup.first.is_visible()
                    ):
                        swal_result = "success"
                        return True
                    return False

                if not monitor.wait_for_condition(find_success, timeout_s=15):
                    log_warn(
                        "Dialog sukses submit tidak muncul; skipping.",
                        idsbr=idsbr or "-",
                    )
                    status = "gagal"
                    note = "Dialog sukses tidak muncul"
                    monitor.bot_goto(TARGET_URL)
                    continue
                if swal_result == "rate_limit":
                    handle_rate_limit_abort()
                    status = "error"
                    detail = rate_limit_info["text"].strip()
                    note = (
                        "Submit ditolak server (HTTP 429/rate limit)."
                        if not detail
                        else f"Submit ditolak server (HTTP 429/rate limit). {detail}"
                    )
                    monitor.bot_goto(TARGET_URL)
                    continue

            success_popup = page.locator(
                ".swal2-popup", has_text=success_text
            )
            ok_button = success_popup.locator(
                ".swal2-confirm", has_text="OK"
            )
            if ok_button.count() == 0:
                log_warn(
                    "Tombol OK pada dialog sukses tidak ditemukan.",
                    idsbr=idsbr or "-",
                )
                status = "gagal"
                note = "Dialog sukses tanpa tombol OK"
                monitor.bot_goto(TARGET_URL)
                continue
            monitor.bot_click(ok_button.first)
            monitor.wait_for_condition(
                lambda: page.locator(".swal2-popup").count() == 0
                or not page.locator(".swal2-popup").first.is_visible(),
                timeout_s=10,
            )
            monitor.wait_for_condition(
                lambda: is_visible(page, "#search-idsbr")
                or page.locator(".usaha-card-header").count() > 0,
                timeout_s=10,
            )
            if not page.url.startswith(TARGET_URL):
                monitor.bot_goto(TARGET_URL)
            SUBMIT_RATE_LIMITER.reset_penalty()
            if SUBMIT_POST_SUCCESS_DELAY_S > 0:
                monitor.wait_for_condition(
                    lambda: False, timeout_s=SUBMIT_POST_SUCCESS_DELAY_S
                )
            status = "berhasil"
            note = build_success_note("")
            session_submit_count += 1
        except Exception as exc:
            message = str(exc)
            if "Run stopped by user." in message:
                log_warn(
                    "Run stopped by user.",
                    idsbr=idsbr or "-",
                    row=batch_index,
                    reason=message,
                )
                status = "stopped"
                note = message
                stop_reason = message
                stop_processing = True
            elif "Idle timeout reached" in message:
                log_warn(
                    "Idle timeout reached; stopping run.",
                    idsbr=idsbr or "-",
                    row=batch_index,
                    reason=message,
                )
                status = "error"
                note = message
                idle_reason = message
                stop_processing = True
            else:
                log_error(
                    "Error while processing row.",
                    idsbr=idsbr or "-",
                    error=message,
                )
                status = "error"
                note = message
        finally:
            combined_note = note or ""
            if extra_notes:
                extra_text = "; ".join(extra_notes)
                if combined_note:
                    combined_note = f"{combined_note}; {extra_text}"
                else:
                    combined_note = extra_text
            run_log_rows.append(
                {
                    "no": excel_row,
                    "idsbr": idsbr or "",
                    "nama_usaha": nama_usaha or "",
                    "alamat": alamat or "",
                    "keberadaanusaha_gc": hasil_gc if hasil_gc is not None else "",
                    "latitude": latitude_value or "",
                    "latitude_source": latitude_source or "",
                    "latitude_before": latitude_before or "",
                    "latitude_after": latitude_after or "",
                    "longitude": longitude_value or "",
                    "longitude_source": longitude_source or "",
                    "longitude_before": longitude_before or "",
                    "longitude_after": longitude_after or "",
                    "hasil_gc_before": hasil_gc_before or "",
                    "hasil_gc_after": hasil_gc_after or "",
                    "nama_usaha_before": nama_before or "",
                    "nama_usaha_after": nama_after or "",
                    "alamat_before": alamat_before or "",
                    "alamat_after": alamat_after or "",
                    "status": status or "error",
                    "catatan": combined_note,
                }
            )
            summary_status = status or "error"
            summary_note = combined_note or "-"
            summary_fields = {
                "row": batch_index,
                "row_excel": excel_row,
                "idsbr": idsbr or "-",
                "status": summary_status,
                "note": summary_note,
            }
            if summary_status == "berhasil":
                log_info("Row summary.", **summary_fields)
            elif summary_status in {"gagal", "skipped"}:
                log_warn("Row summary.", **summary_fields)
            else:
                log_error("Row summary.", **summary_fields)
            if progress_callback:
                try:
                    progress_callback(
                        stats["processed"],
                        selected_rows,
                        excel_row,
                    )
                except Exception:
                    pass
            checkpoint_run_log()
        if stop_processing:
            checkpoint_run_log(force=True)
            break

    if stop_reason:
        log_warn(
            "Processing stopped by user.",
            _spacer=True,
            _divider=True,
            **stats,
        )
    elif idle_reason:
        log_warn(
            "Processing stopped due to idle timeout.",
            _spacer=True,
            _divider=True,
            **stats,
        )
    elif cooldown_reason:
        log_warn(
            "Processing stopped due to server cooldown.",
            _spacer=True,
            _divider=True,
            **stats,
        )
    else:
        log_info("Processing completed.", _spacer=True, _divider=True, **stats)
    checkpoint_run_log(force=True)
    log_info("Run log saved.", path=str(run_log_path))
    if stop_reason:
        raise RuntimeError(stop_reason)
    if idle_reason:
        raise RuntimeError(idle_reason)
