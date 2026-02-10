import argparse
import json
import os
import re
import sys
import time
from urllib.parse import parse_qs, urlparse

from playwright.sync_api import sync_playwright

from .browser import (
    ActivityMonitor,
    apply_rate_limit_profile,
    ensure_on_dirgc,
    is_visible,
    install_user_activity_tracking,
    set_server_cooldown,
    wait_for_block_ui_clear,
)
from .credentials import load_credentials
from .logging_utils import log_info, log_warn
from .vpn import ensure_vpn_connected
from .processor import process_excel_rows
from .recap import run_recap
from .settings import (
    DEFAULT_CREDENTIALS_FILE,
    DEFAULT_EXCEL_FILE,
    DEFAULT_IDLE_TIMEOUT_MS,
    DEFAULT_RATE_LIMIT_PROFILE,
    DEFAULT_RECAP_LENGTH,
    DEFAULT_SESSION_REFRESH_EVERY,
    DEFAULT_SUBMIT_MODE,
    DEFAULT_WEB_TIMEOUT_S,
    BLOCK_RESOURCE_DOMAINS,
    BLOCK_RESOURCE_TYPES,
    BLOCK_UI_SELECTOR,
    ENABLE_RESOURCE_BLOCKING,
    MATCHAPRO_HOST,
    RATE_LIMIT_PROFILES,
    RECAP_LENGTH_MAX,
    RECAP_LENGTH_WARN_THRESHOLD,
    DEFAULT_VPN_PREFIXES,
    SUBMIT_MODES,
)


def build_parser():
    parser = argparse.ArgumentParser(
        description=(
            "Login, then process Excel rows or stop at the DIRGC page."
        )
    )
    parser.add_argument(
        "--headless",
        action="store_true",
        help="Run browser in headless mode (may not work for SSO).",
    )
    parser.add_argument(
        "-m",
        "--manual-only",
        action="store_true",
        help="Skip auto-fill and always use manual login.",
    )
    parser.add_argument(
        "-c",
        "--credentials-file",
        help=(
            "Path to JSON credentials file with username/password. "
            f"Defaults to {DEFAULT_CREDENTIALS_FILE} if present."
        ),
    )
    parser.add_argument(
        "-e",
        "--excel-file",
        help=(
            "Path to Excel file. "
            f"Defaults to {DEFAULT_EXCEL_FILE} if present."
        ),
    )
    parser.add_argument(
        "-start",
        "--start",
        dest="start_row",
        type=int,
        help="Start row (1-based) to process from the Excel file.",
    )
    parser.add_argument(
        "-end",
        "--end",
        dest="end_row",
        type=int,
        help="End row (1-based, inclusive) to process from the Excel file.",
    )
    parser.add_argument(
        "-t",
        "--idle-timeout-ms",
        type=int,
        default=DEFAULT_IDLE_TIMEOUT_MS,
        help="Stop if no user input or bot action for this long.",
    )
    parser.add_argument(
        "--web-timeout-s",
        type=int,
        default=DEFAULT_WEB_TIMEOUT_S,
        help="Default timeout (seconds) for web loading and waits.",
    )
    parser.add_argument(
        "-k",
        "--keep-open",
        action="store_true",
        default=True,
        help="Keep the browser open until you press Enter (default: on).",
    )
    parser.add_argument(
        "--no-keep-open",
        action="store_false",
        dest="keep_open",
        help="Close the browser automatically after the run.",
    )
    parser.add_argument(
        "--dirgc-only",
        action="store_true",
        help="Stop after reaching the DIRGC page (skip filter/input).",
    )
    parser.add_argument(
        "--recap",
        dest="recap",
        action="store_true",
        help="Fetch DIRGC data via API and export Excel rekap.",
    )
    parser.add_argument(
        "--recap-length",
        dest="recap_length",
        type=int,
        default=DEFAULT_RECAP_LENGTH,
        help=(
            f"Pagination size for recap (default: {DEFAULT_RECAP_LENGTH}). "
            "Bigger = fewer requests but heavier per page; server may cap. "
            f"Max {RECAP_LENGTH_MAX}; above {RECAP_LENGTH_WARN_THRESHOLD} warns."
        ),
    )
    parser.add_argument(
        "--recap-output-dir",
        dest="recap_output_dir",
        help="Output directory for rekap Excel (default: logs/recap).",
    )
    parser.add_argument(
        "--recap-sleep-ms",
        dest="recap_sleep_ms",
        type=int,
        default=800,
        help="Delay between recap requests in ms (default: 800).",
    )
    parser.add_argument(
        "--recap-max-retries",
        dest="recap_max_retries",
        type=int,
        default=3,
        help="Max retries per page request (default: 3).",
    )
    parser.add_argument(
        "--recap-backup-every",
        dest="recap_backup_every",
        type=int,
        default=10,
        help=(
            "Create .bak backup every N batches during recap "
            "(default: 10). Use 0 to disable."
        ),
    )
    parser.add_argument(
        "--recap-no-resume",
        dest="recap_no_resume",
        action="store_true",
        help="Ignore existing checkpoint and start a new rekap run.",
    )
    parser.add_argument(
        "--api-log",
        action="store_true",
        help=(
            "Log API endpoints + sample responses from MatchaPro "
            "to logs/api."
        ),
    )
    parser.add_argument(
        "--api-log-dir",
        help="Output directory for API log (default: logs/api).",
    )
    parser.add_argument(
        "--api-log-max",
        type=int,
        default=30,
        help="Maximum unique endpoints to capture (default: 30).",
    )
    parser.add_argument(
        "--api-log-body-limit",
        type=int,
        default=2000,
        help="Max response preview chars per endpoint (default: 2000).",
    )
    parser.add_argument(
        "--api-log-wait-s",
        type=int,
        default=20,
        help=(
            "Seconds to wait for API responses after DIRGC loads "
            "(default: 20)."
        ),
    )
    parser.add_argument(
        "--edit-nama-alamat",
        action="store_true",
        help="Aktifkan toggle edit Nama/Alamat Usaha dan isi dari Excel.",
    )
    parser.add_argument(
        "--update-mode",
        action="store_true",
        help="Gunakan mode update data (klik Edit Hasil).",
    )
    parser.add_argument(
        "--update-fields",
        help=(
            "Daftar field yang di-update saat --update-mode, "
            "pisahkan dengan koma (hasil_gc,nama_usaha,alamat,latitude,longitude,koordinat)."
        ),
    )
    parser.add_argument(
        "--rate-limit-profile",
        choices=sorted(RATE_LIMIT_PROFILES.keys()),
        default=DEFAULT_RATE_LIMIT_PROFILE,
        help=(
            "Profil jeda untuk mengurangi rate limit saat submit "
            f"(default: {DEFAULT_RATE_LIMIT_PROFILE})."
        ),
    )
    parser.add_argument(
        "--submit-mode",
        choices=sorted(SUBMIT_MODES),
        default=DEFAULT_SUBMIT_MODE,
        help=(
            "Metode submit: ui (default) memakai klik UI, "
            "request memakai POST langsung ke endpoint konfirmasi."
        ),
    )
    parser.add_argument(
        "--session-refresh-every",
        type=int,
        default=DEFAULT_SESSION_REFRESH_EVERY,
        help=(
            "Refresh session otomatis setiap N submit sukses "
            "(0 untuk menonaktifkan)."
        ),
    )
    parser.add_argument(
        "--vpn-prefixes",
        help=(
            "Prefix IP VPN yang dianggap valid (contoh: 10.,172.16.). "
            "Pisahkan dengan koma."
        ),
    )
    parser.add_argument(
        "--stop-on-cooldown",
        action="store_true",
        help=(
            "Hentikan proses jika server meminta jeda (HTTP 429) "
            "dan simpan posisi terakhir untuk dilanjutkan."
        ),
    )
    coord_group = parser.add_mutually_exclusive_group()
    coord_group.add_argument(
        "--prefer-excel-coords",
        action="store_true",
        help=(
            "Utamakan koordinat dari Excel (default). "
            "Jika ada nilai Excel, overwrite nilai web."
        ),
    )
    coord_group.add_argument(
        "--prefer-web-coords",
        action="store_true",
        help=(
            "Utamakan koordinat dari web. "
            "Jika field web sudah terisi, tidak dioverwrite."
        ),
    )
    return parser


def validate_row_range(start_row, end_row):
    if start_row is not None and start_row < 1:
        raise ValueError("--start must be >= 1.")
    if end_row is not None and end_row < 1:
        raise ValueError("--end must be >= 1.")
    if start_row is not None and end_row is not None and start_row > end_row:
        raise ValueError("--start must be <= --end.")


def run_dirgc(
    *,
    headless=False,
    manual_only=False,
    credentials_file=None,
    excel_file=None,
    start_row=None,
    end_row=None,
    idle_timeout_ms=DEFAULT_IDLE_TIMEOUT_MS,
    web_timeout_s=DEFAULT_WEB_TIMEOUT_S,
    keep_open=False,
    dirgc_only=False,
    edit_nama_alamat=False,
    prefer_excel_coords=True,
    update_mode=False,
    update_fields=None,
    credentials=None,
    stop_event=None,
    progress_callback=None,
    wait_for_close=None,
    rate_limit_profile=DEFAULT_RATE_LIMIT_PROFILE,
    submit_mode=DEFAULT_SUBMIT_MODE,
    session_refresh_every=DEFAULT_SESSION_REFRESH_EVERY,
    vpn_prefixes=None,
    stop_on_cooldown=False,
    api_log=False,
    api_log_dir=None,
    api_log_max=30,
    api_log_body_limit=2000,
    api_log_wait_s=20,
    recap=False,
    recap_length=DEFAULT_RECAP_LENGTH,
    recap_output_dir=None,
    recap_sleep_ms=800,
    recap_max_retries=3,
    recap_resume=True,
    recap_backup_every=10,
):
    prefixes = None
    if isinstance(vpn_prefixes, str) and vpn_prefixes.strip():
        prefixes = [item.strip() for item in vpn_prefixes.split(",")]
    elif isinstance(DEFAULT_VPN_PREFIXES, (list, tuple)):
        prefixes = list(DEFAULT_VPN_PREFIXES)
    ensure_vpn_connected(prefixes)
    ensure_playwright_browsers()
    apply_rate_limit_profile(rate_limit_profile)
    credentials_value = credentials
    if not manual_only and credentials_value is None:
        credentials_value = load_credentials(credentials_file)

    web_timeout_s = max(5, int(web_timeout_s or 0))
    timeout_scale = web_timeout_s / DEFAULT_WEB_TIMEOUT_S

    with sync_playwright() as p:
        launch_args = [
            "--disable-blink-features=AutomationControlled",
            "--disable-dev-shm-usage",
            "--no-sandbox",
            "--disable-setuid-sandbox",
            "--disable-infobars",
            "--window-position=-5,-5",
            "--disable-extensions",
            "--user-agent=Mozilla/5.0 (Linux; Android 12; M2010J19CG Build/SKQ1.211202.001; wv) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/143.0.7499.192 Mobile Safari/537.36",
        ]
        browser = p.chromium.launch(headless=headless, args=launch_args)

        # CONTEXT ANDROID WEBVIEW
        context = browser.new_context(
            viewport={"width": 390, "height": 844},
            screen={"width": 1080, "height": 2340},
            device_scale_factor=2.625,
            is_mobile=True,
            has_touch=True,
            user_agent="Mozilla/5.0 (Linux; Android 12; M2010J19CG Build/SKQ1.211202.001; wv) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/143.0.7499.192 Mobile Safari/537.36",
            extra_http_headers={
                "Sec-Ch-Ua": '"Android WebView";v="143", "Chromium";v="143", "Not A(Brand";v="24"',
                "Sec-Ch-Ua-Mobile": "?1",
                "Sec-Ch-Ua-Platform": '"Android"',
            },
            java_script_enabled=True,
            permissions=["geolocation"],
        )

        resource_blocking_enabled = (
            ENABLE_RESOURCE_BLOCKING
            and not api_log
            and (BLOCK_RESOURCE_TYPES or BLOCK_RESOURCE_DOMAINS)
        )
        if resource_blocking_enabled:
            def _route_handler(route):
                domain = urlparse(route.request.url).netloc.lower()
                if domain in BLOCK_RESOURCE_DOMAINS:
                    route.abort()
                    return
                if route.request.resource_type in BLOCK_RESOURCE_TYPES:
                    route.abort()
                else:
                    route.continue_()

            context.route("**/*", _route_handler)
        page = context.new_page()
        install_429_response_logger(page)
        api_log_state = None
        if api_log:
            output_dir = api_log_dir or os.path.join(os.getcwd(), "logs", "api")
            api_log_state = install_api_logger(
                page,
                output_dir=output_dir,
                max_entries=api_log_max,
                body_limit=api_log_body_limit,
            )
            log_info(
                "API logging enabled.",
                output_dir=output_dir,
                max_entries=api_log_max,
                body_limit=api_log_body_limit,
            )
        page.set_default_timeout(web_timeout_s * 1000)
        page.set_default_navigation_timeout(web_timeout_s * 1000)
        if resource_blocking_enabled:
            def speed_route(route, request):
                rt = request.resource_type
                url = request.url

                # block heavy third-party + visual assets
                if "fonts.gstatic.com" in url or "fonts.googleapis.com" in url:
                    return route.abort()
                if rt in ("image", "font", "media"):
                    return route.abort()

                return route.continue_()

            page.route("**/*", speed_route)

        context.add_init_script("""
            // Anti-redefine conflict
            (function() {
                if (navigator.webdriver !== undefined) {
                    Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
                }
                window.chrome = {runtime: {}};
                Object.defineProperty(navigator, 'plugins', {get: () => [1,2,3,4,5]});
                Object.defineProperty(navigator, 'languages', {get: () => ['id-ID','en']});
            })();
        """)


        monitor = ActivityMonitor(
            page,
            idle_timeout_ms,
            stop_event=stop_event,
            timeout_scale=timeout_scale,
        )
        install_user_activity_tracking(page, monitor.mark_activity)

        ensure_on_dirgc(
            page,
            monitor=monitor,
            use_saved_credentials=not manual_only,
            credentials=credentials_value,
        )
        if dirgc_only or api_log or recap:
            wait_for_dirgc_ready(page, monitor, timeout_s=30)
            if dirgc_only:
                ensure_dirgc_interactive(page, monitor, timeout_s=15)
        if api_log and api_log_state and api_log_wait_s:
            wait_for_api_capture(
                monitor,
                api_log_state,
                timeout_s=api_log_wait_s,
            )
        if recap:
            run_recap(
                page,
                monitor,
                status_filter="semua",
                length=recap_length,
                output_dir=recap_output_dir,
                sleep_ms=recap_sleep_ms,
                max_retries=recap_max_retries,
                resume=recap_resume,
                backup_every=recap_backup_every,
                progress_callback=progress_callback,
            )
        elif dirgc_only:
            log_info(
                "DIRGC page ready; skipping Excel processing.",
                url=page.url,
            )
            if progress_callback:
                try:
                    progress_callback(0, 0, 0)
                except Exception:
                    pass
        else:
            process_excel_rows(
                page,
                monitor=monitor,
                excel_file=excel_file,
                use_saved_credentials=not manual_only,
                credentials=credentials_value,
                edit_nama_alamat=edit_nama_alamat,
                prefer_excel_coords=prefer_excel_coords,
                update_mode=update_mode,
                update_fields=update_fields,
                start_row=start_row,
                end_row=end_row,
                progress_callback=progress_callback,
                submit_mode=submit_mode,
                session_refresh_every=session_refresh_every,
                stop_on_cooldown=stop_on_cooldown,
            )

        if keep_open:
            if wait_for_close:
                wait_for_close()
            else:
                input("Press Enter to close the browser...")

        context.close()
        browser.close()


def _parse_int(value):
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return None


def _extract_cooldown_from_message(message):
    if not message:
        return None
    text = str(message).lower()
    minute_match = re.search(r"tunggu\s+(\d+)\s*(menit|minute)", text)
    if minute_match:
        minutes = _parse_int(minute_match.group(1))
        if minutes is not None:
            return minutes * 60
    second_match = re.search(r"tunggu\s+(\d+)\s*(detik|second)", text)
    if second_match:
        seconds = _parse_int(second_match.group(1))
        if seconds is not None:
            return seconds
    return None


def _extract_cooldown_seconds(body_text, headers):
    seconds = 0
    reason = ""

    retry_after = None
    if headers:
        retry_after = headers.get("retry-after") or headers.get("Retry-After")
    retry_seconds = _parse_int(retry_after)
    if retry_seconds is not None:
        seconds = max(seconds, retry_seconds)

    body = (body_text or "").strip()
    if body:
        parsed = None
        if body.startswith("{") or body.startswith("["):
            try:
                parsed = json.loads(body)
            except json.JSONDecodeError:
                parsed = None
        if isinstance(parsed, dict):
            message = parsed.get("message")
            reason = message or reason
            msg_seconds = _extract_cooldown_from_message(message)
            if msg_seconds:
                seconds = max(seconds, msg_seconds)
            retry_after_body = _parse_int(parsed.get("retry_after"))
            if retry_after_body is not None:
                seconds = max(seconds, retry_after_body)
        else:
            msg_seconds = _extract_cooldown_from_message(body)
            if msg_seconds:
                seconds = max(seconds, msg_seconds)
    return seconds, reason


def install_429_response_logger(page, body_limit=800, request_body_limit=800):
    def _truncate(text, limit):
        if text is None:
            return ""
        text = str(text)
        if len(text) <= limit:
            return text
        return f"{text[:limit]}... (truncated {len(text) - limit} chars)"

    def _redact_payload(text):
        if not text:
            return ""
        sensitive = {
            "password",
            "pass",
            "otp",
            "token",
            "_token",
            "gc_token",
            "authorization",
        }
        try:
            if text.strip().startswith("{"):
                parsed = json.loads(text)
                if isinstance(parsed, dict):
                    for key in list(parsed.keys()):
                        if key.lower() in sensitive:
                            parsed[key] = "***"
                    return json.dumps(parsed, ensure_ascii=False)
        except Exception:
            pass
        try:
            parsed = parse_qs(text, keep_blank_values=True)
            if parsed:
                redacted = {}
                for key, values in parsed.items():
                    if key.lower() in sensitive:
                        redacted[key] = ["***"]
                    else:
                        redacted[key] = values
                parts = []
                for key, values in redacted.items():
                    for value in values:
                        parts.append(f"{key}={value}")
                return "&".join(parts)
        except Exception:
            pass
        return text

    def handle_response(response):
        try:
            status = response.status
        except Exception:
            return
        if status < 400:
            return

        url = ""
        method = ""
        request_body = ""
        try:
            url = response.url or ""
        except Exception:
            url = ""
        if MATCHAPRO_HOST not in url:
            return
        try:
            request = response.request
            method = request.method or ""
            request_body = request.post_data or ""
        except Exception:
            method = ""
            request_body = ""

        headers = {}
        try:
            headers = response.headers or {}
        except Exception:
            headers = {}
        retry_after = headers.get("retry-after") or headers.get("Retry-After")
        content_type = headers.get("content-type") or headers.get(
            "Content-Type"
        )

        body_text = ""
        body_error = ""
        try:
            body_text = response.text() or ""
        except Exception as exc:
            try:
                body_bytes = response.body()
                body_text = body_bytes.decode("utf-8", errors="replace")
            except Exception as inner:
                body_error = str(inner) or str(exc)

        raw_body_text = body_text.strip() if body_text else ""
        body_text = _truncate(raw_body_text, body_limit) if raw_body_text else ""
        request_body = _truncate(
            _redact_payload(request_body.strip()), request_body_limit
        )

        log_warn(
            "HTTP error response captured.",
            status=status,
            method=method or "-",
            url=url or "-",
            retry_after=retry_after or "-",
            content_type=content_type or "-",
            body=body_text,
            request_body=request_body or "-",
            body_error=body_error,
        )

        if status == 429:
            cooldown_seconds, reason = _extract_cooldown_seconds(
                raw_body_text, headers
            )
            if cooldown_seconds:
                set_server_cooldown(cooldown_seconds, reason=reason)

    page.on("response", handle_response)


def wait_for_dirgc_ready(page, monitor, timeout_s=30):
    def is_ready():
        selectors = [
            "#search-idsbr",
            "#toggle-filter",
            ".usaha-card-header",
            ".empty-state",
            ".no-data",
            ".no-results",
        ]
        for selector in selectors:
            locator = page.locator(selector)
            if locator.count() > 0 and locator.first.is_visible():
                return True
        return False

    if monitor.wait_for_condition(is_ready, timeout_s=timeout_s):
        return True

    log_warn("DIRGC UI not ready; retrying with reload.")
    try:
        page.reload(wait_until="domcontentloaded")
    except Exception:
        return False
    return monitor.wait_for_condition(is_ready, timeout_s=timeout_s)



def ensure_dirgc_interactive(page, monitor, timeout_s=15):
    try:
        wait_for_block_ui_clear(page, monitor, timeout_s=timeout_s)
    except Exception:
        pass
    if is_visible(page, BLOCK_UI_SELECTOR):
        log_warn(
            "DIRGC UI masih loading; mencoba melepas overlay untuk manual.",
            selector=BLOCK_UI_SELECTOR,
        )
        try:
            page.evaluate(
                """
                () => {
                  const overlays = document.querySelectorAll(
                    '.blockUI.blockOverlay, .blockUI.blockMsg'
                  );
                  overlays.forEach((el) => {
                    el.style.display = 'none';
                    el.style.pointerEvents = 'none';
                    if (el.parentNode) {
                      el.parentNode.removeChild(el);
                    }
                  });
                }
                """
            )
        except Exception as exc:
            log_warn("Gagal melepas overlay loading.", error=str(exc))
            return False
    try:
        def loading_overlay_gone():
            return not page.evaluate(
                """
                () => {
                  const viewportW = window.innerWidth || 0;
                  const viewportH = window.innerHeight || 0;
                  const minW = viewportW * 0.7;
                  const minH = viewportH * 0.7;
                  const isVisible = (el) => {
                    if (!el) return false;
                    const style = getComputedStyle(el);
                    if (style.display === 'none' || style.visibility === 'hidden') return false;
                    if (parseFloat(style.opacity || '1') <= 0) return false;
                    const r = el.getBoundingClientRect();
                    return r.width > 1 && r.height > 1;
                  };
                  const candidates = Array.from(
                    document.querySelectorAll('.loading')
                  );
                  for (const el of candidates) {
                    if (!isVisible(el)) continue;
                    const r = el.getBoundingClientRect();
                    const coversScreen = r.width >= minW && r.height >= minH;
                    const style = getComputedStyle(el);
                    const isOverlay =
                      style.position === 'fixed' || style.position === 'absolute';
                    if (coversScreen && isOverlay) {
                      return false;
                    }
                  }
                  return true;
                }
                """
            )
        if not monitor.wait_for_condition(
            loading_overlay_gone, timeout_s=timeout_s
        ):
            removed = page.evaluate(
                """
                () => {
                  const viewportW = window.innerWidth || 0;
                  const viewportH = window.innerHeight || 0;
                  const minW = viewportW * 0.7;
                  const minH = viewportH * 0.7;
                  const isVisible = (el) => {
                    if (!el) return false;
                    const style = getComputedStyle(el);
                    if (style.display === 'none' || style.visibility === 'hidden') return false;
                    if (parseFloat(style.opacity || '1') <= 0) return false;
                    const r = el.getBoundingClientRect();
                    return r.width > 1 && r.height > 1;
                  };
                  let removedCount = 0;
                  const candidates = Array.from(
                    document.querySelectorAll('.loading')
                  );
                  candidates.forEach((el) => {
                    if (!isVisible(el)) return;
                    const r = el.getBoundingClientRect();
                    const coversScreen = r.width >= minW && r.height >= minH;
                    const style = getComputedStyle(el);
                    const isOverlay =
                      style.position === 'fixed' || style.position === 'absolute';
                    if (coversScreen && isOverlay) {
                      el.style.display = 'none';
                      el.style.pointerEvents = 'none';
                      if (el.parentNode) {
                        el.parentNode.removeChild(el);
                      }
                      removedCount += 1;
                    }
                  });
                  return removedCount;
                }
                """
            )
            if removed:
                log_warn("Menghapus overlay loading .loading.", count=removed)
    except Exception as exc:
        log_warn("Gagal mengecek overlay .loading.", error=str(exc))
    return True


def wait_for_api_capture(monitor, api_log_state, timeout_s=20):
    def has_capture():
        entries = api_log_state.get("entries") or []
        for entry in entries:
            if entry.get("response_preview"):
                return True
            if entry.get("response_error") and entry.get("response_error") != "-":
                return True
        return False

    if monitor.wait_for_condition(has_capture, timeout_s=timeout_s):
        return True
    log_warn(
        "No API response captured yet. Keep the browser open and try scrolling.",
        wait_s=timeout_s,
        output_path=api_log_state.get("output_path", "-"),
    )
    return False


def install_api_logger(
    page,
    *,
    output_dir=None,
    max_entries=30,
    body_limit=2000,
):
    if not output_dir:
        output_dir = os.path.join(os.getcwd(), "logs", "api")
    os.makedirs(output_dir, exist_ok=True)
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    output_path = os.path.join(output_dir, f"api_endpoints_{timestamp}.json")
    seen = set()
    entries = []
    state = {
        "entries": entries,
        "output_path": output_path,
    }

    def _truncate_text(text, limit):
        if text is None:
            return ""
        text = str(text)
        if len(text) <= limit:
            return text
        return f"{text[:limit]}... (truncated {len(text) - limit} chars)"

    def _redact_payload(text):
        if not text:
            return ""
        sensitive = {
            "password",
            "pass",
            "otp",
            "token",
            "_token",
            "gc_token",
            "authorization",
        }
        try:
            if text.strip().startswith("{"):
                parsed = json.loads(text)
                if isinstance(parsed, dict):
                    for key in list(parsed.keys()):
                        if key.lower() in sensitive:
                            parsed[key] = "***"
                    return json.dumps(parsed, ensure_ascii=False)
        except Exception:
            pass
        try:
            parsed = parse_qs(text, keep_blank_values=True)
            if parsed:
                redacted = {}
                for key, values in parsed.items():
                    if key.lower() in sensitive:
                        redacted[key] = ["***"]
                    else:
                        redacted[key] = values
                parts = []
                for key, values in redacted.items():
                    for value in values:
                        parts.append(f"{key}={value}")
                return "&".join(parts)
        except Exception:
            pass
        return text

    def _redact_obj(obj, depth=0, max_depth=4):
        sensitive = {
            "password",
            "pass",
            "otp",
            "token",
            "_token",
            "gc_token",
            "authorization",
            "access_token",
            "refresh_token",
        }
        if depth > max_depth:
            return obj
        if isinstance(obj, dict):
            redacted = {}
            for key, value in obj.items():
                if str(key).lower() in sensitive:
                    redacted[key] = "***"
                else:
                    redacted[key] = _redact_obj(value, depth + 1, max_depth)
            return redacted
        if isinstance(obj, list):
            limited = obj[:5]
            tail = "..." if len(obj) > 5 else None
            cleaned = [
                _redact_obj(value, depth + 1, max_depth)
                for value in limited
            ]
            if tail is not None:
                cleaned.append(tail)
            return cleaned
        return obj

    def _sample_payload(payload):
        if isinstance(payload, dict) and isinstance(payload.get("data"), list):
            payload = dict(payload)
            payload["data"] = payload["data"][:3]
            return payload
        if isinstance(payload, list):
            return payload[:3]
        return payload

    def _write_entries():
        try:
            with open(output_path, "w", encoding="utf-8") as handle:
                json.dump(entries, handle, ensure_ascii=False, indent=2)
        except Exception:
            return

    def handle_response(response):
        if len(entries) >= max_entries:
            return

        try:
            url = response.url or ""
        except Exception:
            return

        if MATCHAPRO_HOST not in url:
            return

        try:
            request = response.request
            method = request.method or "GET"
            resource_type = request.resource_type or ""
            request_body = request.post_data or ""
        except Exception:
            method = "GET"
            resource_type = ""
            request_body = ""

        if resource_type and resource_type not in ("xhr", "fetch"):
            return

        try:
            headers = response.headers or {}
        except Exception:
            headers = {}
        content_type = headers.get("content-type") or headers.get("Content-Type") or ""
        if "json" not in content_type.lower():
            return

        try:
            status = response.status
        except Exception:
            status = None

        endpoint = urlparse(url)
        endpoint_key = f"{method} {endpoint.scheme}://{endpoint.netloc}{endpoint.path}"
        if endpoint_key in seen:
            return
        seen.add(endpoint_key)

        response_preview = ""
        response_error = ""
        response_size = headers.get("content-length") or ""
        try:
            body_bytes = response.body()
            response_text = ""
            if body_bytes:
                response_text = body_bytes.decode("utf-8", errors="replace")
            if response_text.strip().startswith("{") or response_text.strip().startswith("["):
                parsed = json.loads(response_text)
                sampled = _sample_payload(parsed)
                redacted = _redact_obj(sampled)
                response_preview = _truncate_text(
                    json.dumps(redacted, ensure_ascii=False), body_limit
                )
            else:
                response_preview = _truncate_text(response_text, body_limit)
        except Exception as exc:
            response_error = str(exc)

        entry = {
            "endpoint": endpoint_key,
            "sample_url": url,
            "method": method,
            "status": status,
            "content_type": content_type or "-",
            "content_length": response_size or "-",
            "request_body": _truncate_text(
                _redact_payload(request_body), body_limit
            ),
            "response_preview": response_preview,
            "response_error": response_error or "-",
            "captured_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        }
        entries.append(entry)
        _write_entries()
        if response_preview:
            log_info(
                "API response captured.",
                endpoint=endpoint_key,
                status=status,
                output_path=output_path,
            )
        elif response_error and response_error != "-":
            log_warn(
                "API response captured but body could not be read.",
                endpoint=endpoint_key,
                status=status,
                error=response_error,
                output_path=output_path,
            )

    page.on("response", handle_response)
    return state


def main():
    parser = build_parser()
    args = parser.parse_args()

    try:
        validate_row_range(args.start_row, args.end_row)
    except ValueError as exc:
        parser.error(str(exc))

    update_fields = None
    if args.update_fields:
        raw_fields = [item.strip().lower() for item in args.update_fields.split(",")]
        raw_fields = [item for item in raw_fields if item]
        if raw_fields:
            update_fields = raw_fields

    run_dirgc(
        headless=args.headless,
        manual_only=args.manual_only,
        credentials_file=args.credentials_file,
        excel_file=args.excel_file,
        start_row=args.start_row,
        end_row=args.end_row,
        idle_timeout_ms=args.idle_timeout_ms,
        web_timeout_s=args.web_timeout_s,
        keep_open=args.keep_open,
        dirgc_only=args.dirgc_only,
        edit_nama_alamat=args.edit_nama_alamat,
        prefer_excel_coords=not args.prefer_web_coords,
        update_mode=args.update_mode,
        update_fields=update_fields,
        rate_limit_profile=args.rate_limit_profile,
        submit_mode=args.submit_mode,
        session_refresh_every=args.session_refresh_every,
        vpn_prefixes=args.vpn_prefixes,
        stop_on_cooldown=args.stop_on_cooldown,
        api_log=args.api_log,
        api_log_dir=args.api_log_dir,
        api_log_max=args.api_log_max,
        api_log_body_limit=args.api_log_body_limit,
        api_log_wait_s=args.api_log_wait_s,
        recap=args.recap,
        recap_length=args.recap_length,
        recap_output_dir=args.recap_output_dir,
        recap_sleep_ms=args.recap_sleep_ms,
        recap_max_retries=args.recap_max_retries,
        recap_backup_every=args.recap_backup_every,
        recap_resume=not args.recap_no_resume,
    )


def ensure_playwright_browsers():
    if os.getenv("PLAYWRIGHT_BROWSERS_PATH"):
        return

    if getattr(sys, "frozen", False):
        base_dir = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
        bundled_path = os.path.join(base_dir, "playwright-browsers")
        if os.path.exists(bundled_path):
            os.environ["PLAYWRIGHT_BROWSERS_PATH"] = bundled_path
            return

    local_path = os.path.join(os.getcwd(), "playwright-browsers")
    if os.path.exists(local_path):
        os.environ["PLAYWRIGHT_BROWSERS_PATH"] = local_path


if __name__ == "__main__":
    main()
