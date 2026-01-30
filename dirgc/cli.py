import argparse
import json
import os
import re
import sys
from urllib.parse import urlparse

from playwright.sync_api import sync_playwright

from .browser import (
    ActivityMonitor,
    apply_rate_limit_profile,
    ensure_on_dirgc,
    install_user_activity_tracking,
    set_server_cooldown,
)
from .credentials import load_credentials
from .logging_utils import log_info, log_warn
from .processor import process_excel_rows
from .settings import (
    DEFAULT_CREDENTIALS_FILE,
    DEFAULT_EXCEL_FILE,
    DEFAULT_IDLE_TIMEOUT_MS,
    DEFAULT_RATE_LIMIT_PROFILE,
    DEFAULT_WEB_TIMEOUT_S,
    BLOCK_RESOURCE_DOMAINS,
    BLOCK_RESOURCE_TYPES,
    ENABLE_RESOURCE_BLOCKING,
    RATE_LIMIT_PROFILES,
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
        help="Keep the browser open until you press Enter.",
    )
    parser.add_argument(
        "--dirgc-only",
        action="store_true",
        help="Stop after reaching the DIRGC page (skip filter/input).",
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
    stop_on_cooldown=False,
):
    ensure_playwright_browsers()
    apply_rate_limit_profile(rate_limit_profile)
    credentials_value = credentials
    if not manual_only and credentials_value is None:
        credentials_value = load_credentials(credentials_file)

    web_timeout_s = max(5, int(web_timeout_s or 0))
    timeout_scale = web_timeout_s / DEFAULT_WEB_TIMEOUT_S

    with sync_playwright() as p:
        # browser = p.chromium.launch(headless=headless)
        # context = browser.new_context()
        # page = context.new_page()
        browser = p.chromium.launch(
            headless=headless,
            args=[
                '--disable-blink-features=AutomationControlled',
                '--disable-dev-shm-usage',
                '--no-sandbox',
                '--disable-setuid-sandbox',
                '--disable-infobars',
                '--window-position=-5,-5',
                '--disable-extensions',
                '--user-agent=Mozilla/5.0 (Linux; Android 12; M2010J19CG Build/SKQ1.211202.001; wv) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/143.0.7499.192 Mobile Safari/537.36'
            ]
        )
        
        # CONTEXT ANDROID WEBVIEW
        context = browser.new_context(
            viewport={'width': 390, 'height': 844},
            screen={'width': 1080, 'height': 2340},
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
            permissions=["geolocation"]
        )
        if ENABLE_RESOURCE_BLOCKING and (BLOCK_RESOURCE_TYPES or BLOCK_RESOURCE_DOMAINS):
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
        page.set_default_timeout(web_timeout_s * 1000)
        page.set_default_navigation_timeout(web_timeout_s * 1000)
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
        if dirgc_only:
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


def install_429_response_logger(page, body_limit=800):
    def _truncate(text, limit):
        if text is None:
            return ""
        text = str(text)
        if len(text) <= limit:
            return text
        return f"{text[:limit]}... (truncated {len(text) - limit} chars)"

    def handle_response(response):
        try:
            status = response.status
        except Exception:
            return
        if status != 429:
            return

        method = ""
        url = ""
        try:
            url = response.url or ""
        except Exception:
            url = ""
        try:
            request = response.request
            method = request.method or ""
        except Exception:
            method = ""

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

        log_warn(
            "HTTP 429 response captured.",
            status=status,
            method=method or "-",
            url=url or "-",
            retry_after=retry_after or "-",
            content_type=content_type or "-",
            body=body_text,
            body_error=body_error,
        )

        cooldown_seconds, reason = _extract_cooldown_seconds(
            raw_body_text, headers
        )
        if cooldown_seconds:
            set_server_cooldown(cooldown_seconds, reason=reason)

    page.on("response", handle_response)


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
        stop_on_cooldown=args.stop_on_cooldown,
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
