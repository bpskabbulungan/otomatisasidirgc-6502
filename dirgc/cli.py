import argparse
import os
import sys

from playwright.sync_api import sync_playwright

from .browser import ActivityMonitor, ensure_on_dirgc, install_user_activity_tracking
from .credentials import load_credentials
from .processor import process_excel_rows
from .settings import (
    DEFAULT_CREDENTIALS_FILE,
    DEFAULT_EXCEL_FILE,
    DEFAULT_IDLE_TIMEOUT_MS,
    DEFAULT_WEB_TIMEOUT_S,
)


def build_parser():
    parser = argparse.ArgumentParser(
        description="Login, process Excel rows, and stop after filling GC fields."
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
    credentials=None,
    stop_event=None,
    progress_callback=None,
    wait_for_close=None,
):
    ensure_playwright_browsers()
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
                "X-Requested-With": "com.matchapro.app"
            },
            java_script_enabled=True,
            permissions=["geolocation"]
        )
        page = context.new_page()
        page.set_default_timeout(web_timeout_s * 1000)
        page.set_default_navigation_timeout(web_timeout_s * 1000)
        
        # STEALTH SCRIPTS - HAPUS SEMUA DETECTION FLAGS
        page.add_init_script("""
            // Hapus webdriver flag
            Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
            
            // Override Chrome detection
            window.chrome = {runtime: {}};
            
            // Permissions & languages Android
            Object.defineProperty(navigator, 'permissions', {
                get: () => ({query: () => Promise.resolve({state: 'granted'})})});
            
            // Plugins empty (mobile)
            Object.defineProperty(navigator, 'plugins', {get: () => [1,2,3,4,5]});
            
            // Languages Indonesia
            Object.defineProperty(navigator, 'languages', {get: () => ['id-ID', 'id', 'en-US', 'en']});
            
            // WebGL fingerprint spoof
            const getParameter = WebGLRenderingContext.getParameter;
            WebGLRenderingContext.prototype.getParameter = function(parameter) {
                if (parameter === 37445) return 'Intel Inc.';
                if (parameter === 37446) return 'Intel(R) UHD Graphics 630';
                return getParameter(parameter);
            };
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
        process_excel_rows(
            page,
            monitor=monitor,
            excel_file=excel_file,
            use_saved_credentials=not manual_only,
            credentials=credentials_value,
            start_row=start_row,
            end_row=end_row,
            progress_callback=progress_callback,
        )

        if keep_open:
            if wait_for_close:
                wait_for_close()
            else:
                input("Press Enter to close the browser...")

        context.close()
        browser.close()


def main():
    parser = build_parser()
    args = parser.parse_args()

    try:
        validate_row_range(args.start_row, args.end_row)
    except ValueError as exc:
        parser.error(str(exc))

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
