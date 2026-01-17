import argparse

from playwright.sync_api import sync_playwright

from .browser import ActivityMonitor, ensure_on_dirgc, install_user_activity_tracking
from .credentials import load_credentials
from .processor import process_excel_rows
from .settings import (
    DEFAULT_CREDENTIALS_FILE,
    DEFAULT_EXCEL_FILE,
    DEFAULT_IDLE_TIMEOUT_MS,
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
        "-k",
        "--keep-open",
        action="store_true",
        help="Keep the browser open until you press Enter.",
    )
    return parser


def main():
    parser = build_parser()
    args = parser.parse_args()

    if args.start_row is not None and args.start_row < 1:
        parser.error("--start must be >= 1.")
    if args.end_row is not None and args.end_row < 1:
        parser.error("--end must be >= 1.")
    if (
        args.start_row is not None
        and args.end_row is not None
        and args.start_row > args.end_row
    ):
        parser.error("--start must be <= --end.")

    with sync_playwright() as p:
        # browser = p.chromium.launch(headless=args.headless)
        # context = browser.new_context()
        # page = context.new_page()
        browser = p.chromium.launch(
            headless=args.headless,
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

        monitor = ActivityMonitor(page, args.idle_timeout_ms)
        install_user_activity_tracking(page, monitor.mark_activity)

        credentials = None
        if not args.manual_only:
            credentials = load_credentials(args.credentials_file)

        ensure_on_dirgc(
            page,
            monitor=monitor,
            use_saved_credentials=not args.manual_only,
            credentials=credentials,
        )
        process_excel_rows(
            page,
            monitor=monitor,
            excel_file=args.excel_file,
            use_saved_credentials=not args.manual_only,
            credentials=credentials,
            start_row=args.start_row,
            end_row=args.end_row,
        )

        if args.keep_open:
            input("Press Enter to close the browser...")

        context.close()
        browser.close()


if __name__ == "__main__":
    main()
