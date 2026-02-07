import random
import re
import time

from playwright.sync_api import TimeoutError as PWTimeoutError
from .logging_utils import log_info, log_warn
from .settings import (
    AUTO_LOGIN_RESULT_TIMEOUT_S,
    BLOCK_UI_SELECTOR,
    DEFAULT_RATE_LIMIT_PROFILE,
    HASIL_GC_LABELS,
    LOGIN_PATH,
    MATCHAPRO_HOST,
    RATE_LIMIT_PROFILES,
    SSO_HOST,
    TARGET_URL,
)


class ActivityMonitor:
    def __init__(self, page, idle_timeout_ms, stop_event=None, timeout_scale=1.0):
        self.page = page
        self.idle_timeout_s = idle_timeout_ms / 1000
        self.last_activity = time.monotonic()
        self.stop_event = stop_event
        self._stop_logged = False
        self.timeout_scale = timeout_scale if timeout_scale and timeout_scale > 0 else 1.0

    def _check_stop(self):
        if self.stop_event and self.stop_event.is_set():
            if not self._stop_logged:
                self._stop_logged = True
                log_warn("Run stopped by user.")
            raise RuntimeError("Run stopped by user.")

    def mark_activity(self, _reason=None):
        self.last_activity = time.monotonic()

    def idle_check(self):
        self._check_stop()
        if time.monotonic() - self.last_activity > self.idle_timeout_s:
            total_seconds = int(round(self.idle_timeout_s))
            minutes = total_seconds // 60
            seconds = total_seconds % 60
            if minutes >= 1:
                message = (
                    "Idle timeout reached "
                    f"({minutes} minutes {seconds} seconds without activity)."
                )
            else:
                message = (
                    "Idle timeout reached "
                    f"({total_seconds} seconds without activity)."
                )
            raise RuntimeError(message)

    def scale_timeout(self, timeout_s):
        if timeout_s is None:
            return None
        return timeout_s * self.timeout_scale

    def wait_for_condition(self, condition, timeout_s=None, poll_ms=1500):
        timeout_s = self.scale_timeout(timeout_s)
        start = time.monotonic()
        while True:
            if condition():
                return True
            if timeout_s is not None and time.monotonic() - start > timeout_s:
                return False
            self.idle_check()
            self.page.wait_for_timeout(poll_ms)

    def bot_click(self, selector_or_locator):
        self._check_stop()
        self.mark_activity("bot")
        if isinstance(selector_or_locator, str):
            self.page.click(selector_or_locator)
        else:
            selector_or_locator.click()

    def bot_fill(self, selector, value):
        self._check_stop()
        self.mark_activity("bot")
        self.page.fill(selector, "" if value is None else str(value))

    def bot_select_option(self, selector, **kwargs):
        self._check_stop()
        self.mark_activity("bot")
        self.page.select_option(selector, **kwargs)

    def bot_goto(self, url):
        self._check_stop()
        self.mark_activity("bot")
        self.page.goto(url, wait_until="domcontentloaded")


class RequestRateLimiter:
    """
    Simple rate limiter to throttle automated filter submissions.

    DIRGC mulai memunculkan HTTP 429 jika filter ditembak terlalu cepat.
    Kelas ini menjaga jarak antar request dan menambahkan exponential backoff
    ketika server sudah menolak permintaan.
    """

    def __init__(
        self,
        min_interval_s=1.2,
        penalty_initial_s=5,
        penalty_max_s=40,
        jitter_s=0.0,
        cooldown_after=0,
        cooldown_s=0.0,
    ):
        self.min_interval_s = max(0.1, float(min_interval_s))
        self.penalty_initial_s = max(0.0, float(penalty_initial_s))
        self.penalty_max_s = max(self.penalty_initial_s, float(penalty_max_s))
        self.jitter_s = max(0.0, float(jitter_s))
        self.cooldown_after = max(0, int(cooldown_after or 0))
        self.cooldown_s = max(0.0, float(cooldown_s))
        self._last_request_ts = 0.0
        self._pending_penalty_s = 0.0
        self._last_penalty_s = 0.0
        self._consecutive_rate_limits = 0

    def wait_for_slot(self, monitor):
        wait_seconds = 0.0
        if self._pending_penalty_s > 0:
            wait_seconds = self._pending_penalty_s
            self._pending_penalty_s = 0.0
        now = time.monotonic()
        if self._last_request_ts > 0:
            elapsed = now - self._last_request_ts
            if elapsed < self.min_interval_s:
                wait_seconds = max(wait_seconds, self.min_interval_s - elapsed)
        if self.jitter_s > 0:
            wait_seconds += random.uniform(0.0, self.jitter_s)
        if wait_seconds > 0:
            monitor.wait_for_condition(lambda: False, timeout_s=wait_seconds)
        self._last_request_ts = time.monotonic()

    def penalize(self):
        self._consecutive_rate_limits += 1
        if self._last_penalty_s <= 0:
            next_penalty = self.penalty_initial_s
        else:
            next_penalty = min(self._last_penalty_s * 2, self.penalty_max_s)
        self._pending_penalty_s = max(self._pending_penalty_s, next_penalty)
        self._last_penalty_s = next_penalty
        extra_cooldown = 0.0
        if self.cooldown_after and self._consecutive_rate_limits >= self.cooldown_after:
            extra_cooldown = self.cooldown_s
            if extra_cooldown > 0:
                self._pending_penalty_s = max(
                    self._pending_penalty_s, next_penalty + extra_cooldown
                )
        return next_penalty + extra_cooldown

    def reset_penalty(self):
        self._last_penalty_s = 0.0
        self._consecutive_rate_limits = 0

    def configure(
        self,
        *,
        min_interval_s=None,
        penalty_initial_s=None,
        penalty_max_s=None,
        jitter_s=None,
        cooldown_after=None,
        cooldown_s=None,
    ):
        if min_interval_s is not None:
            self.min_interval_s = max(0.1, float(min_interval_s))
        if penalty_initial_s is not None:
            self.penalty_initial_s = max(0.0, float(penalty_initial_s))
        if penalty_max_s is not None:
            self.penalty_max_s = max(
                self.penalty_initial_s, float(penalty_max_s)
            )
        if jitter_s is not None:
            self.jitter_s = max(0.0, float(jitter_s))
        if cooldown_after is not None:
            self.cooldown_after = max(0, int(cooldown_after or 0))
        if cooldown_s is not None:
            self.cooldown_s = max(0.0, float(cooldown_s))
        self.reset_penalty()


SERVER_COOLDOWN_UNTIL = 0.0
SERVER_COOLDOWN_REASON = ""


def set_server_cooldown(seconds, reason=""):
    global SERVER_COOLDOWN_UNTIL, SERVER_COOLDOWN_REASON
    try:
        seconds = float(seconds)
    except (TypeError, ValueError):
        return False
    if seconds <= 0:
        return False
    now = time.monotonic()
    until = now + seconds
    if until <= SERVER_COOLDOWN_UNTIL:
        return False
    SERVER_COOLDOWN_UNTIL = until
    SERVER_COOLDOWN_REASON = reason or ""
    log_warn(
        "Server cooldown set.",
        wait_s=round(seconds, 1),
        reason=SERVER_COOLDOWN_REASON or "-",
    )
    return True


def get_server_cooldown_remaining():
    remaining = SERVER_COOLDOWN_UNTIL - time.monotonic()
    if remaining <= 0:
        return 0.0, ""
    return remaining, SERVER_COOLDOWN_REASON


def wait_with_keepalive(
    monitor, total_s, step_s=30, log_interval_s=60, log_context=""
):
    try:
        remaining = float(total_s)
    except (TypeError, ValueError):
        return
    if remaining <= 0:
        return
    step_s = max(1.0, float(step_s))
    log_interval_s = (
        max(10.0, float(log_interval_s)) if log_interval_s else 0.0
    )
    next_log = time.monotonic()
    while remaining > 0:
        now = time.monotonic()
        if log_interval_s and now >= next_log:
            seconds_left = max(0, int(round(remaining)))
            minutes = seconds_left // 60
            seconds = seconds_left % 60
            log_warn(
                "Cooldown aktif; menunggu sebelum lanjut.",
                sisa=f"{minutes:02d}:{seconds:02d}",
                context=log_context or "-",
            )
            next_log = now + log_interval_s
        monitor.mark_activity("cooldown")
        chunk = min(step_s, remaining)
        monitor.wait_for_condition(lambda: False, timeout_s=chunk)
        remaining -= chunk
    log_info(
        "Cooldown selesai; melanjutkan proses.",
        context=log_context or "-",
    )


FILTER_RATE_LIMITER = RequestRateLimiter()
MAX_RATE_LIMIT_RETRIES = 5


SUBMIT_RATE_LIMITER = RequestRateLimiter()
SUBMIT_POST_SUCCESS_DELAY_S = 0.0


def apply_rate_limit_profile(profile_name):
    profile_key = (profile_name or "").strip().lower()
    if profile_key not in RATE_LIMIT_PROFILES:
        profile_key = DEFAULT_RATE_LIMIT_PROFILE
    profile = RATE_LIMIT_PROFILES[profile_key]

    FILTER_RATE_LIMITER.configure(
        min_interval_s=profile.get("filter_min_interval_s"),
        penalty_initial_s=profile.get("filter_penalty_initial_s"),
        penalty_max_s=profile.get("filter_penalty_max_s"),
        jitter_s=profile.get("filter_jitter_s"),
    )
    SUBMIT_RATE_LIMITER.configure(
        min_interval_s=profile.get("submit_min_interval_s"),
        penalty_initial_s=profile.get("submit_penalty_initial_s"),
        penalty_max_s=profile.get("submit_penalty_max_s"),
        jitter_s=profile.get("submit_jitter_s"),
        cooldown_after=profile.get("submit_cooldown_after"),
        cooldown_s=profile.get("submit_cooldown_s"),
    )
    global SUBMIT_POST_SUCCESS_DELAY_S
    SUBMIT_POST_SUCCESS_DELAY_S = float(
        profile.get("submit_success_delay_s") or 0.0
    )
    return profile_key


apply_rate_limit_profile(DEFAULT_RATE_LIMIT_PROFILE)


def install_user_activity_tracking(page, mark_activity):
    page.expose_function("reportActivity", lambda: mark_activity("user"))
    page.add_init_script(
        """
        (() => {
          function isRelevantInput(target) {
            if (!target) return false;
            const id = (target.id || "").toLowerCase();
            const name = (target.name || "").toLowerCase();
            const autocomplete = (target.autocomplete || "").toLowerCase();
            if (id === "username" || id === "password") return true;
            if (name === "username" || name === "password") return true;
            if (autocomplete === "one-time-code") return true;
            const markers = ["otp", "verif", "kode", "mfa"];
            return markers.some((marker) => id.includes(marker) || name.includes(marker));
          }

          function reportIfCredentialInput(event) {
            const target = event.target;
            if (!isRelevantInput(target)) return;
            if (window.reportActivity) {
              window.reportActivity();
            }
          }
          document.addEventListener("input", reportIfCredentialInput, true);
          document.addEventListener("change", reportIfCredentialInput, true);
        })();
        """
    )


def ensure_on_dirgc(
    page,
    monitor,
    use_saved_credentials,
    credentials,
):
    def is_on_target():
        return page.url.startswith(TARGET_URL)

    def is_on_login_page():
        return MATCHAPRO_HOST in page.url and LOGIN_PATH in page.url

    def is_on_matchapro():
        return MATCHAPRO_HOST in page.url

    def is_on_sso_login():
        if SSO_HOST in page.url:
            return True
        return page.locator("#kc-login").count() > 0

    def is_on_otp_challenge():
        if not is_on_sso_login():
            return False
        otp_selectors = [
            "input[autocomplete='one-time-code']",
            "input[name*='otp']",
            "input[id*='otp']",
            "input[name*='verif']",
            "input[id*='verif']",
            "input[name*='kode']",
            "input[id*='kode']",
        ]
        for selector in otp_selectors:
            locator = page.locator(selector)
            if locator.count() > 0 and locator.first.is_visible():
                return True
        text_markers = [
            "OTP",
            "Kode OTP",
            "kode otp",
            "verification code",
            "kode verifikasi",
        ]
        for marker in text_markers:
            locator = page.locator(f"text={marker}")
            if locator.count() > 0 and locator.first.is_visible():
                return True
        return False

    def click_if_present(selector):
        locator = page.locator(selector)
        if locator.count() == 0:
            return False
        monitor.bot_click(locator.first)
        return True

    def attempt_auto_login(username, password):
        if not username or not password:
            log_warn(
                "Saved credentials missing; switching to manual login."
            )
            return False

        if not monitor.wait_for_condition(
            lambda: (
                page.locator("#username").count() > 0
                or page.locator("input[name='username']").count() > 0
            ),
            30,
        ):
            log_warn("Login fields not found; switching to manual login.")
            return False

        username_locator = page.locator("#username")
        if username_locator.count() == 0:
            username_locator = page.locator("input[name='username']")
        if username_locator.count() == 0:
            log_warn("Login fields not found; switching to manual login.")
            return False
        if not username_locator.first.is_visible():
            if not monitor.wait_for_condition(
                lambda: username_locator.first.is_visible(), 15
            ):
                log_warn(
                    "Login fields not visible; switching to manual login."
                )
                return False

        monitor.bot_fill("input#username, input[name='username']", username)
        monitor.bot_fill("input#password, input[name='password']", password)
        monitor.bot_click("#kc-login")

        error_selectors = [
            "#input-error",
            "#kc-error-message",
            ".kc-feedback-text",
            ".alert-error",
            ".pf-c-alert__title",
        ]

        start = time.monotonic()
        while True:
            if is_on_matchapro():
                return True
            for selector in error_selectors:
                locator = page.locator(selector)
                if locator.count() > 0 and locator.first.is_visible():
                    return False
            if time.monotonic() - start > monitor.scale_timeout(
                AUTO_LOGIN_RESULT_TIMEOUT_S
            ):
                return False
            monitor.idle_check()
            page.wait_for_timeout(500)

    allow_autofill = use_saved_credentials
    autofill_attempted = False
    username, password = credentials or (None, None)

    if not is_on_target():
        monitor.bot_goto(TARGET_URL)

    while True:
        if is_on_target():
            log_info("On target page.", url=page.url)
            return

        if is_on_login_page():
            if click_if_present("#login-sso"):
                log_info("Redirecting to SSO login.")
                monitor.wait_for_condition(
                    lambda: is_on_sso_login() or is_on_matchapro(),
                    timeout_s=30,
                )
                continue
            monitor.wait_for_condition(
                lambda: page.locator("#login-sso").count() > 0
                or not is_on_login_page(),
                timeout_s=10,
            )
            continue

        if is_on_sso_login():
            if allow_autofill and not autofill_attempted:
                autofill_attempted = True
                if attempt_auto_login(username, password):
                    monitor.wait_for_condition(is_on_matchapro, timeout_s=60)
                    continue
                allow_autofill = False
                log_warn("Auto-fill login failed; switching to manual login.")

            if is_on_otp_challenge():
                log_info("OTP required; waiting for manual input.")
            else:
                log_info("Waiting for manual login.")
            monitor.wait_for_condition(is_on_matchapro)
            continue

        if is_on_matchapro() and not is_on_target():
            monitor.bot_goto(TARGET_URL)
            continue

        monitor.wait_for_condition(lambda: False, timeout_s=2)


def is_visible(page, selector):
    locator = page.locator(selector)
    return locator.count() > 0 and locator.first.is_visible()


def wait_for_block_ui_clear(page, monitor, timeout_s=15):
    monitor.wait_for_condition(
        lambda: page.locator(BLOCK_UI_SELECTOR).count() == 0
        or not is_visible(page, BLOCK_UI_SELECTOR),
        timeout_s=timeout_s,
    )


def ensure_filter_panel_open(page, monitor):
    if is_visible(page, "#search-idsbr"):
        return
    toggle = page.locator("#toggle-filter")
    if toggle.count() > 0:
        monitor.bot_click(toggle.first)
        monitor.wait_for_condition(
            lambda: is_visible(page, "#search-idsbr"), timeout_s=10
        )


def apply_filter(page, monitor, idsbr, nama_usaha, alamat):
    ensure_filter_panel_open(page, monitor)

    def get_results_snapshot():
        header_locator = page.locator(".usaha-card-header")
        count = header_locator.count()
        first_text = ""
        last_text = ""
        if count > 0:
            try:
                first_text = header_locator.first.inner_text().strip()
            except Exception:
                first_text = ""
            if count > 1:
                try:
                    last_text = (
                        header_locator.nth(count - 1)
                        .inner_text()
                        .strip()
                    )
                except Exception:
                    last_text = ""
            else:
                last_text = first_text
        return count, first_text, last_text

    def results_changed(previous_snapshot):
        return get_results_snapshot() != previous_snapshot

    def wait_for_results(previous_snapshot, timeout_s=15):
        monitor.wait_for_condition(
            lambda: is_visible(page, ".empty-state")
            or is_visible(page, ".no-data")
            or is_visible(page, ".no-results")
            or results_changed(previous_snapshot),
            timeout_s=timeout_s,
        )
        wait_for_block_ui_clear(page, monitor, timeout_s=timeout_s)
        return page.locator(".usaha-card-header").count()

    def retry_results_if_slow(count, timeout_s=5):
        if count <= 1:
            return count
        previous_snapshot = get_results_snapshot()
        updated = monitor.wait_for_condition(
            lambda: is_visible(page, ".empty-state")
            or is_visible(page, ".no-data")
            or is_visible(page, ".no-results")
            or results_changed(previous_snapshot),
            timeout_s=timeout_s,
        )
        if not updated:
            return count
        wait_for_block_ui_clear(page, monitor, timeout_s=timeout_s)
        return page.locator(".usaha-card-header").count()

    def set_filter_values(idsbr_value, nama_value, alamat_value):
        monitor.mark_activity("bot")
        page.evaluate(
            """
            ({ idsbrValue, namaValue, alamatValue }) => {
              const setValue = (selector, value) => {
                const input = document.querySelector(selector);
                if (!input) return;
                input.value = value || "";
              };

              setValue("#search-idsbr", idsbrValue);
              setValue("#search-nama", namaValue);
              setValue("#search-alamat", alamatValue);

              const dispatch = (selector) => {
                const input = document.querySelector(selector);
                if (!input) return;
                input.dispatchEvent(new Event("input", { bubbles: true }));
                input.dispatchEvent(new Event("change", { bubbles: true }));
              };

              dispatch("#search-idsbr");
              dispatch("#search-nama");
              dispatch("#search-alamat");
            }
            """,
            {
                "idsbrValue": idsbr_value or "",
                "namaValue": nama_value or "",
                "alamatValue": alamat_value or "",
            },
        )

    def search_with(idsbr_value, nama_value, alamat_value):
        attempts = 0

        while True:
            attempts += 1
            remaining, reason = get_server_cooldown_remaining()
            if remaining > 0:
                log_warn(
                    "Server cooldown aktif; menunda filter.",
                    wait_s=round(remaining, 1),
                    reason=reason or "-",
                )
                wait_with_keepalive(
                    monitor, remaining, log_context="filter"
                )
            FILTER_RATE_LIMITER.wait_for_slot(monitor)
            previous_snapshot = get_results_snapshot()

            def is_gc_card(resp):
                return (
                    "matchapro.web.bps.go.id/direktori-usaha/data-gc-card" in resp.url
                    and resp.request.method == "POST"
                )

            response_obj = None
            try:
                with page.expect_response(is_gc_card, timeout=5000) as resp_info:
                    set_filter_values(idsbr_value, nama_value, alamat_value)
                response_obj = resp_info.value
            except PWTimeoutError:
                response_obj = None
            except Exception:
                response_obj = None

            status_code = response_obj.status if response_obj else None
            if status_code == 429:
                wait_penalty = FILTER_RATE_LIMITER.penalize()
                log_warn(
                    "Server rate limited filter request (HTTP 429).",
                    attempt=attempts,
                    wait_s=round(wait_penalty, 1),
                )
                if attempts >= MAX_RATE_LIMIT_RETRIES:
                    raise RuntimeError(
                        "Server DIRGC terus mengembalikan HTTP 429. "
                        "Mohon beri jeda beberapa menit sebelum melanjutkan otomatisasi."
                    )
                continue

            FILTER_RATE_LIMITER.reset_penalty()
            monitor.wait_for_condition(lambda: False, timeout_s=0.5)
            return wait_for_results(previous_snapshot)

    if idsbr:
        count = search_with(idsbr, "", "")
        if count > 1:
            log_info(
                "Results not unique; rechecking for slow loading.",
                count=count,
            )
            count = retry_results_if_slow(count)
        if count == 1:
            return count
        if nama_usaha or alamat:
            if count == 0:
                log_warn(
                    "IDSBR not found; retry with idsbr + nama_usaha + alamat."
                )
            else:
                log_warn(
                    "Multiple results for IDSBR; retry with idsbr + nama_usaha + alamat.",
                    count=count,
                )
            return search_with(idsbr, nama_usaha, alamat)
        return count

    return search_with("", nama_usaha, alamat)


def _normalize_hasil_label(text):
    if text is None:
        return ""
    cleaned = " ".join(str(text).split())
    cleaned = re.sub(r"^\d+\s*[.\)\-:]\s*", "", cleaned)
    return cleaned.strip().lower()


def hasil_gc_select(page, monitor, code):
    if code is None:
        return False
    monitor.wait_for_condition(
        lambda: page.locator("#tt_hasil_gc").count() > 0, timeout_s=15
    )
    select_locator = page.locator("#tt_hasil_gc")
    try:
        options = select_locator.locator("option").evaluate_all(
            "elements => elements.map(e => ({ value: e.value, text: e.textContent || '' }))"
        )
    except Exception:
        options = []

    code_value = str(code)
    if options:
        if any(opt.get("value") == code_value for opt in options):
            monitor.bot_select_option("#tt_hasil_gc", value=code_value)
            return True
    else:
        value_locator = select_locator.locator(f'option[value="{code}"]')
        if value_locator.count() > 0:
            monitor.bot_select_option("#tt_hasil_gc", value=code_value)
            return True

    label = HASIL_GC_LABELS.get(code)
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
                if target_norm and target_norm in _normalize_hasil_label(
                    opt_text
                ):
                    matched_value = opt.get("value")
                    break
        if matched_value is not None:
            monitor.bot_select_option("#tt_hasil_gc", value=matched_value)
            return True

    if label:
        try:
            monitor.bot_select_option("#tt_hasil_gc", label=label)
            return True
        except Exception:
            return False
    return False
