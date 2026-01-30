import time

from .browser import (
    SUBMIT_POST_SUCCESS_DELAY_S,
    SUBMIT_RATE_LIMITER,
    apply_filter,
    ensure_on_dirgc,
    get_server_cooldown_remaining,
    hasil_gc_select,
    is_visible,
    wait_with_keepalive,
    wait_for_block_ui_clear,
)
from .excel import load_excel_rows
from .logging_utils import log_error, log_info, log_warn
from .matching import select_matching_card
from .resume_state import save_resume_state
from .run_logs import build_run_log_path, write_run_log
from .settings import RUN_LOG_CHECKPOINT_EVERY, TARGET_URL


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


def process_excel_rows(
    page,
    monitor,
    excel_file,
    use_saved_credentials,
    credentials,
    edit_nama_alamat=False,
    prefer_excel_coords=True,
    update_mode=False,
    update_fields=None,
    start_row=None,
    end_row=None,
    progress_callback=None,
    stop_on_cooldown=False,
):
    run_log_path = build_run_log_path()
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
    last_checkpoint = 0

    def checkpoint_run_log(force=False):
        nonlocal last_checkpoint
        if not run_log_rows:
            return
        if checkpoint_every <= 0 and not force:
            return
        if (
            checkpoint_every > 0
            and not force
            and len(run_log_rows) - last_checkpoint < checkpoint_every
        ):
            return
        write_run_log(run_log_rows, run_log_path)
        last_checkpoint = len(run_log_rows)
        if not force:
            log_info(
                "Run log checkpoint saved.",
                rows=last_checkpoint,
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

        update_fields_set = None
        if update_mode:
            if isinstance(update_fields, (list, tuple, set)):
                update_fields_set = {
                    str(item).strip().lower()
                    for item in update_fields
                    if str(item).strip()
                }
            elif update_fields is not None:
                update_fields_set = set()
        update_hasil_gc = (
            not update_mode
            or update_fields_set is None
            or "hasil_gc" in update_fields_set
        )
        update_nama = (
            not update_mode
            or update_fields_set is None
            or "nama_usaha" in update_fields_set
        )
        update_alamat = (
            not update_mode
            or update_fields_set is None
            or "alamat" in update_fields_set
        )
        update_lat = (
            not update_mode
            or update_fields_set is None
            or "latitude" in update_fields_set
            or "koordinat" in update_fields_set
        )
        update_lon = (
            not update_mode
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

        try:
            monitor.idle_check()
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
            if update_mode:
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
                        missing_fields.append("koordinat")
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
                        "Update ditolak: field kosong di Excel.",
                        idsbr=idsbr or "-",
                        fields=missing_label,
                    )
                    status = "gagal"
                    note = f"Update ditolak: field kosong ({missing_label})"
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
            try:
                header_locator.scroll_into_view_if_needed()
            except Exception:
                pass
            try:
                monitor.bot_click(header_locator)
            except Exception as exc:
                if _is_swal_overlay_visible(page):
                    dismissed = dismiss_swal_overlays(
                        page, monitor, context="select card retry", timeout_s=8
                    )
                    if dismissed:
                        monitor.bot_click(header_locator)
                    else:
                        raise
                elif _is_bootstrap_modal_visible(page):
                    dismissed = dismiss_bootstrap_modals(
                        page, monitor, context="select card retry", timeout_s=8
                    )
                    if dismissed:
                        monitor.bot_click(header_locator)
                    else:
                        raise
                else:
                    raise

            if card_scope.count() == 0:
                card_scope = page

            if not update_mode:
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

            action_selector = ".btn-gc-edit" if update_mode else ".btn-tandai"
            action_label = "Edit Hasil" if update_mode else "Tandai"
            action_locator = page.locator(action_selector)
            if action_locator.count() == 0:
                missing_note = (
                    "Tombol Edit Hasil tidak ditemukan (kemungkinan belum GC; gunakan menu Run)."
                    if update_mode
                    else "Tombol Tandai tidak ditemukan."
                )
                log_warn(
                    f"Tombol {action_label} tidak ditemukan; skipping.",
                    idsbr=idsbr or "-",
                )
                stats["skipped_no_tandai"] += 1
                status = "gagal"
                note = missing_note
                continue
            if not action_locator.first.is_visible():
                log_warn(
                    f"Tombol {action_label} tidak terlihat; skipping.",
                    idsbr=idsbr or "-",
                )
                stats["skipped_no_tandai"] += 1
                status = "gagal"
                note = f"Tombol {action_label} tidak terlihat"
                continue

            wait_for_block_ui_clear(page, monitor, timeout_s=15)
            try:
                action_locator.first.scroll_into_view_if_needed()
            except Exception:
                pass
            try:
                monitor.bot_click(action_locator.first)
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
                lambda: page.locator("#tt_hasil_gc").count() > 0,
                timeout_s=30,
            )
            if not form_ready:
                log_warn(
                    "Form Hasil GC tidak muncul; skipping.",
                    idsbr=idsbr or "-",
                )
                stats["skipped_no_tandai"] += 1
                status = "gagal"
                note = "Form Hasil GC tidak muncul"
                continue

            if update_hasil_gc:
                select_locator = page.locator("#tt_hasil_gc")
                try:
                    hasil_gc_before = select_locator.first.input_value() or ""
                except Exception:
                    hasil_gc_before = ""
                if hasil_gc_select(page, monitor, hasil_gc):
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
                    note = "Hasil GC tidak valid/kosong"
            else:
                select_locator = page.locator("#tt_hasil_gc")
                try:
                    hasil_gc_before = select_locator.first.input_value() or ""
                except Exception:
                    hasil_gc_before = ""
                hasil_gc_after = hasil_gc_before

            def safe_fill(selector, value, field_name, allow_overwrite):
                locator = page.locator(selector)
                if locator.count() == 0 or not locator.first.is_visible():
                    log_warn(
                        "Field tidak ditemukan; lewati.",
                        idsbr=idsbr or "-",
                        field=field_name,
                    )
                    return "", "missing", "", ""
                try:
                    current_value = locator.first.input_value()
                except Exception:
                    current_value = ""
                current_value = (current_value or "").strip()
                desired_value = str(value).strip() if value is not None else ""

                if allow_overwrite:
                    if desired_value:
                        if current_value != desired_value:
                            monitor.bot_fill(selector, desired_value)
                        return desired_value, "excel", current_value, desired_value
                    if current_value:
                        return current_value, "web", current_value, current_value
                    return "", "empty", current_value, ""

                if current_value:
                    return current_value, "web", current_value, current_value
                if not desired_value:
                    return "", "empty", current_value, ""
                monitor.bot_fill(selector, desired_value)
                return desired_value, "excel", current_value, desired_value

            lat_value = latitude if update_lat else None
            lon_value = longitude if update_lon else None
            allow_overwrite_lat = prefer_excel_coords if update_lat else False
            allow_overwrite_lon = prefer_excel_coords if update_lon else False
            latitude_value, latitude_source, latitude_before, latitude_after = (
                safe_fill(
                    "#tt_latitude_cek_user",
                    lat_value,
                    "latitude",
                    allow_overwrite_lat,
                )
            )
            longitude_value, longitude_source, longitude_before, longitude_after = (
                safe_fill(
                    "#tt_longitude_cek_user",
                    lon_value,
                    "longitude",
                    allow_overwrite_lon,
                )
            )

            def ensure_edit_field(toggle_selector, input_selector, value, field_name):
                desired_value = str(value).strip() if value else ""
                toggle = page.locator(toggle_selector)
                if toggle.count() == 0 or not toggle.first.is_visible():
                    log_warn(
                        "Toggle edit tidak ditemukan; lewati.",
                        idsbr=idsbr or "-",
                        field=field_name,
                    )
                    return "", ""
                try:
                    toggle_checked = toggle.first.is_checked()
                except Exception:
                    toggle_checked = False
                if not toggle_checked:
                    try:
                        monitor.bot_click(toggle.first)
                    except Exception as exc:
                        log_warn(
                            "Toggle edit gagal diklik; lewati.",
                            idsbr=idsbr or "-",
                            field=field_name,
                            error=str(exc),
                        )
                        return "", ""

                input_locator = page.locator(input_selector)
                if (
                    input_locator.count() == 0
                    or not input_locator.first.is_visible()
                ):
                    log_warn(
                        "Field edit tidak ditemukan; lewati.",
                        idsbr=idsbr or "-",
                        field=field_name,
                    )
                    return "", ""
                if not monitor.wait_for_condition(
                    lambda: input_locator.count() > 0
                    and input_locator.first.is_editable(),
                    timeout_s=5,
                ):
                    log_warn(
                        "Field edit tidak bisa diedit; lewati.",
                        idsbr=idsbr or "-",
                        field=field_name,
                    )
                    return "", ""
                try:
                    current_value = input_locator.first.input_value()
                except Exception:
                    current_value = ""
                current_value = (current_value or "").strip()
                if not desired_value:
                    return current_value, current_value
                if current_value == desired_value:
                    return current_value, desired_value
                monitor.bot_fill(input_selector, desired_value)
                return current_value, desired_value

            if update_mode:
                if update_nama:
                    nama_before, nama_after = ensure_edit_field(
                        "#toggle_edit_nama",
                        "#tt_nama_usaha_gc",
                        nama_usaha,
                        "nama_usaha",
                    )
                if update_alamat:
                    alamat_before, alamat_after = ensure_edit_field(
                        "#toggle_edit_alamat",
                        "#tt_alamat_usaha_gc",
                        alamat,
                        "alamat",
                    )
            elif edit_nama_alamat:
                ensure_edit_field(
                    "#toggle_edit_nama",
                    "#tt_nama_usaha_gc",
                    nama_usaha,
                    "nama_usaha",
                )
                ensure_edit_field(
                    "#toggle_edit_alamat",
                    "#tt_alamat_usaha_gc",
                    alamat,
                    "alamat",
                )

            if status == "gagal" and note == "Hasil GC tidak valid/kosong":
                monitor.bot_goto(TARGET_URL)
                continue

            submit_locator = page.locator("#save-tandai-usaha-btn")
            if submit_locator.count() == 0:
                log_warn(
                    "Tombol submit tidak ditemukan; skipping.",
                    idsbr=idsbr or "-",
                )
                status = "gagal"
                note = "Tombol submit tidak ditemukan"
                monitor.bot_goto(TARGET_URL)
                continue
            if not submit_locator.first.is_visible():
                log_warn(
                    "Tombol submit tidak terlihat; skipping.",
                    idsbr=idsbr or "-",
                )
                status = "gagal"
                note = "Tombol submit tidak terlihat"
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
                submit_locator.first.scroll_into_view_if_needed()
            except Exception:
                pass
            try:
                monitor.bot_click(submit_locator.first)
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
                if not latitude_value and not longitude_value:
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
            note = "Submit sukses"
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
