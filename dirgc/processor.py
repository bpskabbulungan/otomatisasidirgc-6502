from .browser import (
    apply_filter,
    ensure_on_dirgc,
    hasil_gc_select,
    is_visible,
    wait_for_block_ui_clear,
)
from .excel import load_excel_rows
from .logging_utils import log_error, log_info, log_warn
from .matching import select_matching_card
from .run_logs import build_run_log_path, write_run_log
from .settings import TARGET_URL


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

    for offset, row in enumerate(rows):
        batch_index = offset + 1
        excel_row = start_row + offset
        stats["processed"] += 1
        status = None
        note = ""
        extra_notes = []

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
        ensure_on_dirgc(
            page,
            monitor=monitor,
            use_saved_credentials=use_saved_credentials,
            credentials=credentials,
        )

        try:
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
            monitor.bot_click(header_locator)

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

            def find_swal():
                nonlocal swal_result
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
            status = "berhasil"
            note = "Submit sukses"
        except Exception as exc:
            log_error(
                "Error while processing row.",
                idsbr=idsbr or "-",
                error=str(exc),
            )
            status = "error"
            note = str(exc)
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

    log_info("Processing completed.", _spacer=True, _divider=True, **stats)
    write_run_log(run_log_rows, run_log_path)
    log_info("Run log saved.", path=str(run_log_path))
