import re

from .browser import wait_for_block_ui_clear
from .logging_utils import log_info, log_warn
from .settings import MAX_MATCH_LOGS


def normalize_match_text(value):
    if value is None:
        return ""
    text = str(value).strip().lower()
    if not text:
        return ""
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return " ".join(text.split())


def match_tokens(value):
    normalized = normalize_match_text(value)
    if not normalized:
        return []
    tokens = normalized.split()
    filtered = []
    for token in tokens:
        if token.isdigit():
            filtered.append(token)
        elif len(token) >= 3:
            filtered.append(token)
    if not filtered:
        return [normalized]
    return filtered


def contains_tokens(haystack, tokens):
    if not tokens:
        return False
    return all(token in haystack for token in tokens)


def count_tokens_present(haystack, tokens):
    if not tokens:
        return 0
    return sum(1 for token in tokens if token in haystack)


def join_tokens(tokens):
    if not tokens:
        return "-"
    return ", ".join(tokens)


def summarize_match(index, flags, score, text, max_len=140):
    snippet = " ".join(text.split())
    if len(snippet) > max_len:
        snippet = snippet[: max_len - 3] + "..."
    return (
        f"card#{index + 1} score={score} "
        f"idsbr={'Y' if flags['idsbr'] else 'N'} "
        f"nama={'Y' if flags['nama'] else 'N'} "
        f"alamat={'Y' if flags['alamat'] else 'N'} "
        f"text='{snippet}'"
    )


def select_matching_card(page, monitor, idsbr, nama_usaha, alamat):
    header_locator = page.locator(".usaha-card-header")
    count = header_locator.count()
    if count == 0:
        return None

    wait_for_block_ui_clear(page, monitor, timeout_s=15)

    idsbr_norm = normalize_match_text(idsbr)
    nama_tokens = match_tokens(nama_usaha)
    alamat_tokens = match_tokens(alamat)

    def get_card_text(header, card_scope):
        visible_text = ""
        full_text = ""
        try:
            visible_text = card_scope.inner_text()
        except Exception:
            try:
                visible_text = header.inner_text()
            except Exception:
                visible_text = ""
        try:
            full_text = card_scope.text_content() or ""
        except Exception:
            try:
                full_text = header.text_content() or ""
            except Exception:
                full_text = ""
        if not full_text:
            full_text = visible_text
        if not visible_text:
            visible_text = full_text
        return visible_text or "", full_text or ""

    candidates = []
    for idx in range(count):
        header = header_locator.nth(idx)
        card_scope = header.locator(
            "xpath=ancestor::*[contains(@class, 'usaha-card')]"
        )
        if card_scope.count() == 0:
            card_scope = header

        visible_text, full_text = get_card_text(header, card_scope)
        haystack = normalize_match_text(full_text)
        flags = {
            "idsbr": bool(idsbr_norm and idsbr_norm in haystack),
            "nama": contains_tokens(haystack, nama_tokens),
            "alamat": contains_tokens(haystack, alamat_tokens),
        }
        nama_hits = count_tokens_present(haystack, nama_tokens)
        alamat_hits = count_tokens_present(haystack, alamat_tokens)
        score = 0
        if flags["idsbr"]:
            score += 3
        if flags["nama"]:
            score += 2
        if flags["alamat"]:
            score += 1

        candidates.append(
            {
                "header": header,
                "card": card_scope,
                "flags": flags,
                "score": score,
                "text": visible_text,
                "nama_hits": nama_hits,
                "alamat_hits": alamat_hits,
            }
        )

    def is_acceptable(flags):
        if idsbr and flags["idsbr"]:
            return True
        if nama_usaha and flags["nama"]:
            if alamat:
                return flags["alamat"] or flags["idsbr"]
            return True
        if not nama_usaha and alamat and flags["alamat"]:
            return True
        return False

    if count == 1:
        candidate = candidates[0]
        if is_acceptable(candidate["flags"]):
            log_info("Single result matched.")
            log_info(
                summarize_match(
                    0,
                    candidate["flags"],
                    candidate["score"],
                    candidate["text"],
                )
            )
            return candidate["header"], candidate["card"]

        soft_ok = False
        if idsbr and candidate["flags"]["idsbr"]:
            soft_ok = True
        elif nama_usaha and candidate["nama_hits"] > 0:
            if alamat:
                soft_ok = (
                    candidate["alamat_hits"] > 0
                    or candidate["flags"]["idsbr"]
                )
            else:
                soft_ok = True
        elif not nama_usaha and alamat and candidate["alamat_hits"] > 0:
            soft_ok = True
        if soft_ok:
            log_info(
                "Single result matched (relaxed).",
                nama_hits=f"{candidate['nama_hits']}/{len(nama_tokens)}",
                alamat_hits=f"{candidate['alamat_hits']}/{len(alamat_tokens)}",
            )
            log_info(
                summarize_match(
                    0,
                    candidate["flags"],
                    candidate["score"],
                    candidate["text"],
                )
            )
            return candidate["header"], candidate["card"]
        log_warn(
            "Single result mismatch; skipping.",
            idsbr=idsbr_norm or "-",
            nama_tokens=join_tokens(nama_tokens),
            alamat_tokens=join_tokens(alamat_tokens),
        )
        log_info(
            summarize_match(
                0,
                candidate["flags"],
                candidate["score"],
                candidate["text"],
            )
        )
        return None

    idsbr_matches = [c for c in candidates if c["flags"]["idsbr"]]
    if idsbr_matches:
        candidates = idsbr_matches

    candidates.sort(key=lambda c: c["score"], reverse=True)
    if not candidates or candidates[0]["score"] == 0:
        log_warn(
            "No matching result found; skipping.",
            idsbr=idsbr_norm or "-",
            nama_tokens=join_tokens(nama_tokens),
            alamat_tokens=join_tokens(alamat_tokens),
        )
        for idx, candidate in enumerate(candidates[:MAX_MATCH_LOGS]):
            log_info(
                summarize_match(
                    idx,
                    candidate["flags"],
                    candidate["score"],
                    candidate["text"],
                )
            )
        return None
    if len(candidates) > 1 and candidates[0]["score"] == candidates[1]["score"]:
        log_warn("Ambiguous match (multiple results); skipping.")
        log_info(
            summarize_match(
                0,
                candidates[0]["flags"],
                candidates[0]["score"],
                candidates[0]["text"],
            )
        )
        log_info(
            summarize_match(
                1,
                candidates[1]["flags"],
                candidates[1]["score"],
                candidates[1]["text"],
            )
        )
        return None

    best = candidates[0]
    if not is_acceptable(best["flags"]):
        log_warn(
            "Best match failed validation; skipping.",
            idsbr=idsbr_norm or "-",
            nama_tokens=join_tokens(nama_tokens),
            alamat_tokens=join_tokens(alamat_tokens),
        )
        log_info(
            summarize_match(
                0,
                best["flags"],
                best["score"],
                best["text"],
            )
        )
        return None

    log_info("Best match selected.")
    log_info(
        summarize_match(
            0,
            best["flags"],
            best["score"],
            best["text"],
        )
    )
    return best["header"], best["card"]
