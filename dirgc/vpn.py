import os
import re
import socket
import subprocess

from .logging_utils import log_info, log_warn
from .settings import DEFAULT_VPN_PREFIXES


def _load_prefixes(prefixes):
    if prefixes:
        return tuple(p.strip() for p in prefixes if p and str(p).strip())
    env_value = os.getenv("DIRGC_VPN_PREFIXES")
    if env_value:
        parts = [part.strip() for part in env_value.split(",")]
        return tuple(part for part in parts if part)
    return tuple(DEFAULT_VPN_PREFIXES)


def _get_ipv4_from_psutil():
    try:
        import psutil
    except Exception:
        return []
    addresses = []
    try:
        for _, addrs in psutil.net_if_addrs().items():
            for addr in addrs:
                if getattr(addr, "family", None) == socket.AF_INET:
                    if addr.address:
                        addresses.append(addr.address)
    except Exception:
        return []
    return addresses


def _get_ipv4_from_ipconfig():
    try:
        output = subprocess.check_output(
            ["ipconfig"], text=True, errors="ignore"
        )
    except Exception:
        return []
    matches = re.findall(
        r"(?:IPv4 Address|Alamat IPv4)[^:]*:\s*([0-9.]+)", output
    )
    return [match.strip() for match in matches if match.strip()]


def get_ipv4_addresses():
    addresses = _get_ipv4_from_psutil()
    if addresses:
        return addresses
    return _get_ipv4_from_ipconfig()


def is_vpn_connected(prefixes=None):
    prefixes = _load_prefixes(prefixes)
    if not prefixes:
        return False
    addresses = get_ipv4_addresses()
    for addr in addresses:
        for prefix in prefixes:
            if addr.startswith(prefix):
                return True
    return False


def ensure_vpn_connected(prefixes=None):
    prefixes = _load_prefixes(prefixes)
    if is_vpn_connected(prefixes):
        log_info("VPN terdeteksi.", prefixes=",".join(prefixes))
        return True
    log_warn(
        "VPN tidak terdeteksi; hentikan proses.",
        prefixes=",".join(prefixes) if prefixes else "-",
    )
    raise RuntimeError(
        "VPN tidak terdeteksi. Aktifkan VPN lalu jalankan ulang."
    )
