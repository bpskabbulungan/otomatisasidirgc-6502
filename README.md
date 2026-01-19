# Otomatisasi Ground Check Direktori SBR

## Ringkasan

Aplikasi CLI berbasis Playwright untuk membantu otomatisasi isian field GC berdasarkan data Excel, dilengkapi log untuk pemantauan. Tersedia GUI berbasis PyQt5 + QFluentWidgets untuk pengguna non-terminal.

Jika ingin langsung menggunakan versi Installer GUI, lihat [packaging](https://github.com/bpskabbulungan/otomatisasidirgc-6502/tree/main/packaging). Tapi, sebaiknya baca dokumentasinya dulu secara lengkap dan cermat sampai selesai bair lebih memahami alur kerjanya.

## Daftar Isi

- [Ringkasan](#ringkasan)
- [Ringkasan Fitur](#ringkasan-fitur)
- [Struktur Folder](#struktur-folder)
- [Prasyarat](#prasyarat)
- [Inisiasi](#inisiasi)
- [Requirements](#requirements)
- [Konfigurasi Akun SSO](#konfigurasi-akun-sso)
- [File Excel](#file-excel)
- [Cara Menjalankan](#cara-menjalankan)
- [Catatan](#catatan)
- [Output Log Excel](#output-log-excel)
- [Kredit](#kredit)

## Ringkasan Fitur

- Auto-login (bila kredensial tersedia) dan fallback ke login manual/OTP.
- Filter dan pemilihan usaha berdasarkan IDSBR/nama/alamat.
- Pengisian Hasil GC, latitude, dan longitude dengan validasi sederhana.
- Opsi pemrosesan parsial lewat `--start` dan `--end`.
- Log terminal terstruktur dan file log Excel per run.
- GUI modern untuk menjalankan proses tanpa terminal.

## Struktur Folder

```text
.
|- dirgc/                 # Modul utama aplikasi
|  `- gui/                # GUI (PyQt5 + QFluentWidgets)
|- config/                # Konfigurasi lokal (contoh: credentials)
|- data/                  # File input (Excel)
|- logs/                  # Output log per run (Excel)
|- run_dirgc.py           # Entry point CLI (wrapper)
|- run_dirgc_gui.py       # Entry point GUI
|- requirements.txt
`- README.md
```

## Prasyarat

- Python 3 sudah terpasang.
- Akun dengan akses ke DIRGC (MatchaPro).
- File Excel input tersedia di `data/` atau ditentukan lewat `--excel-file`.

## Inisiasi

Buka terminal CMD atau PowerShell, clone project dan masuk ke foldernya:

```bash
git clone https://github.com/bpskabbulungan/otomatisasidirgc-6502.git
cd otomatisasidirgc-6502
```

## Requirements

Install dependensi:

```bash
pip install -r requirements.txt
playwright install chromium
```

Catatan: `playwright install chromium` cukup dijalankan sekali per environment.

## Konfigurasi Akun SSO

Untuk CLI, letakkan file JSON di `config/credentials.json`:

```json
{
  "username": "usernamesso",
  "password": "passwordsso"
}
```

Atau gunakan environment variables:

- `DIRGC_USERNAME`
- `DIRGC_PASSWORD`

Pencarian file kredensial juga mendukung fallback `credentials.json` di root project.
Jika file dan environment variables tersedia, isi file akan diprioritaskan.
Untuk GUI, isi kredensial lewat menu `Akun SSO` (tidak disimpan ke file).

## File Excel

Default: `data/Direktori_SBR_20260114.xlsx` (bisa diganti via `--excel-file`).
Jika tidak ditemukan, sistem akan mencoba `Direktori_SBR_20260114.xlsx` di root project.

Kolom yang dikenali:

- `idsbr`
- `nama_usaha` (atau `nama usaha` / `namausaha` / `nama`)
- `alamat` (atau `alamat usaha` / `alamat_usaha`)
- `latitude` / `lat`
- `longitude` / `lon` / `long`
- `hasil_gc` / `hasil gc` / `hasilgc` / `ag` / `keberadaanusaha_gc`

Kode `hasil_gc` yang valid:

- 0 = Tidak Ditemukan
- 1 = Ditemukan
- 3 = Tutup
- 4 = Ganda

Jika kolom `hasil_gc` tidak ditemukan, sistem memakai kolom ke-6 (`keberadaanusaha_gc`).

## Cara Menjalankan

GUI (direkomendasikan untuk pengguna non-terminal):

```bash
python run_dirgc_gui.py
```

Di GUI, buka menu `Akun SSO` untuk mengisi username dan password jika ingin auto-login.

CLI:

```bash
python run_dirgc.py
```

Opsi CLI penting:

```bash
python run_dirgc.py --excel-file data/Direktori_SBR_20260114.xlsx --credentials-file config/credentials.json --manual-only --keep-open --start 1 --end 5
```

Gunakan `--start` dan `--end` untuk membatasi baris yang diproses (1-based, inklusif).

Opsi tambahan:

- `--headless` untuk menjalankan browser tanpa UI (SSO sering butuh mode non-headless).
- `--idle-timeout-ms` untuk batas idle (default 300000 / 5 menit).
- `--web-timeout-s` untuk toleransi loading web (default 30 detik).
- `--manual-only` untuk selalu login manual (tanpa auto-fill kredensial).

Auto-login akan mencoba kredensial terlebih dulu; jika gagal/OTP muncul, akan beralih ke manual login.

## Catatan

- Untuk login SSO, mode non-headless disarankan.
- Log terminal sudah diperkaya dengan timestamp dan detail langkah.

## Output Log Excel

Setiap run akan menghasilkan file log Excel di folder `logs/YYYYMMDD/`.
Nama file mengikuti pola `run{N}_{HHMM}.xlsx` (contoh: `run1_0930.xlsx`).

Kolom log:

- `no`
- `idsbr`
- `nama_usaha`
- `alamat`
- `keberadaanusaha_gc`
- `latitude`
- `longitude`
- `status` (berhasil/gagal/error/skipped)
- `catatan`

Nilai `skipped` biasanya muncul jika data sudah GC atau terdeteksi duplikat.

## Kredit

Semoga panduan ini membantu. Jika ada pertanyaan, hubungi tim IPDS BPS Kabupaten Bulungan.
