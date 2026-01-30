# Otomatisasi Ground Check Direktori SBR

## Ringkasan

Aplikasi CLI berbasis Playwright untuk membantu otomatisasi isian field GC berdasarkan data Excel, dilengkapi log untuk pemantauan. Tersedia GUI berbasis PyQt5 + QFluentWidgets untuk pengguna non-terminal.

Tersedia dua versi: GUI dan Script atau Terminal. Alur singkatnya: mulai dari [Persiapan (Python dan Git)](#persiapan-python-dan-git), lanjut ke [Unduh Project](#unduh-project) dan [Instal Dependensi](#instal-dependensi), siapkan [File Excel](#file-excel), lalu jalankan via [GUI](#cara-menjalankan---gui) atau [Script atau Terminal](#cara-menjalankan---script-atau-terminal), jangan lupa [Konfigurasi Akun SSO](#konfigurasi-akun-sso) jika memilih via terminal.

Jika ingin langsung menggunakan versi Installer GUI, lihat [packaging](https://github.com/bpskabbulungan/otomatisasidirgc-6502/tree/main/packaging). Tersedia juga versi .exe yang siap pakai, lihat di https://drive.google.com/file/d/1rewDhHUY_tDCxnPDB52lAVMijJvz1sDw/view?usp=sharing. 

Namun, sebaiknya baca dokumentasi ini dulu agar alur kerjanya lebih jelas.

## Daftar Isi

- [Ringkasan](#ringkasan)
- [Struktur Folder](#struktur-folder)
- [Persiapan (Python dan Git)](#persiapan-python-dan-git)
- [Unduh Project](#unduh-project)
- [Instal Dependensi](#instal-dependensi)
- [File Excel](#file-excel)
- [Cara Menjalankan - GUI](#cara-menjalankan---gui)
- [Cara Menjalankan - Script atau Terminal](#cara-menjalankan---script-atau-terminal)
- [Konfigurasi Akun SSO](#konfigurasi-akun-sso)
- [Catatan](#catatan)
- [Output Log Excel](#output-log-excel)
- [Kredit](#kredit)

## Struktur Folder

```text
.
|- dirgc/                 # Modul utama aplikasi
|  |- browser.py          # Util Playwright: login/redirect DIRGC, filter, rate limit, pilih hasil GC
|  |- cli.py              # Parsing argumen CLI + orkestrasi run (Playwright -> proses Excel)
|  |- credentials.py      # Load kredensial dari env/JSON + fallback path
|  |- excel.py            # Baca Excel + normalisasi kolom/nilai (idsbr, nama, alamat, koordinat, hasil_gc)
|  |- logging_utils.py    # Logger konsol + formatter (juga untuk GUI via handler)
|  |- matching.py         # Pencocokan usaha dari hasil filter (token match + scoring)
|  |- processor.py        # Proses utama per baris Excel: filter, pilih usaha, isi form, submit, log
|  |- run_logs.py         # Generate path & tulis file log Excel (logs/YYYYMMDD/runN_HHMM.xlsx)
|  |- settings.py         # Konstanta & default config (URL, timeout, file default, dsb)
|  `- gui/                # GUI (PyQt5 + QFluentWidgets)
|     `- app.py           # Seluruh UI: halaman Run/Update/SSO/Settings + worker thread
|- config/                # Konfigurasi lokal (contoh: credentials)
|- data/                  # File input (Excel)
|- logs/                  # Output log per run (Excel)
|- run_dirgc.py           # Entry point CLI (wrapper -> dirgc.cli.main)
|- run_dirgc_gui.py       # Entry point GUI (wrapper -> dirgc.gui.app.main)
|- requirements.txt
`- README.md
```

## Persiapan (Python dan Git)

Bagian ini untuk pengguna awam. Jika Python dan Git sudah terpasang, lanjut ke [Unduh Project](#unduh-project).

### 1) Install Python

Unduh dan instal Python 3 dari:
https://www.python.org/downloads/

Saat instalasi di Windows, centang opsi "Add Python to PATH", lalu lanjutkan sampai selesai.

### 2) Install Git

Unduh dan instal Git dari:
https://git-scm.com/downloads

### 3) Cek instalasi

Buka PowerShell atau CMD, lalu jalankan:

```bash
python --version
```

Contoh output yang benar:
`Python 3.11.6`

Lalu cek Git:

```bash
git --version
```

Contoh output yang benar:
`git version 2.44.0.windows.1`

Jika `python` tidak dikenali, tutup dan buka ulang terminal. Jika masih gagal, ulangi instalasi Python dan pastikan opsi "Add Python to PATH" dicentang.

## Unduh Project

Buka PowerShell atau CMD, lalu jalankan perintah berikut satu per satu:

```bash
git clone https://github.com/bpskabbulungan/otomatisasidirgc-6502.git
```

Perintah di atas mengunduh project dari GitHub.

```bash
cd otomatisasidirgc-6502
```

Perintah di atas masuk ke folder project (wajib sebelum menjalankan perintah lain).

## Instal Dependensi

Pastikan masih berada di folder project `otomatisasidirgc-6502`, lalu jalankan:

```bash
python -m pip install -r requirements.txt
```

Perintah di atas menginstal semua library Python yang dibutuhkan.

```bash
python -m playwright install chromium
```

Perintah di atas mengunduh browser Chromium yang diperlukan Playwright. Cukup dijalankan sekali per environment.

## File Excel

Default: `data/Direktori_SBR_20260114.xlsx` (bisa diganti via `--excel-file`).
Jika file tidak ditemukan, sistem akan mencoba `Direktori_SBR_20260114.xlsx` di root project.

Kolom yang dikenali:

- `idsbr`
- `nama_usaha` (atau `nama usaha` / `namausaha` / `nama`)
- `alamat` (atau `alamat usaha` / `alamat_usaha`)
- `latitude` / `lat`
- `longitude` / `lon` / `long`
- `hasil_gc` / `hasil gc` / `hasilgc` / `ag` / `keberadaanusaha_gc`

Kode `hasil_gc` yang valid:

- 0 / 99 = Tidak Ditemukan
- 1 = Ditemukan
- 3 = Tutup
- 4 = Ganda

Jika kolom `hasil_gc` tidak ditemukan, sistem memakai kolom ke-6 (`keberadaanusaha_gc`).

## Cara Menjalankan - GUI

GUI direkomendasikan untuk pengguna non-terminal. Pastikan semua langkah di atas sudah dilakukan.

Jalankan perintah berikut dari folder project:

```bash
python run_dirgc_gui.py
```

Setelah GUI terbuka:

1. Buka menu `Akun SSO`, isi username dan password jika ingin auto-login.
2. Pilih file Excel (atau pastikan file default sudah ada di `data/`).
3. Buka menu `Run` untuk input baru, atau menu `Update` untuk memperbarui data.
4. Klik tombol mulai/update sesuai menu yang dipilih.
5. Di menu `Update`, pilih field yang ingin diperbarui (Hasil GC, Nama, Alamat, Koordinat).
   Jika field dipilih tetapi nilai Excel kosong, baris akan ditolak (status `gagal`).
   Untuk koordinat, boleh isi salah satu saja (latitude atau longitude).
6. Jika sering muncul pesan *Something Went Wrong* saat submit, buka menu
   `Anti Rate Limit` dan pilih mode agar jeda antar submit lebih panjang dan 429 lebih jarang muncul.

## Cara Menjalankan - Script atau Terminal

### Perintah dasar

Jalankan perintah berikut dari folder project:

```bash
python run_dirgc.py
```

Perintah ini akan menggunakan file Excel default di `data/` dan mencoba auto-login jika kredensial tersedia.

### Menentukan file Excel dan kredensial

```bash
python run_dirgc.py --excel-file data/Direktori_SBR_20260114.xlsx --credentials-file config/credentials.json
```

Perintah di atas memakai file Excel tertentu dan kredensial dari file JSON.

### Membatasi baris yang diproses

```bash
python run_dirgc.py --start 1 --end 5
```

Perintah di atas hanya memproses baris 1 sampai 5 (1-based, inklusif).

### Opsi CLI tambahan

- `--headless` untuk menjalankan browser tanpa UI (SSO sering butuh mode non-headless).
- `--idle-timeout-ms` untuk batas idle (default 300000 / 5 menit).
- `--web-timeout-s` untuk toleransi loading web (default 30 detik).
- `--manual-only` untuk selalu login manual (tanpa auto-fill kredensial).
- `--dirgc-only` untuk berhenti di halaman DIRGC (tanpa filter/input).
- `--edit-nama-alamat` untuk mengaktifkan toggle edit Nama/Alamat Usaha dan isi dari Excel.
- `--keep-open` untuk menahan browser tetap terbuka setelah proses.
- `--update-mode` untuk menggunakan tombol Edit Hasil (update data).
- `--prefer-web-coords` untuk mempertahankan koordinat yang sudah terisi di web.
- `--update-fields` untuk memilih field yang di-update (contoh: `hasil_gc,nama_usaha,alamat,koordinat`).
- `--rate-limit-profile` untuk mengatur kecepatan submit (normal/safe/ultra).

Auto-login akan mencoba kredensial terlebih dulu; jika gagal/OTP muncul, akan beralih ke manual login.
Secara default, koordinat diisi dari Excel (jika ada), meskipun web sudah berisi.

## Konfigurasi Akun SSO

Untuk CLI, buat file `config/credentials.json` dengan isi berikut:

```json
{
  "username": "usernamesso",
  "password": "passwordsso"
}
```

Atau gunakan environment variables:

- `DIRGC_USERNAME`
- `DIRGC_PASSWORD`

Pencarian file kredensial juga mendukung fallback `credentials.json` di root project. Jika file dan environment variables tersedia, isi file akan diprioritaskan.

Untuk GUI, isi kredensial lewat menu `Akun SSO` (tidak disimpan ke file).

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
- `latitude_source` (web/excel/empty/missing/unknown)
- `latitude_before`
- `latitude_after`
- `longitude`
- `longitude_source` (web/excel/empty/missing/unknown)
- `longitude_before`
- `longitude_after`
- `hasil_gc_before`
- `hasil_gc_after`
- `nama_usaha_before`
- `nama_usaha_after`
- `alamat_before`
- `alamat_after`
- `status` (berhasil/gagal/error/skipped)
- `catatan`

Nilai `skipped` biasanya muncul jika data sudah GC atau terdeteksi duplikat.

## Kredit

Semoga panduan ini membantu. Jika ada pertanyaan, hubungi tim IPDS BPS Kabupaten Bulungan.
