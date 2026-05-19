# LMS Scrapper

Project ini adalah tool Python untuk mengambil data nilai peserta dari LMS KitaKerja secara otomatis. Scraper membuka halaman grade report LMS, membaca tabel nilai per peserta, menyesuaikan nama kolom nilai ke format yang sudah ditentukan, memvalidasi peserta terhadap file master, lalu menyimpan hasilnya ke file Excel.

Saat ini konfigurasi default diarahkan untuk kelas B:

- LMS course ID: `48`
- File master peserta: `master_peserta_b.csv`
- File output: `nilai_tugas_lms_kelas_b.xlsx`

## Struktur File

```text
.
├── login_lms.py                 # Login manual LMS dan simpan session browser
├── scrape_lms_playwright.py     # Script utama scraping nilai LMS
├── req.txt                      # Daftar dependency Python
├── master_peserta_b.csv         # Master data peserta kelas B
├── nilai_tugas_lms_kelas_b.xlsx # Contoh/hasil output scraping kelas B
├── storage_state.json           # Session login Playwright, dibuat oleh login_lms.py
└── output/                      # Folder output tambahan jika dibutuhkan
```

File `.env`, `.venv`, `storage_state.json`, dan folder `output/` sudah diabaikan oleh Git melalui `.gitignore`.

## Prasyarat

- Python 3.9 atau lebih baru
- Akses akun LMS KitaKerja
- Browser Chromium untuk Playwright
- File master peserta dalam format CSV dengan kolom:
  - `Nama`
  - `Kelas`
  - `Email`

Contoh format master peserta:

```csv
Nama,Kelas,Email
Nama Peserta,B,emailpeserta@example.com
```

## Instalasi

1. Buat virtual environment:

   ```bash
   python -m venv .venv
   ```

2. Aktifkan virtual environment:

   ```bash
   source .venv/bin/activate
   ```

3. Install dependency:

   ```bash
   pip install -r req.txt
   ```

4. Install browser Playwright:

   ```bash
   playwright install chromium
   ```

## Login LMS

Scraper memakai session browser yang disimpan di `storage_state.json`. Session ini dibuat sekali melalui login manual.

Jalankan:

```bash
python login_lms.py
```

Setelah browser terbuka:

1. Login ke LMS seperti biasa.
2. Pastikan sudah berhasil masuk ke dashboard LMS.
3. Kembali ke terminal.
4. Tekan `ENTER`.

Jika berhasil, terminal akan menampilkan pesan bahwa session berhasil disimpan ke `storage_state.json`.

Jalankan ulang `login_lms.py` jika:

- Session LMS sudah expired.
- Password akun berubah.
- Scraper menampilkan error bahwa session belum login.
- File `storage_state.json` terhapus.

## Menjalankan Scraper

Pastikan virtual environment aktif, lalu jalankan:

```bash
python scrape_lms_playwright.py
```

Script akan:

1. Membuka halaman grade report LMS berdasarkan `COURSE_ID`.
2. Membaca halaman pertama gradebook.
3. Mendeteksi pagination dari parameter `page=...`.
4. Mengambil data peserta dari semua halaman.
5. Menghindari duplikasi peserta berdasarkan email.
6. Menormalkan kolom nilai ke format output.
7. Menentukan status kelulusan berdasarkan `PASSING_GRADE`.
8. Memvalidasi data LMS terhadap file master peserta.
9. Menulis hasil ke Excel.

Output default:

```text
nilai_tugas_lms_kelas_b.xlsx
```

## Isi File Output

File Excel output berisi beberapa sheet:

- `Nilai Tugas`: nilai peserta hasil scraping LMS sesuai kolom target.
- `Validasi Peserta`: hasil pencocokan peserta master dengan peserta dari LMS.
- `Extra LMS`: peserta yang ada di LMS, tetapi tidak ada di file master.

Status kelulusan ditentukan dari kolom `nilai akhir`:

```text
Lulus       jika nilai akhir >= 70
Tidak Lulus jika nilai akhir < 70
```

Nilai batas kelulusan diatur melalui konstanta:

```python
PASSING_GRADE = 70
```

## Cara Scraping Data Tiap Kelas

Untuk scraping kelas lain, ubah konfigurasi di bagian atas file `scrape_lms_playwright.py`.

Bagian utama yang perlu disesuaikan:

```python
COURSE_ID = 48
BASE_URL = f"https://lms.kitakerja.id/grade/report/grader/index.php?id={COURSE_ID}"
MASTER_FILE = "master_peserta_b.csv"
OUTPUT_FILE = "nilai_tugas_lms_kelas_b.xlsx"
PASSING_GRADE = 70
```

Langkah scraping per kelas:

1. Cari `COURSE_ID` kelas di LMS.

   Buka grade report kelas di browser, lalu lihat URL. Contoh:

   ```text
   https://lms.kitakerja.id/grade/report/grader/index.php?id=48
   ```

   Angka `48` adalah `COURSE_ID`.

2. Siapkan file master peserta kelas tersebut.

   Contoh untuk kelas A:

   ```text
   master_peserta_a.csv
   ```

   Format kolom tetap:

   ```csv
   Nama,Kelas,Email
   ```

3. Ubah konfigurasi script.

   Contoh untuk kelas A:

   ```python
   COURSE_ID = 47
   MASTER_FILE = "master_peserta_a.csv"
   OUTPUT_FILE = "nilai_tugas_lms_kelas_a.xlsx"
   ```

4. Jalankan scraper:

   ```bash
   python scrape_lms_playwright.py
   ```

5. Cek output Excel sesuai nama file yang diatur di `OUTPUT_FILE`.

Ulangi langkah yang sama untuk kelas lain dengan mengganti `COURSE_ID`, file master, dan file output.

## Kolom Nilai yang Diambil

Kolom output ditentukan oleh `TARGET_COLUMNS` di `scrape_lms_playwright.py`.

Kolom default:

- `No`
- `First name / Last name`
- `Email address`
- `Pre-test`
- `Assignment: Tugas Praktik_Prinsip Manajemen Akses (Real)`
- `Assignment: Tugas Praktik_Prinsip Enkripsi Hashing Windows Ubuntu Kali (Real)`
- `Assignment: Tugas-Check Domain 1 (Real)`
- `Assignment: Tugas Checkpoint Domain 2 (Real)`
- `Assignment: Tugas Checkpoint Domain 3 (Real)`
- `Assignment: Tugas Chapter 4 (Real)`
- `Assignment: Checkpoint Domain 5 (Real)`
- `Post-test`
- `Capstone Project/PBL`
- `nilai akhir`
- `Status Kelulusan`

Jika nama tugas di LMS berubah, sesuaikan mapping di fungsi `header_match()`. Fungsi ini mencocokkan nama kolom dari LMS dengan kolom target berdasarkan keyword.

Contoh mapping:

```python
"Pre-test": [
    "quizpre test",
    "quiz pre test",
    "pre test",
],
```

Tambahkan keyword baru jika nama kolom LMS berbeda tetapi maksud tugasnya sama.

## Alur Validasi Peserta

Validasi dilakukan oleh fungsi `validate_with_master()`.

Urutan pencocokan:

1. Cocokkan berdasarkan email.
2. Jika email tidak cocok, cocokkan berdasarkan nama exact.
3. Jika masih belum cocok, cocokkan berdasarkan kemiripan nama.

Hasil validasi akan diberi catatan:

- `OK`: peserta cocok dengan master.
- `Email master kosong, peserta ditemukan berdasarkan nama.`
- `Nama ditemukan, tetapi email berbeda.`
- `Nama/email dari master tidak ditemukan di hasil scrape LMS.`
- `Ada di LMS, tetapi tidak ada di master peserta.`

## Troubleshooting

### Session belum login atau expired

Jalankan ulang:

```bash
python login_lms.py
```

Lalu jalankan lagi:

```bash
python scrape_lms_playwright.py
```

### Tidak ada data yang berhasil diambil

Cek beberapa hal berikut:

- `COURSE_ID` sudah benar.
- Akun LMS punya akses ke grade report kelas tersebut.
- Session `storage_state.json` masih valid.
- Struktur halaman gradebook LMS belum berubah.

### Kolom nilai kosong atau tidak masuk ke output

Kemungkinan nama kolom di LMS berbeda dari keyword yang dikenali. Perbaiki mapping di fungsi `header_match()` pada `scrape_lms_playwright.py`.

### Peserta tidak cocok dengan master

Cek file master:

- Pastikan kolom bernama `Nama`, `Kelas`, dan `Email`.
- Pastikan email tidak typo.
- Pastikan nama peserta di master cukup mirip dengan nama di LMS.

## Catatan Keamanan

Jangan commit file berikut ke Git:

- `.env`
- `.venv/`
- `storage_state.json`
- file output yang berisi data peserta jika tidak diperlukan

`storage_state.json` berisi session login browser, sehingga harus diperlakukan seperti credential sementara.
