# Student Grouping System

Program untuk membagi mahasiswa baru ke **batch** dan **kelompok** secara otomatis dan merata (berdasarkan prodi, fakultas, dan gender). Hasilnya berupa file Excel. Program juga mengingat hasil sebelumnya, jadi maba susulan bisa ditambahkan tanpa mengacak ulang kelompok yang sudah ada.

Panduan ini memakai **Visual Studio Code (VS Code)** sebagai workspace.

## Daftar Isi

1. [Persiapan (sekali saja)](#1-persiapan-sekali-saja)
2. [Siapkan data](#2-siapkan-data)
3. [Menjalankan program](#3-menjalankan-program)
4. [Hasilnya](#4-hasilnya)
5. [Menambah maba susulan](#5-menambah-maba-susulan)
6. [Menyesuaikan program](#6-menyesuaikan-program)
   - [6.1 Jumlah kelompok dan ukuran maksimal](#61-jumlah-kelompok-dan-ukuran-maksimal)
   - [6.2 Menambah fakultas](#62-menambah-fakultas)
   - [6.3 Nama kolom Excel yang dibaca](#63-nama-kolom-excel-yang-dibaca)
   - [6.4 Isi gender selain LAKI-LAKI dan PEREMPUAN](#64-isi-gender-selain-laki-laki-dan-perempuan)
   - [6.5 Penalty (seberapa penting tiap faktor)](#65-penalty-seberapa-penting-tiap-faktor)
7. [Kalau ada error](#7-kalau-ada-error)

---

## 1. Persiapan (sekali saja)

1. **Install Python** dari [https://www.python.org/downloads/](https://www.python.org/downloads/). Di Windows, **centang "Add python.exe to PATH"** di layar pertama installer.
2. **Install VS Code** dari [https://code.visualstudio.com/](https://code.visualstudio.com/) lalu jalankan installer-nya.
3. **Download program:** di halaman GitHub repository, klik **Code, lalu Download ZIP**, kemudian klik kanan file ZIP dan pilih **Extract All**. Atau pakai `git clone https://github.com/<username>/<nama-repo>.git`.
4. **Buka folder di VS Code:** jalankan VS Code, pilih **File, lalu Open Folder...**, dan pilih folder hasil extract tadi (yang berisi `student_grouping.py`). Kalau muncul pertanyaan *trust the authors*, klik **Yes**.
5. **Pasang ekstensi Python:** kalau VS Code menawarkan install ekstensi Python, klik **Install**. Atau tekan `Ctrl + Shift + X`, cari **Python** (dari Microsoft), lalu Install.
6. **Buka terminal:** pilih **Terminal, lalu New Terminal** (atau tekan ``Ctrl + ` ``). Terminal terbuka di bagian bawah, **langsung di folder program**. Semua perintah di panduan ini diketik di sini.
7. **Install library:** di terminal tersebut, ketik lalu Enter:

   ```
   py -m pip install pandas openpyxl
   ```

Cek Python sudah terpasang dengan mengetik `py --version` (harus muncul nomor versi). Pengguna Mac/Linux: ganti `py` dengan `python3` di seluruh panduan ini.

---

## 2. Siapkan data

Masukkan file data (Excel `.xlsx` atau CSV) ke folder program. Caranya, **masukkan file data** ke folder yang sama dengan file sourcecode program yang sudah didownload.

> **Penting: kolom di data harus sesuai dengan yang dicari program.** Kalau kolomnya tidak ada atau kosong, program tetap jalan, tapi hasilnya jadi aneh (pemerataan tidak bekerja) tanpa ada error.

| Kolom                             | Wajib?       | Kalau tidak ada / kosong                                                                                           |
| --------------------------------- | ------------ | ------------------------------------------------------------------------------------------------------------------ |
| `NIM`                           | Ya           | Baris itu dibuang. NIM dobel juga dibuang.                                                                         |
| `Nama`                          | Ya           | Baris itu dibuang.                                                                                                 |
| `Gender` atau `Jenis Kelamin` | Ya           | Di Excel, judul tidak terdeteksi dan**0 mahasiswa terbaca**. Sel yang kosong dianggap satu kategori sendiri. |
| `Program Studi`                 | Sangat perlu | Semua mahasiswa dianggap satu prodi, jadi pemerataan prodi tidak jalan.                                            |
| `Fakultas`                      | Sangat perlu | Semua jadi`UNKNOWN`, pemerataan fakultas tidak jalan, dan kolom fakultas di dashboard bernilai 0.                |
| `Email`, `No Telp`            | Opsional     | Kolomnya kosong di hasil.                                                                                          |

Catatan:

- Huruf besar/kecil pada judul kolom tidak berpengaruh. Judul boleh tidak di baris pertama (dicari di 30 baris teratas, harus memuat `NIM`, `Nama`, dan `Gender`/`Jenis Kelamin`). Semua sheet dibaca.
- Untuk **CSV**, judul prodi adalah `prodi` (bukan `program studi`).
- Isi kolom `Fakultas` dan `Gender` juga harus dikenali program (lihat [6.2](#62-menambah-fakultas) dan [6.4](#64-isi-gender-selain-laki-laki-dan-perempuan)).
- Kalau judul kolom di datamu berbeda ejaan, ubah di kode ([6.3](#63-nama-kolom-excel-yang-dibaca)) atau ganti judul di Excel.

---

## 3. Menjalankan program

Di terminal VS Code, ketik lalu Enter:

```
py student_grouping.py --input data_maba.xlsx --output hasil
```

| Opsi                  | Fungsi                                                                                                   |
| --------------------- | -------------------------------------------------------------------------------------------------------- |
| `--input` / `-i`  | **Wajib.** File data mahasiswa. Kalau nama file ada spasi, apit dengan tanda kutip.                |
| `--output` / `-o` | Awalan nama file hasil (dengan atau tanpa`.xlsx`). Default: `hasil_pengelompokkan`.                  |
| `--master` / `-m` | Master Data dari run sebelumnya (lihat[bagian 5](#5-menambah-maba-susulan)). Jangan diisi di run pertama. |

> **Jangan klik tombol ▶ (Run Python File) di pojok kanan atas VS Code.** Program butuh opsi `--input`, jadi kalau dijalankan dengan tombol itu akan muncul error `the following arguments are required: --input`. Selalu jalankan lewat terminal seperti di atas.

Setelah selesai, cek baris **`Ukuran Kelompok Terbesar`**. Kalau tidak melebihi 55 dan tidak ada tulisan `WARNING`, hasilnya aman.

> Run pertama menghasilkan pembagian yang **berbeda tiap kali** karena ada unsur acak. Kalau sudah puas, simpan Master Data-nya dan jangan jalankan ulang dari awal.

---

## 4. Hasilnya

File hasil muncul di panel kiri VS Code. Untuk membukanya, klik kanan file lalu pilih **Reveal in File Explorer**, kemudian buka dengan Excel.

| File                                   | Isi                                                                                                                                                                                                        |
| -------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `hasil_Batch_1.xlsx`, `_2`, `_3` | Satu file per batch: sheet**Dashboard** (rekap per fakultas dan per kelompok) dan satu sheet per kelompok. Kolom **Status** bertuliskan **NEW** (merah) untuk mahasiswa yang baru masuk. |
| `hasil_Master_Data.xlsx`             | Data semua mahasiswa beserta batch dan kelompoknya.**Simpan**, dipakai untuk run berikutnya.                                                                                                         |

Tutup file Excel hasil sebelum menjalankan program lagi, kalau tidak akan muncul `PermissionError`.

---

## 5. Menambah maba susulan

Kelompok yang sudah dibagikan tidak boleh berubah, hanya ditambah:

```
py student_grouping.py --input susulan.xlsx --master hasil_Master_Data.xlsx --output hasil2
```

- File input boleh berisi maba baru saja, atau daftar lengkap. NIM yang sudah ada di master diabaikan.
- Anak lama tetap di kelompoknya. Anak baru masuk ke kelompok yang paling cocok dan diberi tanda **NEW**.
- Untuk susulan berikutnya, pakai Master Data **terbaru** (`hasil2_Master_Data.xlsx`).
- Perbaikan data anak lama (misalnya typo nama) dilakukan **langsung di file master**, karena data di file input tidak menimpa master.
- Memindahkan satu orang: ubah kolom **Batch** dan **Kelompok**-nya di master. Jangan ubah judul kolom.

---

## 6. Menyesuaikan program

Semua pengaturan ada di `student_grouping.py`. Cara mengeditnya di VS Code:

- Klik `student_grouping.py` di panel kiri. Tekan `Ctrl + F` untuk mencari teks, dan `Ctrl + S` untuk menyimpan setelah mengubah.
- **Buat salinan cadangan dulu:** klik kanan file di panel kiri, pilih **Copy**, lalu klik kanan area kosong dan pilih **Paste**.
- Jangan mengubah spasi di awal baris (di Python itu penting). Kalau salah, tekan `Ctrl + Z` untuk membatalkan.

### 6.1 Jumlah kelompok dan ukuran maksimal

```python
GROUPS_PER_BATCH_MAP = {1: 53, 2: 53, 3: 53}
MAX_STUDENTS_PER_GROUP = 55
```

Format `nomor batch: jumlah kelompok maksimal`. Contoh 2 batch: `{1: 40, 2: 40}`. Pastikan **jumlah kelompok × ukuran maksimal lebih besar dari jumlah mahasiswa per batch**. Kalau kurang, muncul `WARNING ... terpaksa melebihi`.

### 6.2 Menambah fakultas

Ada dua tempat yang harus diubah bersama:

```python
FACULTIES_LIST = ["FTE", "FRI", "FIF", "FEB", "FKS", "FIK", "FIT", "FK"]
```

dan di `text_map` (dalam `def resolved_faculty`), format `"TULISAN DI DATA": "SINGKATAN"`:

```python
"FAKULTAS HUKUM": "FH", "FH": "FH",
```

Tulis dengan **HURUF BESAR**. Singkatan yang merupakan potongan singkatan lain (contoh `FK` di dalam `FKS`) harus ditaruh **lebih bawah**. Fakultas yang tidak ada di kamus tidak terhitung di dashboard.

### 6.3 Nama kolom Excel yang dibaca

Di `def _row_to_student`, kata dalam `[ ... ]` adalah judul kolom yang dicari (huruf kecil). Tambahkan ejaan kolommu:

```python
major=parse_str(get_val(["program studi", "prodi", "jurusan"])),
```

Kalau ejaan `Nama`, `NIM`, atau gender yang berbeda, ubah juga `def _find_header_row`. Untuk CSV, ubah di `def _load_from_csv` (bagian `Student(...)`).

### 6.4 Isi gender selain LAKI-LAKI dan PEREMPUAN

Dashboard menghitung berdasarkan dua tulisan itu. Kalau datamu `L`/`P`, di `def _build_dashboard_tab` ganti:

```python
g.gender_counts.get("L", 0)
g.gender_counts.get("P", 0)
```

(Pembagian kelompoknya tetap benar, hanya angka di dashboard yang 0 kalau tidak diganti.)

### 6.5 Penalty (seberapa penting tiap faktor)

Setiap kali menaruh sekelompok kecil mahasiswa (2-3 orang, prodi dan gender sama), program menghitung **skor penalty** tiap kelompok dan memilih yang **skornya paling kecil**. Makin besar angkanya, makin diprioritaskan faktor itu. Cari `def calculate_penalty`:

| Baris di kode                     | Artinya                                                                                 |
| --------------------------------- | --------------------------------------------------------------------------------------- |
| `self.size * 57.0`              | Makin banyak anggota, makin dihindari. Menjaga**ukuran kelompok rata**.           |
| `self.major_counts... * 7.0`    | Makin banyak yang prodinya sama, makin dihindari (**sebar prodi**).               |
| `major_penalty += 300.0`        | Menghindari mahasiswa**sendirian** di kelompok yang belum punya teman satu prodi. |
| `self.gender_counts... * 11.0`  | Sebar gender.                                                                           |
| `self.faculty_counts... * 21.0` | Sebar fakultas.                                                                         |
| `175.0 if ... first_names`      | Menghindari**nama depan kembar** dalam satu kelompok.                             |
| `size_penalty += 999999.0`      | Mencegah kelompok melebihi batas maksimal.**Jangan diubah.**                      |

Contoh: ingin pemerataan fakultas lebih diutamakan, ubah `21.0` jadi `40.0`. Tidak peduli nama kembar, ubah `175.0` jadi `0.0`.

Tips:

- Ubah **satu angka per percobaan**, lalu jalankan dan bandingkan. Karena ada unsur acak, coba beberapa kali sebelum menyimpulkan.
- Cek hasilnya lewat dashboard dan baris `Ukuran Kelompok Terkecil/Terbesar`.
- Jangan membuat angka ukuran kelompok (`57.0`) jauh lebih kecil dari yang lain, karena ukuran kelompok bisa jadi timpang.

Daftar nama depan yang dilewati saat cek nama kembar ada di `EXCLUDED_FIRST_NAMES` (huruf kecil, dipisah koma).

---

## 7. Kalau ada error

| Yang muncul                                              | Solusi                                                                                                                        |
| -------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------- |
| `'py' is not recognized`                               | Python belum terpasang atau*Add to PATH* tidak dicentang. Install ulang Python, lalu **tutup dan buka lagi VS Code**. |
| `No module named 'pandas'` / `'openpyxl'`            | Jalankan`py -m pip install pandas openpyxl` di terminal VS Code.                                                            |
| `the following arguments are required: --input`        | Program dijalankan lewat tombol ▶. Jalankan lewat terminal (bagian 3).                                                       |
| `File input tidak ditemukan`                           | Cek ejaan nama file dan pastikan ada di folder yang dibuka di VS Code.                                                        |
| `Memuat 0 mahasiswa`                                   | Judul kolom tidak terbaca ([bagian 2](#2-siapkan-data) dan [6.3](#63-nama-kolom-excel-yang-dibaca)).                            |
| `PermissionError`                                      | File hasil masih terbuka di Excel. Tutup dulu.                                                                                |
| `WARNING ... terpaksa melebihi 55`                     | Kapasitas kurang, naikkan jumlah kelompok atau ukuran maksimal ([6.1](#61-jumlah-kelompok-dan-ukuran-maksimal)).               |
| Fakultas / Laki-Laki / Perempuan di dashboard bernilai 0 | Lihat[6.2](#62-menambah-fakultas) dan [6.4](#64-isi-gender-selain-laki-laki-dan-perempuan).                                     |
| `IndentationError` setelah mengedit                    | Spasi di awal baris berubah. Tekan`Ctrl + Z`, atau kembalikan dari file cadangan.                                           |

Kalau masih bingung, salin pesan error lengkap dari terminal dan tanya AI :P.
