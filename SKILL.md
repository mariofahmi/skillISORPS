---
name: rps-unirow
description: >
  Skill untuk menyusun, mengisi, dan memformat Rencana Pembelajaran Semester
  (RPS) sesuai format resmi dan standar mutu Universitas PGRI Ronggolawe
  (UNIROW) Tuban berbasis kurikulum OBE 2026. WAJIB gunakan skill ini setiap
  kali pengguna menyebut "RPS", "Rencana Pembelajaran Semester", "susun RPS",
  "buatkan RPS", "isi RPS", "format RPS UNIROW/UNIROW Tuban/Ronggolawe", ingin
  membuat dokumen .docx RPS untuk mata kuliah tertentu, ingin melengkapi CPL/
  CPMK/Sub-CPMK, tabel matriks korelasi CPL terhadap Sub-CPMK, rencana
  pembelajaran mingguan 16 minggu (termasuk UTS di Mg 8 dan UAS di Mg 16), rubrik
  penilaian holistik, konversi nilai akhir, contoh soal UAS, lembar validasi,
  atau ingin mengecek/merevisi RPS yang sudah ada agar presisi sesuai standar
  resmi kampus, bahkan jika tidak secara eksplisit menyebut kata "UNIROW".
---

# RPS UNIROW — Panduan Resmi Penyusunan RPS OBE 2026 (Presisi 100%)

Dokumen acuan standar mutu: `01_1_RPS_Antropologi Budaya_OBE.docx` (Kurikulum OBE 2026, FKIP Universitas PGRI Ronggolawe Tuban).
Template dasar: `assets/Template_RPS_UNIROW.docx`.
Peta placeholder & anatomi tabel: `references/placeholder_map.md`.
Script otomasi generator: `scripts/build_rps.py`.
Live Simulator & Generator Web: [https://mariofahmi.github.io/skillISORPS/#pratinjau](https://mariofahmi.github.io/skillISORPS/#pratinjau)
Katalog Terintegrasi: 63 Mata Kuliah Kurikulum OBE 2026 (Semester 1–7).

Skill ini menghasilkan dokumen `.docx` RPS yang **identik secara format, tata letak, dan tipografi** dengan standar mutu resmi kurikulum OBE 2026, serta selaras 100% dengan portal simulator web presisi.

---

## 1. Spesifikasi Format & Desain Dokumen Acuan

| Parameter | Ketentuan Baku |
|---|---|
| **Orientasi Halaman** | **100% A4 Landscape** (1 *Section* penuh, tanpa rotasi ke portrait) |
| **Ukuran Kertas** | A4 Landscape: `w="16838" h="11909"` DXA (`841.7 × 595.45` pt) |
| **Margin Halaman** | Atas: `39.6 pt`, Bawah: `39.6 pt`, Kiri: `43.2 pt`, Kanan: `43.2 pt` |
| **Tipografi Utama** | **Cambria** (Isi tabel/body: 10.5 pt, Sub-header: 11 pt, Judul: 12 pt bold) |
| **Garis Tabel (Borders)**| Single line, warna `#B0B5B3`, ketebalan `sz="4"` (0.5 pt) di seluruh tabel |
| **Bantalan Sel (Padding)**| Top/Bottom: `100 dxa`, Left/Right: `140 dxa` |
| **Palet Warna Shading** | *Slate Gray* elegan: Header `#EAECEE`, Sub-header `#F2F4F4` / `#F4F6F7`, Kartu Validasi `#FAFAFA` |

---

## 2. Struktur Anatomi Dokumen RPS (4 Tabel + Paragraf Evaluasi)

Setiap dokumen RPS UNIROW wajib memuat komponen berurutan berikut:

### I. TABEL 0: Identitas, Otorisasi, Capaian Pembelajaran, & Detail MK (20 Baris)
1. **Baris 0**: Logo resmi UNIROW (`image1.png`) \| KOP Resmi Universitas, Fakultas, dan Prodi \| Kode Dokumen.
2. **Baris 1**: Judul `RENCANA PEMBELAJARAN SEMESTER` (Merge 8 kolom, Shading `#EAECEE`).
3. **Baris 2–3**: Identitas MK (Nama MK, Kode MK, Rumpun MK, Bobot SKS, Semester, Tgl Penyusunan).
4. **Baris 4–5**: Otorisasi (Dosen Pengembang RPS, Koordinator RMK, Ketua PRODI lengkap dengan gelar).
5. **Baris 6–7**: CPL-PRODI yang dibebankan pada MK (Kode dan deskripsi lengkap CPL).
6. **Baris 8–9**: Capaian Pembelajaran Mata Kuliah (CPMK 1 s.d. CPMK 4).
7. **Baris 10–11**: Kemampuan akhir tiap tahapan belajar (Sub-CPMK 1 s.d. 14) **wajib menyertakan label taksonomi Bloom**, contoh: `[C2, A2]`, `[C4, A3]`, `[C5, A3]`, `[C6, A4]`.
8. **Baris 12–13**: **Matriks Korelasi CPL terhadap Sub-CPMK (Tabel Bersarang / Nested Table 18 Baris)**:
   - Header: `Sub-CPMK / Evaluasi`, kolom masing-masing CPL (`CPL... (%)`), dan `Bobot Penilaian (%)`.
   - Baris 1 s.d. 7: Sub-CPMK 1 s.d. 7 (tanda centang `✓` dan bobot 3%–4%).
   - Baris 8: Evaluasi Tengah Semester (`UTS`, Bobot `25%`, Shading `#F4F6F7`).
   - Baris 9 s.d. 15: Sub-CPMK 8 s.d. 14 (tanda centang `✓` dan bobot 3%–4%).
   - Baris 16: Evaluasi Akhir Semester (`UAS`, Bobot `25%`, Shading `#F4F6F7`).
   - Baris 17: `Total` (Nilai `100` pada tiap CPL dan total bobot `100%`, Shading `#EAECEE`).
9. **Baris 14**: Deskripsi Singkat MK (1 paragraf komprehensif profil kompetensi).
10. **Baris 15**: Bahan Kajian: Materi Pembelajaran (dengan kode Bahan Kajian, misal `BK07...`).
11. **Baris 16–17**: Pustaka Utama (buku fundamental ber-ISBN) dan Pustaka Pendukung (jurnal, perundang-undangan).
12. **Baris 18–19**: Dosen Pengampu dan Matakuliah Syarat (atau `Tidak ada`).

### II. TABEL 1: Rencana Pembelajaran 16 Minggu (19 Baris)
Tabel matriks perkuliahan 16 minggu dengan format 8 kolom:
- **Baris 0–2**: Header 3 baris standar (Nomor kolom `(1)` s.d. `(8)`).
- **Baris 3–9 (Minggu 1 s.d. 7)**: Sesi perkuliahan reguler (Sub-CPMK 1 s.d. 7).
- **Baris 10 (Minggu 8 — UTS)**: Baris evaluasi khusus (Sel 1–6 digabung bertuliskan `UJIAN TENGAH SEMESTER (UTS)\nEvaluasi penguasaan materi perkuliahan minggu 1 s.d. 7`, Bobot `25%`, Shading `#F4F6F7`).
- **Baris 11–17 (Minggu 9 s.d. 15)**: Sesi perkuliahan reguler lanjutan (Sub-CPMK 8 s.d. 14).
- **Baris 18 (Minggu 16 — UAS)**: Baris evaluasi khusus (Sel 1–6 digabung bertuliskan `UJIAN AKHIR SEMESTER (UAS)\nEvaluasi komprehensif penguasaan capaian pembelajaran mata kuliah`, Bobot `25%`, Shading `#F4F6F7`).

> **Aturan Waktu Belajar SN-Dikti**:
> - **PB** (Tatap Muka): $sks \times 50$ menit/minggu (2 SKS = 100 menit).
> - **PT** (Terstruktur): $sks \times 60$ menit/minggu (2 SKS = 120 menit).
> - **KM** (Mandiri): $sks \times 60$ menit/minggu (2 SKS = 120 menit).

### III. TABEL 2: Rubrik Penilaian Holistik (5 Baris × 5 Kolom)
Tabel rubrik penilaian hasil belajar mahasiswa:
- Kolom: `Aspek Penilaian`, `Sangat Baik (85-100)`, `Baik (70-84)`, `Cukup (60-69)`, `Bobot`.
- 4 Aspek Penilaian Utama:
  1. **Pemahaman Konsep (40%)**: Penguasaan teori dan kejelasan definisi.
  2. **Kemampuan Analisis (30%)**: Ketajaman berpikir kritis dan kedalaman penalaran.
  3. **Aplikasi Konsep dalam Konteks Lokal (20%)**: Presisi penerapan teori pada fenomena nyata.
  4. **Sistematika & Komunikasi (10%)**: Kerapian tata tulis akademik, orisinalitas, dan komunikasi.

### IV. BAGIAN EVALUASI, SOAL UAS, & VALIDASI (Paragraf Terstruktur)
1. **Kriteria Kelulusan**: Syarat minimal nilai 60 (C) dengan komposisi:
   - Keaktifan & Partisipasi: `10%`
   - Tugas & Kuis: `15%`
   - UTS: `25%`
   - Proyek Akhir: `15%`
   - UAS: `35%`
2. **Konversi Nilai Akhir (Standar Mutu UNIROW Tuban - Skala 7)**:
   - `A  : 85 - 100 (Sangat Baik)`
   - `AB : 80 - 84`
   - `B  : 75 - 79 (Baik)`
   - `BC : 70 - 74`
   - `C  : 65 - 69 (Cukup - Batas Minimal Kelulusan)`
   - `D  : 60 - 64`
   - `E  : < 60 (Tidak Lulus)`
3. **Contoh Soal Ujian Akhir Semester (UAS)**:
   - **5 butir soal esai analitis tingkat tinggi (HOTS)** mengukur CPMK 1 s.d. CPMK 4 (level kognitif C4–C6).
   - Ditutup dengan kalimat: `--- SELAMAT MENGERJAKAN ---`.
4. **Pengantar Lembar Validasi**:
   - Judul: `LEMBAR VALIDASI RENCANA PEMBELAJARAN SEMESTER (RPS)`.
   - Pernyataan pemenuhan standar mutu kurikulum OBE UNIROW Tuban.
   - Tanggal: `Divalidasi di Tuban, pada tanggal: [dd Bulan yyyy]`.

### V. TABEL 3: Lembar Pengesahan / Validasi (2 Baris × 2 Kolom)
Tabel kartu validasi simetris (Shading `#FAFAFA`):
- Kolom Kiri: `Menyetujui,\nKetua Program Studi [Nama Prodi]` \| Tanda tangan & `[Nama Kaprodi, Gelar]\nNIDN. [NIDN]`.
- Kolom Kanan: `Mengetahui,\nUnit Jaminan Mutu (UJM) Prodi [Nama Prodi]` \| Tanda tangan & `[Tim Penjamin Mutu, Gelar]\nNIDN. [NIDN]`.

---

## 3. Alur Kerja Pembuatan Dokumen RPS

### Langkah 1: Kumpulkan / Identifikasi Data Mata Kuliah
Kumpulkan informasi mata kuliah:
- Nama MK, Kode MK, Rumpun MK, Bobot SKS, Semester, Tanggal Penyusunan.
- Dosen Pengembang, Koordinator RMK, Ka PRODI.
- CPL Prodi, CPMK, Sub-CPMK (lengkap dengan kode taksonomi `[C.., A..]`).
- Deskripsi singkat MK, Bahan Kajian (kode BK), Pustaka Utama & Pendukung.
- Pembagian topik 16 minggu (UTS di Mg 8, UAS di Mg 16).

### Langkah 2: Eksekusi Generator Otomatis (Metode Presisi)
Gunakan script `build_rps.py` yang sudah terintegrasi di dalam skill:

```powershell
# 1. Buat file data JSON atau gunakan data MK yang dikumpulkan
py -3 ".agents\skills\rps-unirow\scripts\build_rps.py" --create-sample "data_mk.json"

# 2. Edit data_mk.json sesuai mata kuliah pengguna, lalu jalankan generator:
py -3 ".agents\skills\rps-unirow\scripts\build_rps.py" --data "data_mk.json" --output "01_1_RPS_Nama_MK_OBE.docx"
```

Script ini secara otomatis:
- Memuat master template resmi ber-logo asli UNIROW (`assets/Template_RPS_UNIROW.docx`).
- Mempertahankan seluruh format border `#B0B5B3`, margin landscape A4, dan font Cambria.
- Memperbarui Tabel 0, Tabel Matriks Korelasi Nested, Tabel 16 Minggu, Tabel Rubrik Penilaian, Konversi Nilai, 5 Soal UAS HOTS, dan Tabel Validasi.

---

## 4. Daftar Cek Kualitas Dokumen (Anti-Defect QA Checklist)

Sebelum menyerahkan file ke pengguna, verifikasi poin-poin berikut:
- [ ] Dokumen **100% berorientasi Landscape** A4 (tidak ada pergantian section ke portrait).
- [ ] Font konsisten **Cambria** (10.5 pt untuk teks isi dan tabel).
- [ ] Garis tabel menggunakan warna `#B0B5B3` dengan ketebalan 0.5 pt.
- [ ] Matriks korelasi CPL terhadap Sub-CPMK berada **di dalam baris ke-13 Tabel 0** sebagai *nested table*.
- [ ] UTS berada di **Minggu ke-8** dan UAS di **Minggu ke-16** pada Tabel 1 (tidak dihitung sebagai Sub-CPMK).
- [ ] Penomoran Sub-CPMK tidak bergeser (*no off-by-one error*): Sub-CPMK 1–7 di Minggu 1–7; Sub-CPMK 8–14 di Minggu 9–15.
- [ ] Total bobot mingguan tepat **100%** (UTS 25% + UAS 25% + perkuliahan mingguan 50%).
- [ ] Terdapat **Tabel Rubrik Penilaian Holistik** (4–5 aspek penilaian).
- [ ] Terdapat **Konversi Nilai Akhir Skala 7** (A, AB, B, BC, C, D, E).
- [ ] Terdapat **5 Butir Contoh Soal UAS HOTS** yang selaras dengan CPMK mata kuliah.
- [ ] Tabel 3 Lembar Validasi memuat tanda tangan pengesahan **Ketua Program Studi** dan **Unit Jaminan Mutu (UJM) Prodi**.
