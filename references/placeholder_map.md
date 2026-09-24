# Peta Struktur Dokumen & Placeholder — Format Resmi RPS UNIROW (OBE 2026)

Dokumen acuan standar: `01_1_RPS_Antropologi Budaya_OBE.docx` (Kurikulum OBE 2026, FKIP Universitas PGRI Ronggolawe Tuban).
Template dasar: `assets/Template_RPS_UNIROW.docx`.
Script pembangun otomatis: `scripts/build_rps.py`.

Dokumen terdiri dari **4 tabel utama** dan blok paragraf terstruktur dalam satu dokumen Word (.docx) berorientasi **100% A4 Landscape** (1 section, tanpa pemisahan section ke portrait).

---

## Spesifikasi Teknis Halaman & Desain

| Parameter | Nilai Baku |
|---|---|
| **Ukuran Kertas** | A4 Landscape (`w="16838" h="11909"` DXA / `841.7 × 595.45` pt) |
| **Margin Halaman** | Top: `39.6 pt` (792 dxa), Bottom: `39.6 pt`, Left: `43.2 pt` (864 dxa), Right: `43.2 pt` |
| **Font Utama** | **Cambria** (Body/Isi: 10.5 pt, Sub-header/Tabel: 11 pt, Judul: 12 pt bold) |
| **Garis Tabel (Borders)** | Single line, warna `#B0B5B3`, ketebalan `sz="4"` (0.5 pt) di seluruh sisi & internal |
| **Bantalan Sel (Padding)** | Top: `100 dxa`, Bottom: `100 dxa`, Left: `140 dxa`, Right: `140 dxa` |
| **Palet Warna Shading** | Header Utama: `#EAECEE`, Sub-header: `#F2F4F4` / `#F4F6F7`, Kartu Validasi: `#FAFAFA` |

---

## 1. TABEL 0 — Identitas MK, Otorisasi, Capaian Pembelajaran, & Detail MK
Struktur: **20 baris**, 8 kolom grid (`tblGrid` widths: `1722, 1971, 2255, 1517, 1914, 1730, 1748, 2271` dxa).

### Anatomi Baris demi Baris:

| Baris | Nama / Fungsi | Struktur Sel & Isi |
|:---:|---|---|
| **0** | **Kop & Kode Dokumen** | • **Cell 0**: Logo resmi UNIROW (`media/image1.png`).<br>• **Cell 1**: Kop teks (Center, Bold, 11 pt): `UNIVERSITAS PGRI RONGGOLAWE TUBAN`, `FAKULTAS [NAMA FAKULTAS]`, `PROGRAM STUDI [NAMA PRODI]`.<br>• **Cell 7**: Kode Dokumen (Center, Bold, 10.5 pt), mis. `PPKn/I/PKN/MKK PPKn`. |
| **1** | **Judul Dokumen** | • **Cell 0** (Merge 8 kolom): `RENCANA PEMBELAJARAN SEMESTER` (Center, Bold, 12 pt, Shading `#EAECEE`). |
| **2** | **Header Identitas MK** | • `MATA KULIAH (MK)` (span 2) \| `KODE` \| `Rumpun MK` (span 2) \| `BOBOT (sks)` \| `SEMESTER` \| `Tgl Penyusunan` (span 2). Shading `#F4F6F7`, Bold 10.5 pt. |
| **3** | **Data Identitas MK** | • `[Nama MK]` (Bold) \| `[Kode MK]` \| `[Rumpun MK]` \| `[X] SKS` \| `[Semester]` \| `[Tanggal Penyusunan]`. |
| **4** | **Header Otorisasi** | • `OTORISASI` (span 2) \| `Pengembang RPS` (span 2) \| `Koordinator RMK` (span 2) \| `Ketua PRODI` (span 2). Shading `#F4F6F7`, Bold 10.5 pt. |
| **5** | **Data Otorisasi** | • `OTORISASI` (vertical span) \| `[Nama Dosen Pengembang, Gelar]` \| `[Nama Koordinator RMK, Gelar]` \| `[Nama Ka PRODI, Gelar]`. |
| **6** | **Header CPL-PRODI** | • `Capaian Pembelajaran (CP)` (span 2) \| `CPL-PRODI yang dibebankan pada MK` (span 6). Shading `#F4F6F7`, Bold 10.5 pt. |
| **7** | **Isi CPL-PRODI** | • Kiri: vertical span. Kanan: Deskripsi lengkap CPL prodi berformat `CPL1: ... \n\nCPL2: ... \n\nCPL8: ...` (Cambria 10.5 pt). |
| **8** | **Header CPMK** | • Kiri: vertical span. Kanan: `Capaian Pembelajaran Mata Kuliah (CPMK)` (Bold 10.5 pt, Shading `#F4F6F7`). |
| **9** | **Isi CPMK** | • Kiri: vertical span. Kanan: Daftar CPMK berformat `CPMK1\t[Deskripsi]\nCPMK2\t[Deskripsi]...`. |
| **10** | **Header Sub-CPMK** | • Kiri: vertical span. Kanan: `Kemampuan akhir tiap tahapan belajar (Sub-CPMK)` (Bold 10.5 pt, Shading `#F4F6F7`). |
| **11** | **Isi Sub-CPMK** | • Kiri: vertical span. Kanan: Daftar 14 Sub-CPMK dengan kode taksonomi Bloom, misal: `Sub-CPMK 1 [C2, A2]: ... \nSub-CPMK 2 [C2, A2]: ...`. |
| **12** | **Header Korelasi** | • Kiri: vertical span. Kanan: `Korelasi antara CPL/CPMK terhadap Sub-CPMK` (Bold 10.5 pt, Shading `#F4F6F7`). |
| **13** | **Tabel Matriks Korelasi (Nested Table)** | • Kiri: vertical span. Kanan: **Tabel Bersarang (Nested Table)** memuat pemetaan matriks korelasi CPL terhadap Sub-CPMK (rincian di bawah). |
| **14** | **Deskripsi Singkat MK** | • Kiri (span 2): `Deskripsi Singkat MK` (Bold). Kanan (span 6): 1 paragraf komprehensif deskripsi ruang lingkup mata kuliah. |
| **15** | **Bahan Kajian** | • Kiri (span 2): `Bahan Kajian: Materi Pembelajaran` (Bold). Kanan (span 6): Pokok bahasan dengan kode Bahan Kajian (misal `BK07 Sosiologi dan Antropologi, mencakup: 1. ... 2. ... s.d. 12. ...`). |
| **16** | **Pustaka Utama** | • Kiri (span 2): `Pustaka` (Bold). Tengah (span 1): `Utama :` (Bold). Kanan (span 5): Referensi buku utama ber-ISBN. |
| **17** | **Pustaka Pendukung** | • Kiri (span 2): vertical span. Tengah (span 1): `Pendukung :` (Bold). Kanan (span 5): Referensi pendukung / jurnal / perundang-undangan. |
| **18** | **Dosen Pengampu** | • Kiri (span 2): `Dosen Pengampu` (Bold). Kanan (span 6): Nama dosen pengampu / tim dosen lengkap dengan gelar. |
| **19** | **Matakuliah Syarat** | • Kiri (span 2): `Matakuliah syarat` (Bold). Kanan (span 6): Nama MK prasyarat jika ada, atau `Tidak ada`. |

---

### Anatomi Tabel Bersarang (Nested Table di Baris 13 Tabel 0):
Tabel ini memetakan korelasi CPL prodi terhadap setiap Sub-CPMK dan evaluasi berkala:
- **Jumlah Baris**: 18 baris tetap.
  - **Baris 0**: Header (`Sub-CPMK / Evaluasi`, `CPL1 (%)`, `CPL2 (%)`, ..., `Bobot Penilaian (%)`). Shading `#EAECEE`, Bold.
  - **Baris 1 s.d. 7**: `Sub-CPMK1` s.d. `Sub-CPMK7` (tanda `✓` pada kolom CPL terkait, bobot 3%–4%).
  - **Baris 8**: `UTS` (Center, Bold, Shading `#F4F6F7`, Bobot: `25%`).
  - **Baris 9 s.d. 15**: `Sub-CPMK8` s.d. `Sub-CPMK14` (tanda `✓` pada kolom CPL terkait, bobot 3%–4%).
  - **Baris 16**: `UAS` (Center, Bold, Shading `#F4F6F7`, Bobot: `25%`).
  - **Baris 17**: `Total` (Bold, Shading `#EAECEE`, nilai `100` pada tiap CPL dan total bobot `100%`).

---

## 2. TABEL 1 — Rencana Pembelajaran 16 Minggu
Struktur: **19 baris**, 8 kolom grid (`tblGrid` widths: `883, 2050, 2016, 1913, 2089, 2314, 2324, 1539` dxa).

### Struktur Header (3 Baris Pertama):
- **Baris 0**: `Mg Ke-` (rowspan 2) \| `Kemampuan akhir tiap tahapan belajar (SubCPMK)` (rowspan 2) \| `Penilaian` (colspan 2) \| `Bentuk Pembelajaran, Metode Pembelajaran, Penugasan Mahasiswa, [Estimasi Waktu]` (colspan 3) \| `Materi Pembelajaran [Pustaka]` (rowspan 2) \| `Bobot Penilaian (%)` (rowspan 2). Shading `#EAECEE`.
- **Baris 1**: `Indikator` \| `Teknik & Kriteria` \| `Luring (Tatap Muka)` \| `Daring (Penugasan Mandiri)`. Shading `#EAECEE`.
- **Baris 2**: Penomoran kolom: `(1)` \| `(2)` \| `(3)` \| `(4)` \| `(5)` \| `(6)` \| `(7)` \| `(8)`. Shading `#F4F6F7`.

### Struktur Baris Pertemuan (Baris 3 s.d. 18):
- **Baris 3 s.d. 9 (Minggu 1 s.d. 7)**: Sesi perkuliahan reguler.
  - Kolom (1): Nomor minggu (`1` s.d. `7`).
  - Kolom (2): `Sub-CPMK X [C..., A...]: Mampu ...`
  - Kolom (3): Indikator kinerja spesifik & terukur.
  - Kolom (4): Teknik tes/non-tes dan kriteria penskoran.
  - Kolom (5): Bentuk/Metode Luring dan jam belajar `[PB: 1x(2x50')] = 100 menit`.
  - Kolom (6): Bentuk Daring / LMS dan jam belajar `[PT+KM: 1x(2x60')+1x(2x60')] = 240 menit`.
  - Kolom (7): Materi kajian dan rujukan buku `[Pustaka: Utama ..., Pendukung ...]`.
  - Kolom (8): Bobot penilaian berkala (`3` atau `4`).
- **Baris 10 (Minggu 8 — UTS)**:
  - Kolom (1): `8` (Center, Bold).
  - Kolom (2 s.d. 6 merged): `UJIAN TENGAH SEMESTER (UTS)\nEvaluasi penguasaan materi perkuliahan minggu 1 s.d. 7` (Center, Bold).
  - Kolom (7 merged): ikut dalam merge.
  - Kolom (8): `25` (Center, Bold).
  - Seluruh baris diberi shading `#F4F6F7`.
- **Baris 11 s.d. 17 (Minggu 9 s.d. 15)**: Sesi perkuliahan reguler lanjutan (Sub-CPMK 8 s.d. 14).
- **Baris 18 (Minggu 16 — UAS)**:
  - Kolom (1): `16` (Center, Bold).
  - Kolom (2 s.d. 6 merged): `UJIAN AKHIR SEMESTER (UAS)\nEvaluasi komprehensif penguasaan capaian pembelajaran mata kuliah` (Center, Bold).
  - Kolom (8): `25` (Center, Bold).
  - Seluruh baris diberi shading `#F4F6F7`.

---

## 3. TABEL 2 — Rubrik Penilaian Holistik
Struktur: **5 baris**, 5 kolom (`tblGrid` widths: `2969, 3228, 3228, 3228, 2475` dxa).

- **Baris 0**: `Aspek Penilaian` \| `Sangat Baik (85-100)` \| `Baik (70-84)` \| `Cukup (60-69)` \| `Bobot`. Shading `#EAECEE`, Bold.
- **Baris 1**: `Pemahaman Konsep ... (40%)` \| Uraian Sangat Baik \| Uraian Baik \| Uraian Cukup \| `40%`.
- **Baris 2**: `Kemampuan Analisis ... (30%)` \| Uraian Sangat Baik \| Uraian Baik \| Uraian Cukup \| `30%`.
- **Baris 3**: `Aplikasi Konsep ... (20%)` \| Uraian Sangat Baik \| Uraian Baik \| Uraian Cukup \| `20%`.
- **Baris 4**: `Sistematika & Komunikasi (10%)` \| Uraian Sangat Baik \| Uraian Baik \| Uraian Cukup \| `10%`.

---

## 4. BAGIAN ASESMEN, SOAL UAS, & VALIDASI (Paragraf Terstruktur)

Setelah Tabel 2, dokumen memuat bagian paragraf berikut:

1. **Kriteria Kelulusan**:
   - Judul: `Kriteria Kelulusan:` (Bold 11 pt).
   - Kalimat pengantar: `Mahasiswa dinyatakan lulus mata kuliah [Nama MK] apabila memperoleh nilai akhir minimal 60 (C) dengan komponen penilaian sebagai berikut:`
   - Poin komponen penilaian:
     - `Keaktifan & Partisipasi: 10%`
     - `Tugas & Kuis: 15%`
     - `UTS: 25%`
     - `Proyek Akhir: 15%`
     - `UAS: 35%`
2. **Konversi Nilai Akhir (Standar Mutu UNIROW Tuban)**:
   - Judul: `Konversi Nilai Akhir:` (Bold 11 pt).
   - `A  : 85 - 100 (Sangat Baik)`
   - `AB : 80 - 84`
   - `B  : 75 - 79 (Baik)`
   - `BC : 70 - 74`
   - `C  : 65 - 69 (Cukup)`
   - `D  : 60 - 64`
   - `E  : < 60 (Tidak Lulus)`
3. **Contoh Soal Ujian Akhir Semester (UAS)**:
   - Judul: `CONTOH SOAL UJIAN AKHIR SEMESTER (UAS)` (Bold 11 pt).
   - Memuat **5 butir soal esai analitis tingkat tinggi (HOTS)** yang mengukur ketercapaian CPMK 1 s.d. CPMK 4 (C4, C5, C6).
   - Penutup: `--- SELAMAT MENGERJAKAN ---` (Center, Italic/No Spacing).
4. **Pengantar Lembar Validasi**:
   - Judul: `LEMBAR VALIDASI RENCANA PEMBELAJARAN SEMESTER (RPS)` (Center, Bold 12 pt).
   - Pernyataan: `Rencana Pembelajaran Semester (RPS) mata kuliah [Nama MK] ini telah divalidasi dan dinyatakan memenuhi standar mutu kurikulum OBE Universitas PGRI Ronggolawe (UNIROW) Tuban.`
   - Tanggal: `Divalidasi di Tuban, pada tanggal: [Tanggal Validasi]`.

---

## 5. TABEL 3 — Lembar Pengesahan / Validasi
Struktur: **2 baris**, 2 kolom simetris (`tblGrid` widths: `7569, 7569` dxa, Shading `#FAFAFA`).

- **Baris 0**:
  - Kolom 0 (Kiri): `Menyetujui,\nKetua Program Studi [Nama Prodi]` (Center, Bold).
  - Kolom 1 (Kanan): `Mengetahui,\nUnit Jaminan Mutu (UJM) Prodi [Nama Prodi]` (Center, Bold).
- **Baris 1**:
  - Kolom 0 (Kiri): Ruang tanda tangan digital / basah, lalu:\n`    [Nama Lengkap Kaprodi, Gelar]\n    NIDN. [NIDN Kaprodi]` (Left, Bold).
  - Kolom 1 (Kanan): Ruang tanda tangan digital / basah, lalu:\n`    [Nama Tim Penjamin Mutu / UJM, Gelar]\n    NIDN. [NIDN UJM]` (Left, Bold).
