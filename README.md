# RPS UNIROW — Generator Rencana Pembelajaran Semester OBE 2026

"Artificial intelligence module (custom skill) for Google Antigravity, Cursor IDE, and Claude Code to compile, populate, and format Semester Learning Plan (RPS) documents compliant with the standard format and quality standards of Universitas PGRI Ronggolawe (UNIROW) Tuban, based on the 2026 OBE (Outcome-Based Education) curriculum.

---

## 🌟 Fitur Utama & Standar Mutu

1. **Format Baku A4 Landscape (100% Sesuai Template Resmi)**:
   - Orientasi **A4 Landscape** penuh (`w="16838" h="11909"` dxa).
   - Margin halaman presisi: Atas/Bawah `39.6 pt`, Kiri/Kanan `43.2 pt`.
   - Tipografi elegan **Cambria** (Body 10.5 pt, Header 11 pt, Judul 12 pt bold).
   - Garis tabel single-line warna `#B0B5B3` (sz 4 / 0.5 pt) dengan padding sel proporsional.

2. **4 Komponen Anatomi Dokumen Lengkap**:
   - **Tabel 0: Identitas, Otorisasi, CPL-CPMK & Korelasi (20 Baris)**:
     - KOP resmi universitas & prodi, nama MK, kode, SKS, semester, tanggal penyusunan.
     - Pejabat pengembang RPS, Koordinator RMK, dan Kaprodi lengkap dengan NIDN/gelar.
     - CPL Prodi, CPMK 1–4, Sub-CPMK 1–14 berlabel taksonomi Bloom (misal: `[C4, A3]`).
     - **Tabel Bersarang (Nested Table 18 Baris)**: Matriks korelasi CPL terhadap Sub-CPMK dengan centang (`✓`) dan persentase bobot presisi 100%.
     - Deskripsi mata kuliah, bahan kajian, pustaka ber-ISBN, dan dosen pengampu.
   - **Tabel 1: Matriks Rencana Pembelajaran 16 Minggu (19 Baris × 8 Kolom)**:
     - Sesi perkuliahan terstruktur 16 minggu dengan beban belajar SN-Dikti Permendikbudristek No. 53/2023.
     - Minggu ke-8: **Ujian Tengah Semester (UTS)** berbobot 25% (shading `#F4F6F7`).
     - Minggu ke-16: **Ujian Akhir Semester (UAS)** berbobot 25% (shading `#F4F6F7`).
   - **Tabel 2: Rubrik Penilaian Holistik (5 Baris × 5 Kolom)**:
     - 4 kriteria: Pemahaman Konsep (40%), Analisis Kritis (30%), Aplikasi Konteks Lokal (20%), Sistematika & Komunikasi (10%).
   - **Bagian Evaluasi, Skala 7, Soal UAS & Lembar Pengesahan**:
     - Standar mutu konversi nilai skala 7 resmi UNIROW (A s.d. E, batas kelulusan C 56.00).
     - 5 paket butir soal UAS berbasis HOTS sesuai Sub-CPMK.
     - Kartu pengesahan/validasi oleh Tim Pengembang Kurikulum.

---

## 📁 Struktur Berkas Skill

```text
rps-unirow/
├── SKILL.md                          # Definisi instruksi skill & panduan operasional AI
├── README.md                         # Dokumentasi repository publik
├── assets/
│   ├── Template_RPS_UNIROW.docx      # Master template Word A4 Landscape resmi
│   └── Template_RPS_UNIROW_OLD.docx  # Template pembanding arsip DIKTI
├── references/
│   └── placeholder_map.md            # Peta pemetaan placeholder & sel tabel
└── scripts/
    └── build_rps.py                  # Engine generator otomatis dokumen Word (.docx)
```

---

## 🚀 Prasyarat & Instalasi

### 1. Prasyarat Sistem
- Python 3.8 atau lebih baru.
- Library `python-docx`:
  ```bash
  pip install python-docx
  ```

### 2. Cara Pemasangan di Google Antigravity

#### Opsi A: Pasang di Proyek Tertentu (Workspace)
Buka terminal di root workspace Anda, lalu clone:
```bash
git clone https://github.com/mariofahmi/skillISORPS.git .agents/skills/rps-unirow
```

#### Opsi B: Pasang Secara Global (Berlaku untuk Semua Proyek)
Clone ke folder konfigurasi global Antigravity:
* **Windows (PowerShell)**:
  ```powershell
  git clone https://github.com/mariofahmi/skillISORPS.git "$env:USERPROFILE\.gemini\config\skills\rps-unirow"
  ```
* **Linux / macOS**:
  ```bash
  git clone https://github.com/mariofahmi/skillISORPS.git ~/.gemini/config/skills/rps-unirow
  ```

---

## 💻 Cara Penggunaan

### 1. Melalui Chat AI Antigravity
Cukup ketik perintah di obrolan Antigravity:
> `"/rps-unirow buatkan RPS mata kuliah Hukum Tata Negara 4 SKS untuk Semester 3 kurikulum OBE 2026"`

AI akan secara otomatis memetakan CPL, merumuskan CPMK & Sub-CPMK berlabel Bloom, menyusun silabus 16 minggu, serta menghasilkan dokumen Word `.docx` siap cetak.

### 2. Melalui CLI (Terminal Python Langsung)
```powershell
py ".agents\skills\rps-unirow\scripts\build_rps.py"
```

### 3. Melalui Simulator & Portal Web Interaktif
Kunjungi live demo di: **[https://mariofahmi.github.io/skillISORPS/](https://mariofahmi.github.io/skillISORPS/)**
- **Unduh Word (.docx)**: Ekspor dokumen OpenXML `.docx` asli kurikulum OBE 2026 secara instan langsung di peramban (client-side via JSZip engine).
- **Cetak / Ekspor PDF A4 Landscape**: Tata letak WYSIWYG A4 Landscape 100% presisi untuk langsung dicetak atau disimpan ke format PDF.
- **Salin Markdown RPS**: Salin naskah ringkasan terstruktur untuk integrasi cepat ke LMS / catatan dosen.


---

## 📄 Lisensi & Kontributor
- **Pengembang**: Mario Fahmi Syahrial, M.Pd. (Program Studi PPKn, FKIP UNIROW Tuban)
- **Kompatibilitas**: Google Antigravity, Antigravity 2.0, Cursor IDE, Claude Code
