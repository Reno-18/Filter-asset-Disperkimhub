# AsetFilter

Aplikasi web Flask untuk parsing, filtering, dan export data aset pemerintah dari file Excel yang kompleks.

## Fitur Utama

- **Upload Excel** - Upload file .xls/.xlsx dengan drag & drop
- **Parsing Otomatis** - Bersihkan data messy dari format Excel kompleks
- **Filter Powerful** - 4 jenis filter: text search, dropdown, range, multi-select
- **Tabel Interaktif** - Sorting per kolom, pagination 20 data/halaman
- **Export Data** - Export hasil filter ke CSV atau Excel
- **UI Modern** - Bootstrap 5 dengan sidebar navigation

## Kolom yang Diekstrak

| Kolom Excel | Field Database |
|-------------|----------------|
| Jenis Barang / Nama Barang | nama_asset |
| KECAMATAN | kecamatan |
| Luas (m2) | luas |
| Satuan Kerja | satuan_kerja |
| Status Tanah | status_tanah |
| CATATAN (TERMANFAATKAN/TERLANTAR) | catatan |
| K3 | k3 |
| PEMETAAN ASET TANAH | pemetaan |
| Nilai / Harga | nilai_harga |
| Kode Aset | kode_aset |
| Tahun | tahun |

## Instalasi

### 1. Clone/Download Project
```powershell
cd "d:\DISPERKIMHUB-ASSET DATA\asetfilter"
```

### 2. Buat Virtual Environment
```powershell
python -m venv venv
.\venv\Scripts\activate
```

### 3. Install Dependencies
```powershell
pip install -r requirements.txt
```

### 4. Jalankan Aplikasi
```powershell
python app.py
```

Aplikasi akan berjalan di: **http://127.0.0.1:5000**

## Cara Penggunaan

### Upload Data
1. Buka halaman Upload (`/upload`)
2. Drag & drop file Excel atau klik untuk memilih file
3. Tunggu proses parsing selesai
4. Lihat preview data yang berhasil diproses

### Filter Data
1. Buka Dashboard (`/`)
2. Gunakan filter yang tersedia:
   - **Nama Asset**: Ketik untuk pencarian partial match
   - **Kecamatan**: Pilih dari dropdown
   - **Luas**: Masukkan range min-max
   - **Status**: Pilih satu atau beberapa status
3. Klik "Terapkan Filter"
4. Klik header kolom untuk sorting

### Export Data
1. Terapkan filter sesuai kebutuhan
2. Klik "Export CSV" atau "Export Excel"
3. File akan terdownload otomatis

## Struktur Project

```
/asetfilter
├── app.py              # Main Flask application
├── config.py           # Configuration settings
├── models.py           # SQLAlchemy models
├── forms.py            # Flask-WTF forms
├── parser.py           # Excel parsing logic
├── requirements.txt    # Python dependencies
├── README.md           # Documentation
├── uploads/            # Temporary upload folder
├── static/
│   ├── css/
│   │   └── style.css   # Custom styles
│   └── js/
│       └── app.js      # Custom JavaScript
└── templates/
    ├── base.html       # Base template
    ├── index.html      # Dashboard page
    ├── upload.html     # Upload page
    ├── 404.html        # Error page
    └── 500.html        # Error page
```

## Cara Kerja Parser Excel

Parser (`parser.py`) menangani file Excel dengan struktur kompleks:

1. **Deteksi Header Row** - Mencari baris yang mengandung "Jenis Barang" atau "Nama Barang"
2. **Skip Baris Summary** - Lewati baris dengan keyword JUMLAH, TOTAL, dll
3. **Normalisasi Kolom** - Map nama kolom Excel ke nama field standar
4. **Clean Nilai Luas** - Handle format aneh seperti "6153:00:00" → 6153
5. **Combine Status** - Gabungkan Status Tanah, CATATAN, K3, PEMETAAN jadi satu field

## API Endpoints

| Method | Route | Deskripsi |
|--------|-------|-----------|
| GET | / | Dashboard dengan filter dan tabel |
| GET | /upload | Halaman upload |
| POST | /upload | Proses upload file |
| POST | /filter | Filter data (AJAX) |
| GET | /export-csv | Export ke CSV |
| GET | /export-excel | Export ke Excel |
| POST | /clear-data | Hapus semua data |
| GET | /api/stats | Statistik data |

## Teknologi

- **Backend**: Flask 3.x, Flask-SQLAlchemy, Flask-WTF
- **Database**: SQLite
- **Frontend**: Bootstrap 5, Bootstrap Icons
- **Data Processing**: pandas, openpyxl, xlrd

## Troubleshooting

### Error "No data could be extracted"
- Pastikan format file sesuai dengan struktur PRESENTASI.xls
- Periksa apakah kolom header ada di file

### Error upload file
- Pastikan file berekstensi .xls atau .xlsx
- Ukuran file maksimal 10MB

### Data tidak lengkap
- Beberapa baris mungkin terlewat karena:
  - Tidak memiliki "Nama Asset"
  - Merupakan baris summary/total
  - Format kolom tidak sesuai

## Lisensi

MIT License - Bebas digunakan dan dimodifikasi.
