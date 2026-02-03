"""
AsetFilter - Database Models
"""
from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()


class Asset(db.Model):
    """Asset model representing a single asset record from the Excel file"""
    __tablename__ = 'assets'
    
    id = db.Column(db.Integer, primary_key=True)
    
    # Core identification
    no_kib = db.Column(db.String(50))  # NO. KIB 2023
    no_urut = db.Column(db.Integer)  # No.
    kode_lokasi = db.Column(db.String(50))  # Kode Lokasi
    kode_aset = db.Column(db.String(100))  # Kode Aset
    
    # Asset details
    satuan_kerja = db.Column(db.String(200))  # Satuan Kerja (Department)
    nama_asset = db.Column(db.String(500))  # Jenis Barang / Nama Barang
    nomor = db.Column(db.String(100))  # Nomor
    
    # Physical attributes
    luas = db.Column(db.Float)  # Luas (m2)
    tahun = db.Column(db.Integer)  # Tahun
    
    # Location
    kecamatan = db.Column(db.String(100))  # KECAMATAN
    alamat = db.Column(db.String(500))  # Letak/Alamat if available
    
    # Status fields
    status_tanah = db.Column(db.String(100))  # Status Tanah
    catatan = db.Column(db.String(200))  # CATATAN (TERMANFAATKAN/TERLANTAR)
    k3 = db.Column(db.String(200))  # K3 (MILIK WARGA/ADA KLAIM, TKD, DLL)
    pemetaan = db.Column(db.String(100))  # PEMETAAN ASET TANAH
    tanah_bangunan = db.Column(db.String(100))  # TANAH (BANGUNAN/TANAH KOSONG)
    
    # Combined status for filtering (derived field)
    status_combined = db.Column(db.String(500))
    
    # Financial
    nilai_harga = db.Column(db.Float)  # Nilai / Harga
    asal_usul = db.Column(db.String(100))  # Asal Usul
    penggunaan = db.Column(db.String(200))  # Penggunaan
    
    # Additional
    jumlah_bidang = db.Column(db.Integer)  # JUMLAH BIDANG
    keterangan = db.Column(db.Text)  # Keterangan
    lain_lain = db.Column(db.Text)  # LAIN-LAIN
    
    def __repr__(self):
        return f'<Asset {self.id}: {self.nama_asset[:50] if self.nama_asset else "Unknown"}>'
    
    def to_dict(self):
        """Convert asset to dictionary for JSON/export"""
        return {
            'id': self.id,
            'no_kib': self.no_kib,
            'no_urut': self.no_urut,
            'kode_lokasi': self.kode_lokasi,
            'kode_aset': self.kode_aset,
            'satuan_kerja': self.satuan_kerja,
            'nama_asset': self.nama_asset,
            'luas': self.luas,
            'tahun': self.tahun,
            'kecamatan': self.kecamatan,
            'status_tanah': self.status_tanah,
            'catatan': self.catatan,
            'k3': self.k3,
            'pemetaan': self.pemetaan,
            'tanah_bangunan': self.tanah_bangunan,
            'status_combined': self.status_combined,
            'nilai_harga': self.nilai_harga,
            'asal_usul': self.asal_usul,
            'penggunaan': self.penggunaan,
            'keterangan': self.keterangan,
            'lain_lain': self.lain_lain
        }
    
    def to_export_dict(self):
        """Convert asset to dictionary for export with columns matching original Excel layout"""
        return {
            'NO. KIB': self.no_kib,
            'No.': self.no_urut,
            'Kode Lokasi': self.kode_lokasi,
            'Satuan Kerja': self.satuan_kerja,
            'Jenis Barang / Nama Barang': self.nama_asset,
            'Nomor': self.nomor,
            'Luas (m2)': self.luas,
            'Tahun': self.tahun,
            'Status Tanah': self.status_tanah,
            'Penggunaan': self.penggunaan,
            'Asal Usul': self.asal_usul,
            'Nilai / Harga': self.nilai_harga,
            'Keterangan': self.keterangan,
            'Kode Aset': self.kode_aset,
            'JUMLAH BIDANG': self.jumlah_bidang,
            'KECAMATAN': self.kecamatan,
            'PEMETAAN ASET TANAH': self.pemetaan,
            'CATATAN (TERMANFAATKAN/TERLANTAR)': self.catatan,
            'K3 (MILIK WARGA/ADA KLAIM, TKD, DLL)': self.k3,
            'TANAH (BANGUNAN/TANAH KOSONG)': self.tanah_bangunan,
            'LAIN-LAIN': self.lain_lain
        }


class UploadHistory(db.Model):
    """Track upload history"""
    __tablename__ = 'upload_history'
    
    id = db.Column(db.Integer, primary_key=True)
    filename = db.Column(db.String(255))
    uploaded_at = db.Column(db.DateTime, default=db.func.now())
    records_count = db.Column(db.Integer)
    status = db.Column(db.String(50))  # success, failed
    error_message = db.Column(db.Text)
    
    def __repr__(self):
        return f'<Upload {self.filename} at {self.uploaded_at}>'
