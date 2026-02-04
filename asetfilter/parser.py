"""
AsetFilter - Excel Parser Module

This module handles the complex parsing of the PRESENTASI.xls file which contains:
- Multiple rows of headers and merged cells
- Data from various government departments
- Inconsistent formatting in numeric fields
- Multiple status columns that need to be combined

The parser normalizes all data into a clean, consistent format for database storage.
"""
import pandas as pd
import numpy as np
import re
from typing import Tuple, List, Optional
import logging

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


# Column name mapping from Excel to normalized field names
COLUMN_MAPPING = {
    'NO. KIB 2023': 'no_kib',
    'No.': 'no_urut',
    'Kode Lokasi': 'kode_lokasi',
    'Satuan Kerja': 'satuan_kerja',
    'Jenis Barang / Nama Barang': 'jenis_barang_nama_barang',  # Keep orig col but don't map to nama_asset
    'Nomor': 'nomor',
    'Luas (m2)': 'luas',
    'Tahun': 'tahun',
    'Status Tanah': 'status_tanah',
    'Penggunaan': 'nama_asset',  # Map Penggunaan to nama_asset as requested
    'Asal Usul': 'asal_usul',
    'Nilai / Harga': 'nilai_harga',
    'Keterangan': 'keterangan',
    'Kode Aset': 'kode_aset',
    'JUMLAH BIDANG': 'jumlah_bidang',
    'KECAMATAN': 'kecamatan',
    'PEMETAAN ASET TANAH': 'pemetaan',
    'CATATAN (TERMANFAATKAN/TERLANTAR)': 'catatan',
    'K3 (MILIK WARGA/ADA KLAIM, TKD, DLL)': 'k3',
    'TANAH (BANGUNAN/TANAH KOSONG)': 'tanah_bangunan',
    'LAIN-LAIN': 'lain_lain',
    'Letak/Alamat': 'alamat',
    'Letak / Alamat': 'alamat', # Handle variations
    'Location/Address': 'alamat' # Handle variations requested by user
}

# Keywords to identify summary/total rows that should be skipped
SKIP_KEYWORDS = ['JUMLAH', 'TOTAL', 'SUB TOTAL', 'GRAND TOTAL', 'REKAPITULASI']

# Status keywords for combined status field
STATUS_KEYWORDS = [
    'TERMANFAATKAN', 'TERLANTAR', 'BERSERTIFIKAT', 'TKD', 
    'BELUM TERPETAKAN', 'SUDAH TERPETAKAN', 'HAK PAKAI',
    'MILIK WARGA', 'ADA KLAIM', 'BANGUNAN', 'TANAH KOSONG',
    'BELUM DISURVEY', 'SUDAH PEMAPARAN'
]


def clean_luas_value(value) -> Optional[float]:
    """
    Clean and convert Luas (area) values to float.
    
    Handles various formats found in the Excel:
    - Standard numbers: 1500.00, 1500
    - Time-like formats: "6153:00:00" (should be 6153)
    - Text with numbers: "1500 m2"
    - Empty/NaN values
    
    Args:
        value: The raw value from Excel
        
    Returns:
        Float value or None if conversion fails
    """
    if pd.isna(value):
        return None
    
    # Convert to string for processing
    str_val = str(value).strip()
    
    if not str_val:
        return None
    
    try:
        # Handle time-like format "6153:00:00" -> 6153
        if ':' in str_val:
            # Extract the first part before any colon
            parts = str_val.split(':')
            str_val = parts[0]
        
        # Remove any non-numeric characters except decimal point and minus
        cleaned = re.sub(r'[^\d.\-]', '', str_val)
        
        if cleaned:
            return float(cleaned)
        return None
        
    except (ValueError, TypeError):
        logger.warning(f"Could not convert luas value: {value}")
        return None


def clean_nilai_value(value) -> Optional[float]:
    """
    Clean and convert Nilai/Harga (price) values to float.
    
    Args:
        value: The raw value from Excel
        
    Returns:
        Float value or None if conversion fails
    """
    if pd.isna(value):
        return None
    
    str_val = str(value).strip()
    
    if not str_val:
        return None
    
    try:
        # Remove currency symbols, commas, spaces
        cleaned = re.sub(r'[Rp.,\s]', '', str_val)
        cleaned = re.sub(r'[^\d.\-]', '', cleaned)
        
        if cleaned:
            return float(cleaned)
        return None
        
    except (ValueError, TypeError):
        logger.warning(f"Could not convert nilai value: {value}")
        return None


def clean_tahun_value(value) -> Optional[int]:
    """
    Clean and convert Tahun (year) values to integer.
    
    Args:
        value: The raw value from Excel
        
    Returns:
        Integer year or None if conversion fails
    """
    if pd.isna(value):
        return None
    
    try:
        # Handle float values like 1999.0
        if isinstance(value, float):
            return int(value)
        
        str_val = str(value).strip()
        
        # Extract 4-digit year
        match = re.search(r'(\d{4})', str_val)
        if match:
            year = int(match.group(1))
            # Validate reasonable year range
            if 1900 <= year <= 2100:
                return year
        
        return None
        
    except (ValueError, TypeError):
        return None


def combine_status_fields(row: dict) -> str:
    """
    Combine multiple status fields into a single searchable string.
    
    Combines:
    - Status Tanah (e.g., "Hak Pakai")
    - CATATAN (e.g., "TERMANFAATKAN")
    - K3 (e.g., "TKD")
    - PEMETAAN ASET TANAH (e.g., "BELUM TERPETAKAN")
    - TANAH (e.g., "BANGUNAN")
    
    Args:
        row: Dictionary containing the status fields
        
    Returns:
        Combined status string with unique values
    """
    status_parts = []
    
    status_fields = ['status_tanah', 'catatan', 'k3', 'pemetaan', 'tanah_bangunan']
    
    for field in status_fields:
        value = row.get(field)
        if value and not pd.isna(value):
            str_val = str(value).strip().upper()
            if str_val and str_val not in ['NAN', 'NONE', '-']:
                status_parts.append(str_val)
    
    # Remove duplicates while preserving order
    seen = set()
    unique_parts = []
    for part in status_parts:
        if part not in seen:
            seen.add(part)
            unique_parts.append(part)
    
    return ' | '.join(unique_parts) if unique_parts else ''


def is_data_row(row: pd.Series) -> bool:
    """
    Determine if a row contains actual data (not headers, totals, or empty rows).
    """
    # Check if row is mostly empty - require at least 3 non-empty values
    # Since we used fillna(''), we must check for empty strings, not just notna()
    non_null_count = sum(1 for v in row.values if pd.notna(v) and str(v).strip() != '')
    
    if non_null_count < 3:
        return False
    
    # Check for header-like rows (containing known header keywords)
    row_str = ' '.join(str(v).upper() for v in row.values if pd.notna(v) and str(v).strip() != '')
    
    # Valid data row override
    if 'BEDA' in row_str: 
       pass

    # Specific override for known valid row starting with BEDA
    first_val_check = str(row.iloc[0]).strip().upper() if pd.notna(row.iloc[0]) else ''
    if first_val_check == 'BEDA':
        return True

    if 'LETAK' in row_str and 'ALAMAT' in row_str:
        return False
    if 'PENGADAAN' in row_str and 'HAK' in row_str:
        return False
    
    # Only check FIRST column for skip keywords 
    first_val = str(row.iloc[0]).strip().upper() if pd.notna(row.iloc[0]) else ''
    for keyword in SKIP_KEYWORDS:
        if first_val.startswith(keyword):
            return False
    
    return True


def find_header_row(df: pd.DataFrame) -> int:
    """Find the row containing column headers."""
    search_terms = ['Jenis Barang', 'Nama Barang', 'Satuan Kerja', 'KECAMATAN']
    
    for idx, row in df.iterrows():
        # Check non-empty values
        row_str = ' '.join(str(v) for v in row.values if pd.notna(v) and str(v).strip() != '')
        
        matches = sum(1 for term in search_terms if term in row_str)
        if matches >= 2:
            return idx
    
    return -1


def map_columns(df: pd.DataFrame) -> dict:
    """Create mapping from actual column indices to normalized field names."""
    column_map = {}
    
    # Check if we have secondary headers in the first row
    has_secondary_headers = False
    if not df.empty:
        first_row_str = ' '.join(str(v).upper() for v in df.iloc[0].values if pd.notna(v) and str(v).strip() != '')
        if 'LETAK' in first_row_str and 'ALAMAT' in first_row_str:
            has_secondary_headers = True
            
    for idx in range(len(df.columns)):
        # Get candidate names
        candidates = []
        
        # 1. Primary header
        col_name = df.columns[idx]
        if pd.notna(col_name) and str(col_name).strip() != '':
            candidates.append(str(col_name).strip())
            
        # 2. Secondary header
        if has_secondary_headers and not df.empty:
            val = df.iloc[0, idx]
            if pd.notna(val) and str(val).strip() != '':
                candidates.append(str(val).strip())
        
        # Try to match any candidate to our mapping
        mapped_field = None
        for candidate in candidates:
            # Try exact match
            if candidate in COLUMN_MAPPING:
                mapped_field = COLUMN_MAPPING[candidate]
                break
            
            # Try partial match logic (fallback)
            for excel_col, field_name in COLUMN_MAPPING.items():
                if excel_col.lower() in candidate.lower() or candidate.lower() in excel_col.lower():
                    mapped_field = field_name
                    break
            
            if mapped_field:
                break
        
        if mapped_field:
            column_map[idx] = mapped_field
            
    return column_map


def parse_excel_file(filepath: str) -> Tuple[pd.DataFrame, dict]:
    """
    Parse the Excel file and return cleaned DataFrame.
    
    This is the main entry point for Excel parsing. It handles:
    1. Reading the Excel file using calamine engine
    2. Finding and processing headers
    3. Cleaning and normalizing data
    4. Combining status fields
    
    Args:
        filepath: Path to the Excel file
        
    Returns:
        Tuple of (cleaned DataFrame, parsing statistics dict)
    """
    stats = {
        'total_rows_read': 0,
        'valid_rows': 0,
        'skipped_rows': 0,
        'sheets_processed': [],
        'errors': []
    }
    
    all_data = []
    
    try:
        # Use calamine engine for better performance and accuracy
        # Read the file headerless initially to find the header row
        logger.info(f"Reading {filepath} using calamine engine")
        
        # Read sheet 'A' (or first sheet)
        # Using keep_default_na=False to prevent Pandas from converting empty strings to NaN automatically, 
        # but we also want to explicitly handle empty cells.
        # User requested: "Ensure there is no .dropna() function... Use df.fillna('')"
        try:
            df_raw = pd.read_excel(filepath, sheet_name='A', header=None, engine='calamine')
        except Exception as e:
            logger.info(f"Sheet 'A' not found or error ({e}), trying index 0")
            df_raw = pd.read_excel(filepath, sheet_name=0, header=None, engine='calamine')
            
        # Fill NaN with empty string IMMEDIATELY
        # Convert to object type first to allow mixed types (strings in float columns)
        df_raw = df_raw.astype(object)
        df_raw.fillna('', inplace=True)
        
        stats['total_rows_read'] = len(df_raw)
        stats['sheets_processed'] = ['Sheet A/0']
        
        # Find header row
        header_row_idx = find_header_row(df_raw)
        
        if header_row_idx == -1:
            logger.warning("Could not find header row, defaulting to 6")
            header_row_idx = 6
            
        # Set columns from header row
        headers = df_raw.iloc[header_row_idx]
        df_data = df_raw.iloc[header_row_idx + 1:].copy()
        
        # Clean headers to strings
        df_data.columns = [str(h).strip() for h in headers]
        
        col_map = map_columns(df_data)
        
        if not col_map:
            logger.error("Could not map columns")
            stats['errors'].append("Could not map columns")
            return pd.DataFrame(), stats

        skipped_by_filter = 0
        
        # Process data
        for idx, row in df_data.iterrows():
            if not is_data_row(row):
                stats['skipped_rows'] += 1
                skipped_by_filter += 1
                continue
            
            record = {}
            for col_idx, field_name in col_map.items():
                # col_idx in map_columns corresponds to integer index of the column
                # Since we sliced the dataframe, the columns are now named.
                # BUT map_columns returns indices like {0: 'nama_asset', 1: '...'} relative to the dataframe columns?
                # Let's check map_columns again. 
                # It iterates range(len(df.columns)).
                # So if we use df.iloc[idx], we need positional access.
                if col_idx < len(df_data.columns):
                     val = row.iloc[col_idx]
                     record[field_name] = val
            
            # Data Cleaning & Validation
            
            # Clean kecamatan
            kecamatan = record.get('kecamatan')
            if kecamatan is not None and str(kecamatan).strip() in ['0', '-', '']:
                kecamatan = None
                record['kecamatan'] = None
            
            # Handle blank names
            nama = record.get('nama_asset')
            if not nama or str(nama).strip() == '':
                 if kecamatan:
                     record['nama_asset'] = '(Tanpa Nama)'
                 else:
                     stats['skipped_rows'] += 1
                     continue

            # Clean Numeric Fields
            if 'luas' in record:
                record['luas'] = clean_luas_value(record['luas'])
            if 'nilai_harga' in record:
                record['nilai_harga'] = clean_nilai_value(record['nilai_harga'])
            if 'tahun' in record:
                record['tahun'] = clean_tahun_value(record['tahun'])
            
            # Combine status
            record['status_combined'] = combine_status_fields(record)
            
            # Clean string fields
            for field in ['nama_asset', 'kecamatan', 'satuan_kerja', 'status_tanah', 
                         'catatan', 'k3', 'pemetaan', 'tanah_bangunan', 'kode_aset']:
                if field in record and record[field] is not None:
                    record[field] = str(record[field]).strip()
            
            all_data.append(record)
            
        stats['valid_rows'] = len(all_data)
        logger.info(f"Parsing complete. Valid rows: {len(all_data)}")
        
        return pd.DataFrame(all_data), stats
            
    except Exception as e:
        error_msg = f"Error reading Excel file: {str(e)}"
        logger.error(error_msg)
        stats['errors'].append(error_msg)
        return pd.DataFrame(), stats


def get_unique_values(df: pd.DataFrame, column: str) -> List[str]:
    """
    Get sorted unique non-null values from a column.
    
    Args:
        df: DataFrame
        column: Column name
        
    Returns:
        List of unique values
    """
    if column not in df.columns:
        return []
    
    values = df[column].dropna().unique()
    return sorted([str(v).strip() for v in values if str(v).strip()])


def get_status_options(df: pd.DataFrame) -> List[str]:
    """
    Extract unique status values from the combined status field.
    
    Args:
        df: DataFrame
        
    Returns:
        List of unique status keywords
    """
    if 'status_combined' not in df.columns:
        return STATUS_KEYWORDS
    
    all_statuses = set()
    
    for value in df['status_combined'].dropna():
        parts = str(value).split('|')
        for part in parts:
            cleaned = part.strip().upper()
            if cleaned and cleaned not in ['NAN', 'NONE', '-']:
                all_statuses.add(cleaned)
    
    return sorted(list(all_statuses))


def get_luas_range(df: pd.DataFrame) -> Tuple[float, float]:
    """
    Get min and max values for Luas column.
    
    Args:
        df: DataFrame
        
    Returns:
        Tuple of (min_luas, max_luas)
    """
    if 'luas' not in df.columns:
        return (0, 0)
    
    luas_values = df['luas'].dropna()
    
    if len(luas_values) == 0:
        return (0, 0)
    
    return (float(luas_values.min()), float(luas_values.max()))
