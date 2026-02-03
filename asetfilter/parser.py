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
    'Jenis Barang / Nama Barang': 'nama_asset',
    'Nomor': 'nomor',
    'Luas (m2)': 'luas',
    'Tahun': 'tahun',
    'Status Tanah': 'status_tanah',
    'Penggunaan': 'penggunaan',
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
    'LAIN-LAIN': 'lain_lain'
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
    
    Args:
        row: Pandas Series representing a row
        
    Returns:
        True if this is a valid data row
    """
    # Check if row is mostly empty
    non_null_count = row.notna().sum()
    if non_null_count < 3:
        return False
    
    # Convert row to string for keyword checking
    row_str = ' '.join(str(v).upper() for v in row.values if pd.notna(v))
    
    # Skip summary/total rows
    for keyword in SKIP_KEYWORDS:
        if keyword in row_str:
            return False
    
    return True


def find_header_row(df: pd.DataFrame) -> int:
    """
    Find the row containing column headers.
    
    Looks for rows containing key column names like "Jenis Barang" or "Nama Barang".
    
    Args:
        df: Raw DataFrame from Excel
        
    Returns:
        Row index of headers, or -1 if not found
    """
    search_terms = ['Jenis Barang', 'Nama Barang', 'Satuan Kerja', 'KECAMATAN']
    
    for idx, row in df.iterrows():
        row_str = ' '.join(str(v) for v in row.values if pd.notna(v))
        
        matches = sum(1 for term in search_terms if term in row_str)
        if matches >= 2:
            return idx
    
    return -1


def map_columns(df: pd.DataFrame) -> dict:
    """
    Create mapping from actual column indices to normalized field names.
    
    Args:
        df: DataFrame with header row as column names
        
    Returns:
        Dictionary mapping column index to field name
    """
    column_map = {}
    
    for idx, col_name in enumerate(df.columns):
        if pd.isna(col_name):
            continue
            
        col_str = str(col_name).strip()
        
        # Try exact match first
        if col_str in COLUMN_MAPPING:
            column_map[idx] = COLUMN_MAPPING[col_str]
            continue
        
        # Try partial match
        for excel_col, field_name in COLUMN_MAPPING.items():
            if excel_col.lower() in col_str.lower() or col_str.lower() in excel_col.lower():
                column_map[idx] = field_name
                break
    
    return column_map


def parse_excel_file(filepath: str) -> Tuple[pd.DataFrame, dict]:
    """
    Parse the Excel file and return cleaned DataFrame.
    
    This is the main entry point for Excel parsing. It handles:
    1. Reading the Excel file
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
        # Determine file type and read accordingly
        if filepath.endswith('.xls'):
            xls = pd.ExcelFile(filepath, engine='xlrd')
        else:
            xls = pd.ExcelFile(filepath, engine='openpyxl')
        
        stats['sheets_processed'] = xls.sheet_names
        
        for sheet_name in xls.sheet_names:
            logger.info(f"Processing sheet: {sheet_name}")
            
            try:
                # Read sheet without headers first
                df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)
                stats['total_rows_read'] += len(df_raw)
                
                # Find header row
                header_row = find_header_row(df_raw)
                
                if header_row == -1:
                    logger.warning(f"Could not find header row in sheet: {sheet_name}")
                    # Try using row 6 as default (common in this file format)
                    header_row = 6
                
                # Read again with proper headers
                df = pd.read_excel(xls, sheet_name=sheet_name, header=header_row)
                
                # Create column mapping
                col_map = map_columns(df)
                
                if not col_map:
                    logger.warning(f"Could not map columns in sheet: {sheet_name}")
                    continue
                
                # Process each row
                for idx, row in df.iterrows():
                    if not is_data_row(row):
                        stats['skipped_rows'] += 1
                        continue
                    
                    # Create record dictionary
                    record = {}
                    
                    for col_idx, field_name in col_map.items():
                        if col_idx < len(row):
                            record[field_name] = row.iloc[col_idx]
                    
                    # Skip if no asset name
                    if not record.get('nama_asset') or pd.isna(record.get('nama_asset')):
                        stats['skipped_rows'] += 1
                        continue
                    
                    # Clean specific fields
                    record['luas'] = clean_luas_value(record.get('luas'))
                    record['nilai_harga'] = clean_nilai_value(record.get('nilai_harga'))
                    record['tahun'] = clean_tahun_value(record.get('tahun'))
                    
                    # Combine status fields
                    record['status_combined'] = combine_status_fields(record)
                    
                    # Clean string fields
                    for field in ['nama_asset', 'kecamatan', 'satuan_kerja', 'status_tanah', 
                                 'catatan', 'k3', 'pemetaan', 'tanah_bangunan', 'kode_aset']:
                        if field in record and pd.notna(record[field]):
                            record[field] = str(record[field]).strip()
                        else:
                            record[field] = None
                    
                    all_data.append(record)
                    stats['valid_rows'] += 1
                    
            except Exception as e:
                error_msg = f"Error processing sheet {sheet_name}: {str(e)}"
                logger.error(error_msg)
                stats['errors'].append(error_msg)
        
        # Create final DataFrame
        if all_data:
            result_df = pd.DataFrame(all_data)
            
            # Ensure all expected columns exist
            expected_cols = list(set(COLUMN_MAPPING.values())) + ['status_combined']
            for col in expected_cols:
                if col not in result_df.columns:
                    result_df[col] = None
            
            return result_df, stats
        else:
            return pd.DataFrame(), stats
            
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
