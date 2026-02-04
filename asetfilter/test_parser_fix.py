
import sys
import os
import pandas as pd
import logging

# Setup logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Add current dir to path
sys.path.append(os.getcwd())

from parser import parse_excel_file

FILEPATH = r"d:\DISPERKIMHUB-ASSET DATA\PRESENTASI.xls"

if not os.path.exists(FILEPATH):
    print(f"File not found at {FILEPATH}")
    # Try alternate location if needed, or ask user?
    # Assuming user has it there as per previous context.
    # Try uploads folder if not found
    FILEPATH = r"d:\DISPERKIMHUB-ASSET DATA\asetfilter\uploads\PRESENTASI.xls"

if not os.path.exists(FILEPATH):
    print(f"File not found at {FILEPATH}")
    sys.exit(1)

print(f"Testing parsing of: {FILEPATH}")
try:
    df, stats = parse_excel_file(FILEPATH)
    print("\n--- Stats ---")
    for k, v in stats.items():
        print(f"{k}: {v}")
    
    print(f"\nDataFrame Shape: {df.shape}")
    if not df.empty:
        print("First 3 rows:")
        print(df.head(3).to_string())
    else:
        print("DataFrame is empty!")

except Exception as e:
    print(f"Error: {e}")
    import traceback
    traceback.print_exc()
