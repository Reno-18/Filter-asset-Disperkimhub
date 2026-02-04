
import pandas as pd
import numpy as np

try:
    df = pd.DataFrame({'A': [1.0, 2.0, np.nan]})
    print("Original types:")
    print(df.dtypes)
    
    print("Filling with empty string...")
    try:
        df.fillna('', inplace=True)
        print("Success!")
        print(df)
        print("New types:")
        print(df.dtypes)
    except Exception as e:
        print(f"Error during fillna: {e}")

except Exception as e:
    print(f"General Error: {e}")
