import pandas as pd
import sys

filename = sys.argv[1]
try:
    sheets = pd.read_excel(filename, sheet_name=None, engine="odf")
    for sheet_name, df in sheets.items():
        print(f"--- Sheet: {sheet_name} ---")
        print(df.head(20).to_string())
        print()
except Exception as e:
    print(f"Error: {e}")
