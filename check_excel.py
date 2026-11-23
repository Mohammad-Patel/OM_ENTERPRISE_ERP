import os
import pandas as pd

excel_path = 'om_enterprise.xlsx'
if not os.path.exists(excel_path):
    excel_path = 'data/om_enterprise.xlsx'

if os.path.exists(excel_path):
    xls = pd.ExcelFile(excel_path)
    print("=" * 50)
    print(f"FILE FOUND: {excel_path}")
    print("=" * 50)
    print("SHEETS IN EXCEL:", xls.sheet_names)
    print("-" * 50)
    
    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet)
        print(f"\nSheet Name: [{sheet}]")
        print(f"Total Rows: {len(df)}")
        print("Columns:", df.columns.tolist())
    print("\n" + "=" * 50)
else:
    print("Error: Excel file nahi mili! Path check karein.")