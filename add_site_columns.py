
from openpyxl import load_workbook
from pathlib import Path
import shutil

WORKBOOK_PATH = Path(r"E:\OM_ENTERPRISE_ERP\data\om_enterprise.xlsx")
BACKUP_PATH = WORKBOOK_PATH.with_name("om_enterprise_backup.xlsx")

if not WORKBOOK_PATH.exists():
    raise FileNotFoundError(f"Workbook nahi mili: {WORKBOOK_PATH}")

# Original workbook ki backup banayein, agar backup pehle se nahi hai.
if not BACKUP_PATH.exists():
    shutil.copy2(WORKBOOK_PATH, BACKUP_PATH)
    print(f"Backup ban gaya: {BACKUP_PATH}")

wb = load_workbook(WORKBOOK_PATH)

if "SiteLocations" not in wb.sheetnames:
    raise ValueError("Workbook mein 'SiteLocations' sheet nahi mili.")

ws = wb["SiteLocations"]

# Existing headers read karein.
headers = [
    str(ws.cell(row=1, column=col).value).strip()
    if ws.cell(row=1, column=col).value is not None
    else ""
    for col in range(1, ws.max_column + 1)
]

# Sirf missing headers add karein; existing columns ko nahi hatayein.
for required_header in ("State", "State_Code"):
    if required_header not in headers:
        new_col = ws.max_column + 1
        ws.cell(row=1, column=new_col).value = required_header
        headers.append(required_header)
        print(f"Column add hua: {required_header}")
    else:
        print(f"Column pehle se maujood hai: {required_header}")

wb.save(WORKBOOK_PATH)
wb.close()

print("SiteLocations headers successfully update ho gaye.")
