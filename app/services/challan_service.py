from datetime import datetime
from decimal import Decimal
import pandas as pd
import os
from app.services.excel_service import (
    get_sheet_rows,
    append_sheet_row,
)
from app.services.stock_service import _decimal, StockServiceError


class ChallanServiceError(Exception):
    """Challan processing or validation error."""
    pass


def get_setting(key, default_val, workbook_path=None):
    """Helper to fetch config from Settings sheet."""
    path = workbook_path or os.path.join('data', 'om_enterprise.xlsx')
    if not os.path.exists(path):
        alt_path = 'om_enterprise.xlsx'
        if os.path.exists(alt_path):
            path = alt_path
        else:
            return str(default_val)
    
    try:
        xls = pd.ExcelFile(path)
        if 'Settings' in xls.sheet_names:
            df = pd.read_excel(xls, 'Settings')
            if not df.empty and 'Setting_Key' in df.columns:
                match = df.loc[df['Setting_Key'] == key, 'Setting_Value']
                if not match.empty:
                    val = str(match.values[0]).strip()
                    if val and val != 'nan':
                        return val
    except Exception:
        pass
    return str(default_val)


def generate_challan_number(challan_type, party_type, workbook_path=None):
    """
    Generate dynamic challan number based on client rules:
    - GST Issue Challan: OM/26-27/180, OM/26-27/181...
    - GST Receive Challan: 01, 02, 03...
    - Without-GST Issue Challan: 01, 02, 03...
    - Without-GST Receive Challan: 01, 02, 03...
    """
    is_gst = (party_type.upper() == "GST")
    is_issue = (challan_type.upper() == "ISSUE")
    
    sheet_name = "Issue_Challan" if is_issue else "Receive_Challan"
    
    try:
        rows = get_sheet_rows(sheet_name, workbook_path=workbook_path)
    except Exception:
        rows = []

    max_num = 0

    if is_issue and is_gst:
        # GST Issue Challan Format: OM/26-27/180+
        default_prefix = get_setting("GST_ISSUE_CHALLAN_PREFIX", "OM/26-27/", workbook_path)
        default_start = int(get_setting("GST_ISSUE_CHALLAN_START", "180", workbook_path))
        max_num = default_start - 1

        for row in rows:
            c_no = str(row.get("Challan_No") or "").strip()
            if default_prefix in c_no:
                parts = c_no.split("/")
                for p in reversed(parts):
                    if p.isdigit():
                        num = int(p)
                        if num > max_num:
                            max_num = num
                        break
        next_num = max_num + 1
        return f"{default_prefix}{next_num}"

    else:
        # Numeric format: 01, 02, 03... for GST Receive, Without-GST Issue & Receive
        if is_issue:
            default_start = int(get_setting("WITHOUT_GST_ISSUE_CHALLAN_START", "1", workbook_path))
        else:
            if is_gst:
                default_start = int(get_setting("GST_RECEIVE_CHALLAN_START", "1", workbook_path))
            else:
                default_start = int(get_setting("WITHOUT_GST_RECEIVE_CHALLAN_START", "1", workbook_path))
        
        max_num = default_start - 1

        for row in rows:
            c_no = str(row.get("Challan_No") or "").strip()
            if c_no.isdigit():
                num = int(c_no)
                if num > max_num:
                    max_num = num
            else:
                digits = "".join(filter(str.isdigit, c_no))
                if digits:
                    num = int(digits)
                    if num > max_num and len(c_no) <= 5:
                        max_num = num

        next_num = max_num + 1
        return f"{next_num:02d}"


def create_issue_challan(data, items_list, workbook_path=None):
    party_type = str(data.get("party_type") or "GST").strip().upper()
    
    # Form ke kisi bhi purane ya default challan_no ko ignore karke hamesha naya generate karein
    challan_no = generate_challan_number("ISSUE", party_type, workbook_path=workbook_path)

    party_id = str(data.get("party_id") or "").strip()
    site_id = str(data.get("site_id") or "").strip()
    challan_date = str(data.get("date") or datetime.now().strftime("%d-%m-%Y")).strip()

    if not party_id or not site_id:
        raise ChallanServiceError("Party and Site are required.")

    if not items_list:
        raise ChallanServiceError("At least one item is required in Issue Challan.")

    challan_record = {
        "Challan_No": challan_no,
        "Challan_Date": challan_date,
        "Party_ID": party_id,
        "Site_ID": site_id,
        "Party_Type": party_type,
        "Transporter_Name": data.get("transporter", "SELF"),
        "Vehicle_No": data.get("vehicle_no", ""),
        "EWay_Bill_No": data.get("eway_bill_no", ""),
        "Remarks": data.get("remarks", ""),
        "Status": "ACTIVE",
        "Driver_Name": data.get("driver_name", ""),
        "Driver_Mob": data.get("driver_mob", ""),
        "Driver_Licence": data.get("driver_licence", ""),
        "Weight": data.get("weight", ""),
        "Goods_Value": data.get("goods_value", "0.00")
    }
    
    try:
        append_sheet_row("Issue_Challan", challan_record, workbook_path=workbook_path)
    except Exception as e:
        raise ChallanServiceError(f"Failed to save issue challan record: {str(e)}")

    for item in items_list:
        item_id = str(item.get("item_id") or "").strip()
        qty = _decimal(item.get("quantity"), "Quantity")

        if qty <= 0:
            continue

        ledger_entry = {
            "Reference_ID": challan_no,
            "Transaction_Date": challan_date,
            "Transaction_Type": "ISSUE",
            "Party_ID": party_id,
            "Site_ID": site_id,
            "Item_ID": item_id,
            "Usable_Qty_Change": float(-qty),
            "Damaged_Qty_Change": 0.0,
        }
        append_sheet_row("StockLedger", ledger_entry, workbook_path=workbook_path)

    return {"status": "success", "challan_no": challan_no}


def create_receive_challan(data, items_list, breakage_list=None, workbook_path=None):
    party_type = str(data.get("party_type") or "GST").strip().upper()
    
    # Form ke kisi bhi purane challan_no ko ignore karke hamesha naya generate karein
    challan_no = generate_challan_number("RECEIVE", party_type, workbook_path=workbook_path)

    party_id = str(data.get("party_id") or "").strip()
    site_id = str(data.get("site_id") or "").strip()
    challan_date = str(data.get("date") or datetime.now().strftime("%d-%m-%Y")).strip()

    if not party_id or not site_id:
        raise ChallanServiceError("Party and Site are required.")

    for item in items_list:
        item_id = str(item.get("item_id") or "").strip()
        qty = _decimal(item.get("quantity"), "Quantity")

        if qty <= 0:
            continue

        ledger_entry = {
            "Reference_ID": challan_no,
            "Transaction_Date": challan_date,
            "Transaction_Type": "RECEIVE",
            "Party_ID": party_id,
            "Site_ID": site_id,
            "Item_ID": item_id,
            "Usable_Qty_Change": float(qty),
            "Damaged_Qty_Change": 0.0,
        }
        append_sheet_row("StockLedger", ledger_entry, workbook_path=workbook_path)

    if breakage_list:
        for b in breakage_list:
            item_id = str(b.get("item_id") or "").strip()
            qty = _decimal(b.get("quantity"), "Breakage Qty")
            rate = _decimal(b.get("rate"), "Breakage Rate")

            if qty <= 0:
                continue

            damage_ledger_entry = {
                "Reference_ID": challan_no,
                "Transaction_Date": challan_date,
                "Transaction_Type": "BREAKAGE",
                "Party_ID": party_id,
                "Site_ID": site_id,
                "Item_ID": item_id,
                "Usable_Qty_Change": 0.0,
                "Damaged_Qty_Change": float(qty),
            }
            append_sheet_row("StockLedger", damage_ledger_entry, workbook_path=workbook_path)

    return {"status": "success", "challan_no": challan_no}