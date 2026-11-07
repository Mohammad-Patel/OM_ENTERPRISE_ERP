from datetime import datetime
from decimal import Decimal
from app.services.excel_service import (
    get_sheet_rows,
    append_sheet_row,
)
from app.services.stock_service import _decimal, StockServiceError


class ChallanServiceError(Exception):
    """Challan processing or validation error."""
    pass


def create_issue_challan(data, items_list, workbook_path=None):
    challan_no = str(data.get("challan_no") or "").strip()
    party_id = str(data.get("party_id") or "").strip()
    site_id = str(data.get("site_id") or "").strip()
    challan_date = str(data.get("date") or datetime.now().strftime("%d-%m-%Y")).strip()

    if not challan_no or not party_id or not site_id:
        raise ChallanServiceError("Challan No, Party, and Site are required.")

    if not items_list:
        raise ChallanServiceError("At least one item is required in Issue Challan.")

    challan_record = {
        "Challan_No": challan_no,
        "Challan_Date": challan_date,
        "Party_ID": party_id,
        "Site_ID": site_id,
        "Party_Type": data.get("party_type", "GST"),
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
    except Exception:
        pass

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
    challan_no = str(data.get("challan_no") or "").strip()
    party_id = str(data.get("party_id") or "").strip()
    site_id = str(data.get("site_id") or "").strip()
    challan_date = str(data.get("date") or datetime.now().strftime("%d-%m-%Y")).strip()

    if not challan_no or not party_id or not site_id:
        raise ChallanServiceError("Challan No, Party, and Site are required.")

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
