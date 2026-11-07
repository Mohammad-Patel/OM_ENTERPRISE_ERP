from __future__ import annotations

import json
import re
import uuid
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

from app.services.excel_service import (
    WORKBOOK_PATH,
    ExcelServiceError,
    append_row,
    get_sheet_rows,
    initialize_workbook,
)

class ItemServiceError(Exception):
    pass


DEFAULT_ITEMS = [
    ("CUPLOCK/VERTICAL 3 MTR", "3 MTR", "Vertical", 13.0),
    ("CUPLOCK/VERTICAL 2.5 MTR", "2.5 MTR", "Vertical", 10.83),
    ("CUPLOCK/VERTICAL 2 MTR", "2 MTR", "Vertical", 8.66),
    ("CUPLOCK/VERTICAL 1.5 MTR", "1.5 MTR", "Vertical", 6.5),
    ("CUPLOCK/VERTICAL 1 MTR", "1 MTR", "Vertical", 4.33),
    ("CUPLOCK/VERTICAL 0.5 MTR", "0.5 MTR", "Vertical", 2.16),
    ("LEDGER/HORIZONTAL 2 MTR (CENTER TO CENTER)", "2 MTR", "Ledger", 6.5),
    ("LEDGER/HORIZONTAL 1.8 MTR (CENTER TO CENTER)", "1.8 MTR", "Ledger", 6.0),
    ("LEDGER/HORIZONTAL 1.5 MTR (CENTER TO CENTER)", "1.5 MTR", "Ledger", 5.0),
    ("LEDGER/HORIZONTAL 1.2 MTR (CENTER TO CENTER)", "1.2 MTR", "Ledger", 4.0),
    ("LEDGER/HORIZONTAL 1.0 MTR (CENTER TO CENTER)", "1.0 MTR", "Ledger", 3.4),
    ("LEDGER/HORIZONTAL 915 MM (CENTER TO CENTER)", "915 MM", "Ledger", 3.1),
    ("LEDGER/HORIZONTAL 600 MM (CENTER TO CENTER)", "600 MM", "Ledger", 2.0),
    ("PROPS 3*3", "3*3", "Props", 26.0),
    ("PROPS 2*3", "2*3", "Props", 21.0),
    ("PROPS 2*2", "2*2", "Props", 18.0),
    ("S/PLATE 3*2", "3*2", "S/Plate", 22.0),
    ("S/PLATE 3*1.5", "3*1.5", "S/Plate", 16.5),
    ("S/PLATE 3*1", "3*1", "S/Plate", 11.0),
    ("ST. CHANNEL 10'", "10 FT", "Steel Channel", 23.0),
    ("ST. CHANNEL 9'", "9 FT", "Steel Channel", 21.0),
    ("ST. CHANNEL 8'", "8 FT", "Steel Channel", 19.0),
    ("U-JACK", "", "Jack", 3.2),
    ("BASE JACK", "", "Jack", 3.2),
    ("JOINTER PIN (SPIGOT PIN)", "", "Accessory", 0.65),
    ("CLAMP", "", "Accessory", 0.85),
    ("PIPE 6 MTR", "6 MTR", "Pipe", 20.0),
    ("ACROW SPAN", "", "Accessory", 38.0),
    ("STEEL CHALLI", "", "Accessory", 15.0),
]


def _now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _normalise_name(value):
    return " ".join(str(value or "").strip().split())


def _decimal(value, label, required=False):
    if value is None or str(value).strip() == "":
        if required:
            raise ItemServiceError(f"{label} is required.")
        return None
    try:
        number = Decimal(str(value).strip())
    except (InvalidOperation, ValueError):
        raise ItemServiceError(f"{label} must be a valid number.")
    if not number.is_finite() or number < 0:
        raise ItemServiceError(f"{label} cannot be negative.")
    return float(number)


def _items():
    try:
        return get_sheet_rows("Items")
    except (ExcelServiceError, Exception) as exc:
        raise ItemServiceError(f"Could not read Items sheet: {exc}") from exc


def list_items(search="", active_only=False):
    query = str(search or "").strip().casefold()
    records = _items()
    result = []
    for item in records:
        if active_only and str(item.get("Is_Active") or "YES").upper() not in ("YES", "Y", "TRUE", "1"):
            continue
        if query and not any(query in str(item.get(field) or "").casefold()
                             for field in ("Item_Code", "Item_Name", "Size", "Category")):
            continue
        result.append(item)
    return sorted(result, key=lambda item: str(item.get("Item_Code") or ""))


def get_item(item_id):
    target = str(item_id or "").strip()
    for item in _items():
        if str(item.get("Item_ID") or "") == target:
            return item
    return None


def _validate(data, exclude_id=None):
    name = _normalise_name(data.get("Item_Name"))
    if not name:
        raise ItemServiceError("Item Name is required.")

    code = str(data.get("Item_Code") or "").strip().upper()
    for item in _items():
        if exclude_id and str(item.get("Item_ID") or "") == str(exclude_id):
            continue
        if str(item.get("Item_Name") or "").strip().casefold() == name.casefold():
            raise ItemServiceError("An item with this name already exists.")
        if code and str(item.get("Item_Code") or "").strip().upper() == code:
            raise ItemServiceError("Item Code already exists.")

    weight = _decimal(data.get("Unit_Weight_KG"), "Standard Weight per Piece (kg)", required=True)
    rate = _decimal(data.get("Daily_Rental_Rate"), "Default Daily Rental Rate")
    goods_rate = _decimal(data.get("Goods_Value_Rate"), "Goods Value Rate")

    return {
        "Item_Name": name,
        "Size": str(data.get("Size") or "").strip(),
        "Category": str(data.get("Category") or "").strip(),
        "Description": str(data.get("Description") or "").strip(),
        "Unit": str(data.get("Unit") or "Nos").strip() or "Nos",
        "Unit_Weight_KG": weight,
        "Daily_Rental_Rate": rate,
        "Goods_Value_Rate": goods_rate,
    }


def _audit(action, item_id, before=None, after=None, reason=""):
    audit_id = "AUD-" + uuid.uuid4().hex[:10].upper()
    append_row("Audit_Log", {
        "Audit_ID": audit_id,
        "Event_Time": _now(),
        "Username": "Administrator",
        "Action": action,
        "Entity_Type": "Item",
        "Entity_ID": str(item_id),
        "Reason": reason,
        "Before_Data": json.dumps(before, default=str, ensure_ascii=False) if before else "",
        "After_Data": json.dumps(after, default=str, ensure_ascii=False) if after else "",
    })


def create_item(data):
    initialize_workbook()
    values = _validate(data)
    existing = _items()
    next_number = 1
    for item in existing:
        match = re.fullmatch(r"ITM-(\d+)", str(item.get("Item_Code") or "").strip().upper())
        if match:
            next_number = max(next_number, int(match.group(1)) + 1)

    item_id = "ITEM-" + uuid.uuid4().hex[:12].upper()
    item_code = f"ITM-{next_number:04d}"
    record = {
        "Item_ID": item_id,
        "Item_Code": item_code,
        **values,
        "Is_Active": "YES",
        "Created_At": _now(),
        "Updated_At": _now(),
    }
    try:
        append_row("Items", record)
        _audit("CREATE", item_id, after=record)
    except Exception as exc:
        raise ItemServiceError(f"Could not save item: {exc}") from exc
    return item_id


def update_item(item_id, data, reason=""):
    current = get_item(item_id)
    if current is None:
        raise ItemServiceError("Item was not found.")
    values = _validate(data, exclude_id=item_id)

    try:
        from openpyxl import load_workbook
        from app.services.excel_service import _WORKBOOK_LOCK, _save_atomic

        path = initialize_workbook()
        with _WORKBOOK_LOCK:
            wb = load_workbook(path)
            try:
                ws = wb["Items"]
                headers = [ws.cell(1, col).value for col in range(1, ws.max_column + 1)]
                header_col = {name: i + 1 for i, name in enumerate(headers) if name}
                row_number = None
                for row in range(2, ws.max_row + 1):
                    if str(ws.cell(row, header_col["Item_ID"]).value or "") == str(item_id):
                        row_number = row
                        break
                if row_number is None:
                    raise ItemServiceError("Item was not found.")
                updated = dict(current)
                updated.update(values)
                updated["Updated_At"] = _now()
                for key, value in updated.items():
                    if key in header_col and key not in ("Item_ID", "Item_Code", "Created_At"):
                        ws.cell(row_number, header_col[key]).value = value
                _save_atomic(wb, path)
            finally:
                wb.close()
        _audit("UPDATE", item_id, before=current, after=updated, reason=reason)
        return True
    except ItemServiceError:
        raise
    except Exception as exc:
        raise ItemServiceError(f"Could not update item: {exc}") from exc


def delete_item(item_id, reason=""):
    current = get_item(item_id)
    if current is None:
        raise ItemServiceError("Item was not found.")

    # Do not delete records referenced by challans or stock.
    reference_checks = (
        ("Issue_Challan_Items", "Item_ID"),
        ("Receive_Challan_Items", "Item_ID"),
        ("StockLedger", "Item_ID"),
    )
    for sheet, column in reference_checks:
        try:
            references = get_sheet_rows(sheet)
        except Exception:
            references = []
        if any(str(row.get(column) or "") == str(item_id) for row in references):
            raise ItemServiceError(
                "This item is already referenced by stock/challan records. "
                "It cannot be deleted."
            )

    try:
        from openpyxl import load_workbook
        from app.services.excel_service import _WORKBOOK_LOCK, _save_atomic

        path = initialize_workbook()
        with _WORKBOOK_LOCK:
            wb = load_workbook(path)
            try:
                ws = wb["Items"]
                headers = [ws.cell(1, col).value for col in range(1, ws.max_column + 1)]
                id_col = headers.index("Item_ID") + 1
                for row in range(2, ws.max_row + 1):
                    if str(ws.cell(row, id_col).value or "") == str(item_id):
                        ws.delete_rows(row, 1)
                        _save_atomic(wb, path)
                        _audit("DELETE", item_id, before=current, reason=reason)
                        return True
            finally:
                wb.close()
        raise ItemServiceError("Item was not found.")
    except ItemServiceError:
        raise
    except Exception as exc:
        raise ItemServiceError(f"Could not delete item: {exc}") from exc


def seed_default_items():
    """Insert the 29 supplied items once; existing names are left untouched."""
    existing_names = {
        str(item.get("Item_Name") or "").strip().casefold()
        for item in _items()
    }
    added = 0
    for name, size, category, weight in DEFAULT_ITEMS:
        if name.casefold() in existing_names:
            continue
        create_item({
            "Item_Name": name,
            "Size": size,
            "Category": category,
            "Description": "",
            "Unit": "Nos",
            "Unit_Weight_KG": weight,
            "Daily_Rental_Rate": "",
            "Goods_Value_Rate": "",
        })
        existing_names.add(name.casefold())
        added += 1
    return added
