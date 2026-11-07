from __future__ import annotations

import os
import threading
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from app.config import Config

WORKBOOK_PATH = Path(Config.WORKBOOK_PATH)
_WORKBOOK_LOCK = threading.RLock()
HEADER_COLOR = "18345C"

SCHEMAS: dict[str, list[str]] = {
    "Customers": [
        "Party_ID", "Party_Name", "Party_Type", "GSTIN", "Phone",
        "Address", "State", "State_Code", "Is_Active", "Created_At", "Updated_At",
    ],
    "SiteLocations": [
        "Site_ID", "Party_ID", "Site_Name", "Site_Address", "Contact_Person",
        "Phone", "Is_Active", "Created_At", "Updated_At",
    ],
    "Items": [
        "Item_ID", "Item_Code", "Item_Name", "Size", "Category", "Description",
        "Unit", "Unit_Weight_KG", "Daily_Rental_Rate", "Goods_Value_Rate",
        "Is_Active", "Created_At", "Updated_At",
    ],
    "Issue_Challan": [
        "Issue_ID", "Challan_No", "Party_ID", "Site_ID", "Challan_Date",
        "Party_Type", "Transporter_Name", "Vehicle_No", "EWay_Bill_No",
        "Remarks", "Status", "Created_At", "Created_By", "Cancel_Reason",
        "Driver_Name", "Driver_Mob", "Weight", "Goods_Value", "Driver_Licence",
    ],
    "Issue_Challan_Items": [
        "Line_ID", "Issue_ID", "Item_ID", "Quantity", "Unit_Weight_KG",
        "Total_Weight_KG", "Goods_Value_Rate", "Approx_Goods_Value",
    ],
    "Receive_Challan": [
        "Receive_ID", "Receive_Challan_No", "Issue_ID", "Party_ID", "Site_ID",
        "Receive_Date", "Remarks", "Status", "Created_At", "Created_By", "Cancel_Reason",
    ],
    "Receive_Challan_Items": [
        "Line_ID", "Receive_ID", "Issue_Line_ID", "Item_ID", "Returned_Qty",
        "Damaged_Qty", "Accepted_Qty", "Damage_Rate", "Damage_Amount", "Damage_Status",
    ],
    "StockLedger": [
        "Ledger_ID", "Transaction_Date", "Party_ID", "Site_ID", "Item_ID",
        "Transaction_Type", "Reference_ID", "Usable_Qty_Change", "Damaged_Qty_Change",
        "Unit_Weight_KG", "Weight_Change_KG", "Goods_Value_Rate", "Goods_Value_Change",
        "Reason", "Created_At", "Created_By",
    ],
    "Monthly_Bills": [
        "Bill_ID", "Bill_No", "Party_ID", "Site_ID", "Period_Start", "Period_End",
        "Bill_Date", "Taxable_Value", "CGST_Rate", "CGST_Amount", "SGST_Rate",
        "SGST_Amount", "Total_Amount", "Round_Off", "Net_Total", "Status",
        "Created_At", "Created_By", "Cancel_Reason",
    ],
    "Monthly_Bill_Lines": [
        "Bill_Line_ID", "Bill_ID", "Item_ID", "Item_Name_Snapshot", "Date_From",
        "Date_To", "Quantity", "Rental_Days", "Daily_Rate", "Line_Amount",
        "Source_Issue_ID", "Source_Receive_ID", "Line_Type", "Damage_Charge_ID",
    ],
    "Damage_Charges": [
        "Damage_Charge_ID", "Receive_ID", "Party_ID", "Site_ID", "Item_ID",
        "Damaged_Qty", "Damage_Rate", "Damage_Amount", "Charge_Date", "Status",
        "Billed_Bill_ID", "Created_At", "Created_By",
    ],
    "Audit_Log": [
        "Audit_ID", "Event_Time", "Username", "Action", "Entity_Type",
        "Entity_ID", "Reason", "Before_Data", "After_Data",
    ],
    "Settings": [
        "Setting_Key", "Setting_Value", "Description",
    ],
    "Document_Index": [
        "Document_ID", "Document_Type", "Reference_ID", "Document_Number",
        "File_Path", "Created_At", "Status",
    ],
    "OpeningStock": [
        "Opening_Stock_ID", "Location_Type", "Party_ID", "Site_ID", "Item_ID",
        "Estimated_Qty", "Estimated_Date", "Estimated_Notes", "Verified_Qty",
        "Verified_Date", "Verified_Notes", "Status", "Created_At", "Updated_At",
    ],
}

class ExcelServiceError(Exception):
    pass

def _format_header(cell) -> None:
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = PatternFill(fill_type="solid", fgColor=HEADER_COLOR)
    cell.alignment = Alignment(vertical="center", wrap_text=True)

def _format_sheet(ws, headers: list[str]) -> None:
    for column_number, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=column_number)
        cell.value = header
        _format_header(cell)
        ws.column_dimensions[get_column_letter(column_number)].width = min(max(len(header) + 3, 15), 28)
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    ws.row_dimensions[1].height = 30
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"

def _save_atomic(wb, path: Path) -> None:
    temporary_path = path.with_name(f"{path.stem}.tmp{path.suffix}")
    try:
        wb.save(temporary_path)
        os.replace(temporary_path, path)
    except Exception as exc:
        if temporary_path.exists():
            try:
                temporary_path.unlink()
            except OSError:
                pass
        raise ExcelServiceError(f"Could not save workbook: {path}.") from exc

def initialize_workbook(workbook_path: str | Path | None = None) -> Path:
    path = Path(workbook_path or WORKBOOK_PATH)
    path.parent.mkdir(parents=True, exist_ok=True)
    with _WORKBOOK_LOCK:
        if path.exists() and path.stat().st_size > 0:
            try:
                wb = load_workbook(path)
            except Exception as exc:
                raise ExcelServiceError(f"Cannot open workbook safely: {path}.") from exc
        else:
            wb = Workbook()
            wb.remove(wb.active)
        changed = False
        try:
            for sheet_name, headers in SCHEMAS.items():
                if sheet_name not in wb.sheetnames:
                    ws = wb.create_sheet(sheet_name)
                    _format_sheet(ws, headers)
                    changed = True
                    continue
                ws = wb[sheet_name]
                if ws.max_row == 1 and ws.max_column == 1 and ws["A1"].value is None:
                    _format_sheet(ws, headers)
                    changed = True
                    continue
                existing_headers = [ws.cell(row=1, column=column).value for column in range(1, ws.max_column + 1)]
                for header in headers:
                    if header in existing_headers:
                        continue
                    next_column = ws.max_column + 1
                    cell = ws.cell(row=1, column=next_column, value=header)
                    _format_header(cell)
                    ws.column_dimensions[get_column_letter(next_column)].width = min(max(len(header) + 3, 15), 28)
                    existing_headers.append(header)
                    changed = True
                ws.freeze_panes = "A2"
                ws.sheet_view.showGridLines = False
                ws.row_dimensions[1].height = 30
                ws.auto_filter.ref = ws.dimensions
            
            if "Settings" in wb.sheetnames:
                settings_ws = wb["Settings"]
                existing_keys = {settings_ws.cell(row=row, column=1).value for row in range(2, settings_ws.max_row + 1)}
                defaults = [
                    ("COMPANY_NAME", "OM ENTERPRISE", "Company display name"),
                    ("WORKBOOK_SCHEMA_VERSION", "1", "ERP workbook schema version"),
                ]
                for key, value, description in defaults:
                    if key in existing_keys:
                        continue
                    settings_ws.append([key, value, description])
                    existing_keys.add(key)
                    changed = True

            if changed or not path.exists() or path.stat().st_size == 0:
                _save_atomic(wb, path)
        finally:
            wb.close()
    return path

def get_sheet_rows(sheet_name: str, workbook_path: str | Path | None = None) -> list[dict[str, Any]]:
    if sheet_name not in SCHEMAS:
        raise ExcelServiceError(f"Unknown sheet: {sheet_name}")
    path = Path(workbook_path or WORKBOOK_PATH)
    initialize_workbook(path)
    with _WORKBOOK_LOCK:
        try:
            wb = load_workbook(path, read_only=True, data_only=True)
            try:
                ws = wb[sheet_name]
                rows = ws.iter_rows(values_only=True)
                headers = next(rows, None)
                if not headers:
                    return []
                header_names = [str(header) if header is not None else "" for header in headers]
                result = []
                for row in rows:
                    if not any(value is not None for value in row):
                        continue
                    record = {header_names[index]: (row[index] if index < len(row) else None) for index in range(len(header_names)) if header_names[index]}
                    result.append(record)
                return result
            finally:
                wb.close()
        except ExcelServiceError:
            raise
        except Exception as exc:
            raise ExcelServiceError(f"Could not read sheet '{sheet_name}'.") from exc

def append_row(sheet_name: str, row_data: dict[str, Any], workbook_path: str | Path | None = None) -> int:
    if sheet_name not in SCHEMAS:
        raise ExcelServiceError(f"Unknown sheet: {sheet_name}")
    path = initialize_workbook(workbook_path)
    with _WORKBOOK_LOCK:
        wb = None
        try:
            wb = load_workbook(path)
            ws = wb[sheet_name]
            headers = [ws.cell(row=1, column=column).value for column in range(1, ws.max_column + 1)]
            unknown_fields = set(row_data) - set(headers)
            if unknown_fields:
                raise ExcelServiceError(f"Unknown field(s) for {sheet_name}: {', '.join(sorted(unknown_fields))}")
            ws.append([row_data.get(header) if header is not None else None for header in headers])
            row_number = ws.max_row
            _save_atomic(wb, path)
            return row_number
        except ExcelServiceError:
            raise
        except Exception as exc:
            raise ExcelServiceError(f"Could not append data to '{sheet_name}'.") from exc
        finally:
            if wb is not None:
                wb.close()

def append_sheet_row(sheet_name: str, row_data: dict[str, Any], workbook_path: str | Path | None = None) -> int:
    return append_row(sheet_name, row_data, workbook_path)
