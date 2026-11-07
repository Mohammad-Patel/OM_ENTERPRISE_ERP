from decimal import Decimal, InvalidOperation
from pathlib import Path

from app.services.excel_service import (
    WORKBOOK_PATH,
    ExcelServiceError,
    get_sheet_rows,
    append_sheet_row,
)


class StockServiceError(Exception):
    """Stock service validation or processing error."""
    pass


def _decimal(value, label):
    if value is None or str(value).strip() == "":
        return Decimal("0")

    try:
        number = Decimal(str(value).strip())
    except (InvalidOperation, ValueError):
        raise StockServiceError(
            f"{label} must be a valid number."
        )

    if not number.is_finite():
        raise StockServiceError(
            f"{label} must be a finite number."
        )

    return number


def list_stock_ledger(
    item_id="",
    party_id="",
    site_id="",
    transaction_type="",
    workbook_path=None,
):
    """Return StockLedger records with optional filters."""
    try:
        records = get_sheet_rows(
            "StockLedger",
            workbook_path=workbook_path,
        )
    except Exception as exc:
        raise StockServiceError(
            f"Could not read StockLedger: {exc}"
        ) from exc

    item_id = str(item_id or "").strip()
    party_id = str(party_id or "").strip()
    site_id = str(site_id or "").strip()
    transaction_type = str(
        transaction_type or ""
    ).strip().upper()

    result = []

    for record in records:
        if item_id and str(record.get("Item_ID") or "") != item_id:
            continue

        if party_id and str(record.get("Party_ID") or "") != party_id:
            continue

        if site_id and str(record.get("Site_ID") or "") != site_id:
            continue

        if transaction_type:
            current_type = str(
                record.get("Transaction_Type") or ""
            ).strip().upper()

            if current_type != transaction_type:
                continue

        result.append(record)

    return result


def get_item_stock_balance(
    item_id,
    party_id=None,
    site_id=None,
    workbook_path=None,
):
    """
    Calculate usable and damaged quantity from signed ledger changes.
    """
    item_id = str(item_id or "").strip()

    if not item_id:
        raise StockServiceError("Item ID is required.")

    records = list_stock_ledger(
        item_id=item_id,
        party_id=party_id or "",
        site_id=site_id or "",
        workbook_path=workbook_path,
    )

    usable_balance = Decimal("0")
    damaged_balance = Decimal("0")

    for record in records:
        usable_balance += _decimal(
            record.get("Usable_Qty_Change"),
            "Usable Quantity Change",
        )

        damaged_balance += _decimal(
            record.get("Damaged_Qty_Change"),
            "Damaged Quantity Change",
        )

    return {
        "Item_ID": item_id,
        "Party_ID": party_id or "",
        "Site_ID": site_id or "",
        "Usable_Qty": float(usable_balance),
        "Damaged_Qty": float(damaged_balance),
        "Total_Qty": float(
            usable_balance + damaged_balance
        ),
    }


def get_godown_stock_summary(workbook_path=None):
    """
    Get current available stock in Godown item-wise.
    Godown stock has empty Party_ID and Site_ID or specific Godown flags.
    """
    try:
        items = get_sheet_rows("Items", workbook_path=workbook_path)
    except Exception:
        items = []

    summary = []
    for item in items:
        item_id = str(item.get("Item_ID") or item.get("ID") or "")
        if not item_id:
            continue
        
        # Calculate balance where Party/Site is empty (Godown inventory)
        records = list_stock_ledger(item_id=item_id, workbook_path=workbook_path)
        usable = Decimal("0")
        damaged = Decimal("0")

        for r in records:
            # If transaction is inside godown (Party/Site not set or type is Opening/Godown adjustment)
            p_id = str(r.get("Party_ID") or "").strip()
            s_id = str(r.get("Site_ID") or "").strip()
            
            # Ledger tracks overall movement. Godown stock = Opening + Receive(In) - Issue(Out)
            usable += _decimal(r.get("Usable_Qty_Change"), "Usable")
            damaged += _decimal(r.get("Damaged_Qty_Change"), "Damaged")

        summary.append({
            "Item_ID": item_id,
            "Item_Name": item.get("Item_Name") or item.get("Name") or "",
            "Usable_Qty": float(usable),
            "Damaged_Qty": float(damaged),
            "Total_Qty": float(usable + damaged)
        })

    return summary