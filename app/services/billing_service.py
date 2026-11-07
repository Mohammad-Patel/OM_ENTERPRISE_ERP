from datetime import datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP
from app.services.stock_service import list_stock_ledger, _decimal
from app.services.excel_service import get_sheet_rows


class BillingServiceError(Exception):
    """Billing service processing or calculation error."""
    pass


def parse_date(date_str):
    """Parse string date DD-MM-YYYY to datetime.date object."""
    try:
        return datetime.strptime(str(date_str).strip(), "%d-%m-%Y").date()
    except Exception:
        try:
            return datetime.strptime(str(date_str).strip(), "%Y-%m-%d").date()
        except Exception as exc:
            raise BillingServiceError(f"Invalid date format '{date_str}'. Expected DD-MM-YYYY.") from exc


def generate_monthly_bill(party_id, site_id, start_date_str, end_date_str, workbook_path=None):
    """
    Generates structured line-item breakdown for monthly rental invoice.
    Calculates day-wise duration, total numbers, item rates, GST, and net total.
    """
    party_id = str(party_id or "").strip()
    site_id = str(site_id or "").strip()

    if not party_id or not site_id:
        raise BillingServiceError("Party ID and Site ID are required for billing.")

    start_date = parse_date(start_date_str)
    end_date = parse_date(end_date_str)

    if start_date > end_date:
        raise BillingServiceError("Start date cannot be after End date.")

    # 1. Fetch Item Master with Rates
    items_rows = get_sheet_rows("Items", workbook_path=workbook_path)
    items_map = {}
    for r in items_rows:
        i_id = str(r.get("Item_ID") or r.get("ID") or "").strip()
        if i_id:
            items_map[i_id] = {
                "name": r.get("Item_Name") or r.get("Name") or "",
                "rate": _decimal(r.get("Rent_Rate") or r.get("Rate") or "0", "Rent Rate")
            }

    # 2. Fetch Ledger Entries for this Party and Site
    ledger = list_stock_ledger(party_id=party_id, site_id=site_id, workbook_path=workbook_path)

    # 3. Process Ledger per Item to calculate running balances and period splits
    bill_items = []
    item_counter = 1

    # Group ledger transactions by Item_ID
    items_ledger_map = {}
    for entry in ledger:
        i_id = str(entry.get("Item_ID") or "").strip()
        if i_id not in items_ledger_map:
            items_ledger_map[i_id] = []
        items_ledger_map[i_id].append(entry)

    for item_id, entries in items_ledger_map.items():
        item_info = items_map.get(item_id, {"name": f"Item {item_id}", "rate": Decimal("0")})
        daily_rate = item_info["rate"]

        # Sort entries by date
        sorted_entries = sorted(entries, key=lambda x: parse_date(x.get("Date")))

        # Calculate opening stock before start_date
        opening_qty = Decimal("0")
        period_txns = []

        for e in sorted_entries:
            e_date = parse_date(e.get("Date"))
            # Outstanding stock at site increases with ISSUE, decreases with RECEIVE
            usable_chg = _decimal(e.get("Usable_Qty_Change"), "Qty")
            net_site_change = -usable_chg  # In stock ledger: ISSUE is -qty (from store), so site gets +qty

            if e_date < start_date:
                opening_qty += net_site_change
            elif start_date <= e_date <= end_date:
                period_txns.append((e_date, net_site_change))

        # Build timeline splits
        timeline = {}
        curr_qty = opening_qty

        # Initialize timeline start
        timeline[start_date] = curr_qty

        for t_date, qty_change in period_txns:
            curr_qty += qty_change
            timeline[t_date] = curr_qty

        # Create period segments
        dates = sorted(list(timeline.keys()))
        segments = []

        for idx in range(len(dates)):
            seg_start = dates[idx]
            seg_qty = timeline[seg_start]

            if seg_qty <= 0:
                continue

            if idx < len(dates) - 1:
                seg_end = dates[idx + 1] - timedelta(days=1)
            else:
                seg_end = end_date

            if seg_start <= seg_end:
                days = (seg_end - seg_start).days + 1
                nos = seg_qty * Decimal(days)
                amount = (nos * daily_rate).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

                segments.append({
                    "qty": float(seg_qty),
                    "period_str": f"{seg_start.strftime('%d-%m-%y')} TO {seg_end.strftime('%d-%m-%y')}",
                    "days": days,
                    "nos": float(nos),
                    "rate": float(daily_rate),
                    "amount": float(amount),
                    "amount_dec": amount
                })

        if segments:
            bill_items.append({
                "s_no": item_counter,
                "item_name": item_info["name"],
                "rowspan": len(segments),
                "segments": segments
            })
            item_counter += 1

    # 4. Total Calculation
    taxable_value = sum(
        seg["amount_dec"]
        for b_item in bill_items
        for seg in b_item["segments"]
    )

    cgst = (taxable_value * Decimal("0.09")).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    sgst = (taxable_value * Decimal("0.09")).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    total_amount = taxable_value + cgst + sgst

    net_total_rounded = total_amount.quantize(Decimal("1"), rounding=ROUND_HALF_UP)
    round_off = net_total_rounded - total_amount

    return {
        "party_id": party_id,
        "site_id": site_id,
        "start_date": start_date_str,
        "end_date": end_date_str,
        "items": bill_items,
        "summary": {
            "taxable_value": float(taxable_value),
            "cgst": float(cgst),
            "sgst": float(sgst),
            "total_amount": float(total_amount),
            "round_off": float(round_off),
            "net_total": float(net_total_rounded)
        }
    }