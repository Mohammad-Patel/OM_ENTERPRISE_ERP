from flask import Blueprint, render_template, request, redirect, url_for, flash
from app.services.challan_service import create_receive_challan, ChallanServiceError
from app.services.excel_service import get_sheet_rows

receive_challan_bp = Blueprint("receive_challan", __name__, url_prefix="/receive-challan")


@receive_challan_bp.route("/", methods=["GET"])
def index():
    """List all receive challans."""
    try:
        challans = get_sheet_rows("ReceiveChallans")
    except Exception:
        challans = []
    return render_template("receive_challan/index.html", challans=challans)


@receive_challan_bp.route("/new", methods=["GET", "POST"])
def create():
    """Create Receive Challan with usable & breakage entries."""
    if request.method == "POST":
        try:
            data = {
                "challan_no": request.form.get("challan_no"),
                "date": request.form.get("date"),
                "party_id": request.form.get("party_id"),
                "site_id": request.form.get("site_id"),
                "vehicle_no": request.form.get("vehicle_no"),
                "driver_name": request.form.get("driver_name"),
                "driver_mob": request.form.get("driver_mob"),
                "weight": request.form.get("weight"),
            }

            # 1. Usable Return Items
            item_ids = request.form.getlist("item_id[]")
            quantities = request.form.getlist("quantity[]")
            items_list = []
            for i_id, qty in zip(item_ids, quantities):
                if i_id and qty and float(qty) > 0:
                    items_list.append({"item_id": i_id, "quantity": qty})

            # 2. Breakage / Damage Items
            b_item_ids = request.form.getlist("b_item_id[]")
            b_quantities = request.form.getlist("b_quantity[]")
            b_rates = request.form.getlist("b_rate[]")
            b_descriptions = request.form.getlist("b_description[]")
            
            breakage_list = []
            for b_id, b_qty, b_rate, b_desc in zip(b_item_ids, b_quantities, b_rates, b_descriptions):
                if b_id and b_qty and float(b_qty) > 0:
                    breakage_list.append({
                        "item_id": b_id,
                        "quantity": b_qty,
                        "rate": b_rate or "0",
                        "description": b_desc or ""
                    })

            res = create_receive_challan(data, items_list, breakage_list)
            flash(f"Receive Challan {res['challan_no']} successfully recorded!", "success")
            return redirect(url_for("receive_challan.print_challan", challan_no=res["challan_no"]))

        except ChallanServiceError as e:
            flash(str(e), "danger")
        except Exception as e:
            flash(f"Error processing receive challan: {str(e)}", "danger")

    parties = get_sheet_rows("Parties")
    sites = get_sheet_rows("Sites")
    items = get_sheet_rows("Items")

    return render_template(
        "receive_challan/form.html",
        parties=parties,
        sites=sites,
        items=items,
    )


@receive_challan_bp.route("/print/<challan_no>", methods=["GET"])
def print_challan(challan_no):
    """Render printable Receive Challan layout with Breakage details."""
    try:
        all_challans = get_sheet_rows("ReceiveChallans")
        challan = next((c for c in all_challans if str(c.get("Challan_No")) == str(challan_no)), None)
        
        if not challan:
            flash(f"Receive Challan No. {challan_no} not found.", "danger")
            return redirect(url_for("receive_challan.index"))

        parties = get_sheet_rows("Parties")
        sites = get_sheet_rows("Sites")
        party = next((p for p in parties if str(p.get("Party_ID")) == str(challan.get("Party_ID"))), {})
        site = next((s for s in sites if str(s.get("Site_ID")) == str(challan.get("Site_ID"))), {})

        ledger = get_sheet_rows("StockLedger")
        items_rows = get_sheet_rows("Items")
        items_dict = {str(i.get("Item_ID")): i.get("Item_Name") for i in items_rows}

        usable_items = []
        breakage_items = []
        total_usable_qty = 0
        total_breakage_qty = 0
        total_breakage_amt = 0.0

        u_no = 1
        for entry in ledger:
            if str(entry.get("Challan_No")) == str(challan_no):
                i_id = str(entry.get("Item_ID"))
                i_name = items_dict.get(i_id, f"Item {i_id}")

                if entry.get("Transaction_Type") == "RECEIVE":
                    qty = float(entry.get("Usable_Qty_Change", 0))
                    usable_items.append({"s_no": u_no, "item_name": i_name, "quantity": qty})
                    total_usable_qty += qty
                    u_no += 1

                elif entry.get("Transaction_Type") == "BREAKAGE":
                    b_qty = float(entry.get("Damaged_Qty_Change", 0))
                    b_rate = float(entry.get("Breakage_Rate", 0))
                    b_amt = b_qty * b_rate
                    breakage_items.append({
                        "item_name": i_name,
                        "quantity": b_qty,
                        "rate": b_rate,
                        "amount": b_amt,
                        "description": entry.get("Description", "")
                    })
                    total_breakage_qty += b_qty
                    total_breakage_amt += b_amt

        return render_template(
            "print/receive_challan.html",
            challan=challan,
            party=party,
            site=site,
            items=usable_items,
            breakage_items=breakage_items,
            total_qty=total_usable_qty,
            total_breakage_qty=total_breakage_qty,
            total_breakage_amt=total_breakage_amt
        )
    except Exception as e:
        flash(f"Error displaying print view: {str(e)}", "danger")
        return redirect(url_for("receive_challan.index"))