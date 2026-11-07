from flask import Blueprint, render_template, request
from app.services.stock_service import get_godown_stock_summary, list_stock_ledger
from app.services.excel_service import get_sheet_rows

stock_bp = Blueprint("stock", __name__, url_prefix="/stock")


@stock_bp.route("/godown", methods=["GET"])
def godown_stock():
    """View Current Real-time Available Stock in Godown."""
    godown_summary = get_godown_stock_summary()
    return render_template("stock/godown.html", stock_items=godown_summary)


@stock_bp.route("/party-site", methods=["GET"])
def party_site_stock():
    """View Party & Site-wise Current Rented Outstanding Stock."""
    selected_party_id = request.args.get("party_id", "")
    selected_site_id = request.args.get("site_id", "")

    parties = get_sheet_rows("Parties")
    sites = get_sheet_rows("Sites")
    items_rows = get_sheet_rows("Items")
    items_dict = {str(i.get("Item_ID")): i.get("Item_Name") for i in items_rows}

    outstanding_report = []

    if selected_party_id and selected_site_id:
        ledger = list_stock_ledger(party_id=selected_party_id, site_id=selected_site_id)
        
        # Calculate item-wise outstanding balance on site
        item_balances = {}
        for entry in ledger:
            i_id = str(entry.get("Item_ID"))
            chg = float(entry.get("Usable_Qty_Change", 0))
            # ISSUE is negative in store ledger, meaning positive at site
            site_change = -chg
            item_balances[i_id] = item_balances.get(i_id, 0.0) + site_change

        for i_id, balance in item_balances.items():
            if balance > 0:
                outstanding_report.append({
                    "item_id": i_id,
                    "item_name": items_dict.get(i_id, f"Item {i_id}"),
                    "outstanding_qty": balance
                })

    return render_template(
        "stock/party_site.html",
        parties=parties,
        sites=sites,
        selected_party_id=selected_party_id,
        selected_site_id=selected_site_id,
        report=outstanding_report
    )