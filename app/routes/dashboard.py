from flask import Blueprint, render_template
from app.services.excel_service import get_sheet_rows

dashboard_bp = Blueprint("dashboard", __name__)


@dashboard_bp.route("/")
def dashboard():
    try:
        parties = get_sheet_rows("Parties")
        sites = get_sheet_rows("Sites")
        items = get_sheet_rows("Items")
    except Exception:
        parties, sites, items = [], [], []

    return render_template(
        "dashboard/index.html",
        total_parties=len(parties),
        total_sites=len(sites),
        total_items=len(items)
    )