# from flask import Blueprint, render_template
# from app.services.excel_service import get_sheet_rows

# dashboard_bp = Blueprint("dashboard", __name__)


# @dashboard_bp.route("/")
# def dashboard():
#     try:
#         parties = get_sheet_rows("Parties")
#         sites = get_sheet_rows("Sites")
#         items = get_sheet_rows("Items")
#         challans = get_sheet_rows("Challans")
#     except Exception:
#         parties, sites, items, challans = [], [], [], []

#     active_challans = [
#         c for c in challans if str(c.get("Status", "")).strip().lower() == "active"
#     ]

#     return render_template(
#         "dashboard/index.html",
#         total_parties=len(parties),
#         total_sites=len(sites),
#         total_items=len(items),
#         total_active_challans=len(active_challans),
#         recent_challans=challans[-5:] if challans else []  # Last 5 transactions
#     )

from flask import Blueprint, render_template
from app.routes.challan_store import Book, load_parties, load_items_master, records, _issue_sheet, is_active

dashboard_bp = Blueprint("dashboard", __name__)

@dashboard_bp.route("/")
def dashboard():
    total_parties = 0
    total_sites = 0
    total_items = 0
    total_active_challans = 0
    recent_challans = []

    try:
        book = Book.load()
        
        # 1. Parties Count
        parties = load_parties(book)
        total_parties = len([p for p in parties if sid(p.get('Party_ID')) or p.get('Party_Name')])
        
        # 2. Items Count
        items = load_items_master(book)
        total_items = len([i for i in items if is_active(i)])
        
        # 3. Sites Count (Checking 'SiteLocations' first, followed by fallbacks)
        total_sites = 0
        for sheet_name in ['SiteLocations', 'Sites', 'Site_Master', 'Party_Sites']:
            try:
                sites_df = book.df(sheet_name)
                if not sites_df.empty:
                    total_sites = len(records(sites_df))
                    break
            except Exception:
                continue
        
        if total_sites == 0:
            try:
                challan_df = book.df(_issue_sheet(book))
                challan_records = records(challan_df)
                sites_set = {str(c.get('Site_Name') or c.get('Site_Location') or c.get('Site')).strip() for c in challan_records if c.get('Site_Name') or c.get('Site_Location') or c.get('Site')}
                total_sites = len([s for s in sites_set if s and s.lower() != 'nan' and s != ''])
            except Exception:
                total_sites = 0

        # 4. Challans Count & Recent Transactions
        try:
            challan_df = book.df(_issue_sheet(book))
            challan_records = records(challan_df)
            active_challans = [c for c in challan_records if not str(c.get("Status", "")).strip().lower() in ('cancelled', 'cancel')]
            total_active_challans = len(active_challans)
            recent_challans = challan_records[-5:] if challan_records else []
        except Exception:
            pass

    except Exception as e:
        print("Dashboard Load Error:", e)

    return render_template(
        "dashboard/index.html",
        total_parties=total_parties,
        total_sites=total_sites,
        total_items=total_items,
        total_active_challans=total_active_challans,
        recent_challans=recent_challans
    )

def sid(val):
    if val is None:
        return ''
    val_str = str(val).strip()
    if val_str.endswith('.0'):
        val_str = val_str[:-2]
    return val_str