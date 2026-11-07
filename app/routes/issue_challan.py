from flask import Blueprint, render_template, request, redirect, url_for
import os
import pandas as pd

issue_challan_bp = Blueprint('issue_challan', __name__, url_prefix='/issue-challan')

EXCEL_PATH = 'om_enterprise_master.xlsx'

@issue_challan_bp.route('/')
def index():
    challans = []
    total_records = 0
    try:
        if os.path.exists(EXCEL_PATH):
            df = pd.read_excel(EXCEL_PATH, sheet_name='Issue_Challan')
            challans = df.to_dict(orient='records')
            total_records = len(challans)
    except Exception:
        try:
            excel_file = pd.ExcelFile(EXCEL_PATH)
            sheet_name = excel_file.sheet_names[0]
            df = pd.read_excel(EXCEL_PATH, sheet_name=sheet_name)
            challans = df.to_dict(orient='records')
            total_records = len(challans)
        except:
            challans = []

    return render_template(
        'issue_challan/index.html',
        challans=challans,
        page=1,
        total_pages=1,
        total_records=total_records
    )

@issue_challan_bp.route('/generate', methods=['GET', 'POST'])
def generate():
    if request.method == 'POST':
        # Save logic here if needed
        return redirect(url_for('issue_challan.index'))
    
    parties = []
    sites = []
    items = []
    try:
        if os.path.exists(EXCEL_PATH):
            xl = pd.ExcelFile(EXCEL_PATH)
            if 'Parties' in xl.sheet_names:
                parties = pd.read_excel(EXCEL_PATH, sheet_name='Parties').to_dict(orient='records')
            if 'Sites' in xl.sheet_names:
                sites = pd.read_excel(EXCEL_PATH, sheet_name='Sites').to_dict(orient='records')
            if 'Items' in xl.sheet_names:
                items = pd.read_excel(EXCEL_PATH, sheet_name='Items').to_dict(orient='records')
    except Exception as e:
        print("Error loading dropdown data:", e)

    # Render generate/create form template with dropdown lists
    try:
        return render_template('issue_challan/generate.html', parties=parties, sites=sites, items=items)
    except:
        return render_template('issue_challan/index.html', parties=parties, sites=sites, items=items)

@issue_challan_bp.route('/create', methods=['POST'])
def create():
    return redirect(url_for('issue_challan.index'))
