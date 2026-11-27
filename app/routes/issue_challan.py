"""
Issue Challan module - OM Enterprise ERP
----------------------------------------
Excel (data/om_enterprise.xlsx) based storage.

Sheets used:
    Issue_Challan, Issue_Challan_Items   -> challan header / item lines
    Customers, SiteLocations, Items      -> master data
    Audit_Log                            -> create / update / delete history

Production notes:
    * Every read-modify-write runs inside a lock (waitress is multi-threaded).
    * All sheets of one action are saved together, through a temp file that
      replaces the real file only on success (a crash can never half-write it).
    * Challan numbers are always generated on the server (never trusted from form).
    * Inputs are validated (party/site relation, items, quantities, date).
"""
import json
import logging
import math
import os
import re
import shutil
import tempfile
import threading
from datetime import datetime
from functools import wraps

import pandas as pd
from flask import (Blueprint, flash, redirect, render_template, request,
                   session, url_for)

logger = logging.getLogger(__name__)

issue_challan_bp = Blueprint('issue_challan', __name__, url_prefix='/issue-challan')

EXCEL_PATH = os.path.join('data', 'om_enterprise.xlsx')

CHALLAN_SHEET = 'Issue_Challan'
CHALLAN_ITEMS_SHEET = 'Issue_Challan_Items'
AUDIT_SHEET = 'Audit_Log'

CHALLAN_PREFIX = 'OM/IC/'
DATE_FMT = '%d-%m-%Y'
DEFAULT_GOODS_RATE = 59
DEFAULT_UNIT_WEIGHT = 10
ALLOWED_PER_PAGE = ('10', '25', '50', 'All')

_excel_lock = threading.RLock()


class ChallanError(Exception):
    """User-facing validation error (message is shown with flash)."""


# GST rule + numbering live in ONE place (app/services/numbering_service.py)
from app.services.numbering_service import (is_gst_registered, next_issue_challan_no, take_challan_no,
                                            preview_challan_numbers, ISSUE,
                                            is_valid_gstin)

# --------------------------------------------------------------------------
# Small helpers
# --------------------------------------------------------------------------
def sid(v):
    """ID -> comparable string ('1.0' -> '1')."""
    s = '' if v is None else str(v).strip()
    return re.sub(r'\.0+$', '', s)


def norm_id(v):
    """Digit IDs are stored as int, everything else stays text."""
    s = sid(v)
    return int(s) if s.isdigit() else s


def to_float(value, default=0.0):
    try:
        if value is None or value == '':
            return default
        return float(value)
    except (ValueError, TypeError):
        return default


def to_int(value, default=1):
    try:
        return int(value)
    except (ValueError, TypeError):
        return default


def to_js_json(data):
    """JSON that is safe to place inside an HTML attribute / script."""
    return json.dumps(data, default=str).replace('</', '<\\/')


def current_user():
    return str(session.get('username') or session.get('user') or session.get('user_name') or '')


def _clean_value(v):
    if isinstance(v, float) and v.is_integer():
        return int(v)
    if isinstance(v, pd.Timestamp):
        return v.strftime(DATE_FMT)
    if hasattr(v, 'item') and not isinstance(v, (str, bytes)):
        try:
            return v.item()
        except Exception:
            return v
    return v


def records(df):
    """DataFrame -> list of dicts (NaN -> '', dates -> dd-mm-yyyy, 5.0 -> 5)."""
    if df is None or df.empty:
        return []
    df = df.copy()
    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = df[col].dt.strftime(DATE_FMT)
    df = df.astype(object).where(pd.notna(df), '')
    return [{k: _clean_value(v) for k, v in row.items()}
            for row in df.to_dict(orient='records')]


def next_id(df, col):
    if df is None or df.empty or col not in df.columns:
        return 1
    nums = pd.to_numeric(df[col], errors='coerce').dropna()
    return int(nums.max()) + 1 if not nums.empty else 1


def append_rows(df, rows):
    """Append rows while keeping the existing column order."""
    cols = list(df.columns) if df is not None else []
    for r in rows:
        for k in r:
            if k not in cols:
                cols.append(k)
    new_df = pd.DataFrame(rows, columns=cols)
    if df is None or df.empty:
        return new_df
    return pd.concat([df, new_df], ignore_index=True)[cols]


def is_active(row):
    return str(row.get('Is_Active', '')).strip().lower() not in ('false', 'no', 'n', '0', 'inactive')


# --------------------------------------------------------------------------
# Display formatters (used by print page)
# --------------------------------------------------------------------------
_ONES = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine', 'Ten',
         'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen',
         'Eighteen', 'Nineteen']
_TENS = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety']


def _below_hundred(n):
    if n < 20:
        return _ONES[n]
    return _TENS[n // 10] + (' ' + _ONES[n % 10] if n % 10 else '')


def _below_thousand(n):
    words = []
    if n >= 100:
        words.append(_ONES[n // 100] + ' Hundred')
        n %= 100
    if n:
        words.append(_below_hundred(n))
    return ' '.join(words)


def number_to_words(num):
    """350 -> 'Three Hundred Fifty Only' (Indian system: Thousand / Lakh / Crore)."""
    n = int(round(to_float(num)))
    if n <= 0:
        return 'Zero Only'
    parts = []
    for div, name in ((10000000, 'Crore'), (100000, 'Lakh'), (1000, 'Thousand')):
        if n >= div:
            parts.append(_below_thousand(n // div) + ' ' + name)
            n %= div
    if n:
        parts.append(_below_thousand(n))
    return ' '.join(parts) + ' Only'


def fmt_qty(v):
    """50.0 -> '50', 7300.5 -> '7300.5'"""
    f = to_float(v)
    if f == int(f):
        return str(int(f))
    return f"{f:.2f}".rstrip('0').rstrip('.')


def fmt_money(v):
    """430700 -> '430,700.00'"""
    return f"{to_float(v):,.2f}"


# --------------------------------------------------------------------------
# Excel access
# --------------------------------------------------------------------------
class Book:
    """Snapshot of all sheets (file is closed immediately after reading)."""

    def __init__(self, sheets):
        self.sheets = sheets
        self._lower = {n.lower(): n for n in sheets}

    @classmethod
    def load(cls):
        if not os.path.exists(EXCEL_PATH):
            return cls({})
        return cls(pd.read_excel(EXCEL_PATH, sheet_name=None))

    def find(self, exact=(), keywords=(), exclude=()):
        """Exact names are checked in priority order, then keywords."""
        for ex in exact:
            if ex in self._lower:
                return self._lower[ex]
        for name in self.sheets:
            low = name.lower()
            if any(x in low for x in exclude):
                continue
            if any(k in low for k in keywords):
                return name
        return None

    def df(self, name):
        if name and name in self.sheets:
            return self.sheets[name].copy()
        return pd.DataFrame()


def challan_sheet(book):
    return book.find(exact=('issue_challan',), keywords=('issue_challan',), exclude=('item',))


def challan_items_sheet(book):
    return book.find(exact=('issue_challan_items',))


def audit_sheet(book):
    return book.find(exact=('audit_log',))


def load_parties(book):
    s = book.find(exact=('customers', 'customer', 'parties', 'party'),
                  keywords=('customer', 'client'),
                  exclude=('challan', 'issue', 'site', 'item'))
    return records(book.df(s))


def load_sites(book):
    s = book.find(exact=('sitelocations', 'sitelocation', 'sites', 'site'),
                  keywords=('sitelocation',),
                  exclude=('challan', 'issue', 'item'))
    return records(book.df(s))


def load_items_master(book):
    s = book.find(exact=('items', 'item'),
                  keywords=('product',),
                  exclude=('challan', 'issue'))
    return records(book.df(s))


def save_sheets(updates):
    """
    Save several sheets in ONE operation.
    Written to a temp copy first and swapped in only on success, so the real
    workbook is never left half-written.
    """
    directory = os.path.dirname(os.path.abspath(EXCEL_PATH))
    os.makedirs(directory, exist_ok=True)
    fd, tmp_path = tempfile.mkstemp(suffix='.xlsx', dir=directory)
    os.close(fd)
    try:
        if os.path.exists(EXCEL_PATH):
            shutil.copy2(EXCEL_PATH, tmp_path)
            writer_kwargs = {'mode': 'a', 'if_sheet_exists': 'replace'}
        else:
            writer_kwargs = {'mode': 'w'}
        with pd.ExcelWriter(tmp_path, engine='openpyxl', **writer_kwargs) as writer:
            for sheet, df in updates.items():
                df.to_excel(writer, sheet_name=sheet, index=False)
        os.replace(tmp_path, EXCEL_PATH)
    finally:
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass


def audit_update(book, action, entity_id, before=None, after=None, reason=''):
    """Returns (sheet_name, dataframe) with one audit row appended."""
    name = audit_sheet(book) or AUDIT_SHEET
    df = book.df(name)
    row = {
        'Audit_ID': next_id(df, 'Audit_ID'),
        'Event_Time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'Username': current_user(),
        'Action': action,
        'Entity_Type': 'Issue_Challan',
        'Entity_ID': entity_id,
        'Reason': reason,
        'Before_Data': json.dumps(before, default=str, ensure_ascii=False) if before else '',
        'After_Data': json.dumps(after, default=str, ensure_ascii=False) if after else '',
    }
    return name, append_rows(df, [row])


# --------------------------------------------------------------------------
# Business helpers
# --------------------------------------------------------------------------
def next_challan_no(book, party):
    """GST party -> OM/26-27/180 ...   Non-GST party -> 01, 02 ... (rules: numbering_service)."""
    return next_issue_challan_no(book, party)


def preview_numbers(book, parties):
    """Next number for each series + which parties are GST (used by the create popup)."""
    nums = preview_challan_numbers(book, ISSUE)
    return {
        'gst': nums['gst'],
        'non_gst': nums['non_gst'],
        'gst_party_ids': [sid(p.get('Party_ID')) for p in parties if is_gst_registered(p)],
    }


def pick_challan_row(df, challan_no):
    """
    Index of the challan row for a number.  The same number can legitimately exist twice (a series
    can be restarted from Settings), so the list passes ?id=<Issue_ID>; without it the NEWEST
    challan with that number is used.
    """
    mask = df['Challan_No'].astype(str) == str(challan_no)
    wanted = sid(request.values.get('id', ''))
    if wanted and 'Issue_ID' in df.columns:
        by_id = mask & (df['Issue_ID'].apply(sid) == wanted)
        if by_id.any():
            return df.index[by_id][-1]
    if not mask.any():
        return None
    return df.index[mask][-1]


def normalize_date(value):
    value = (value or '').strip()
    if not value:
        return datetime.now().strftime(DATE_FMT)
    for fmt in ('%d-%m-%Y', '%Y-%m-%d', '%d/%m/%Y'):
        try:
            return datetime.strptime(value, fmt).strftime(DATE_FMT)
        except ValueError:
            continue
    raise ChallanError('Invalid date. Please use DD-MM-YYYY.')


def parse_any_date(value):
    """'05-09-2026' / '2026-09-05' / '05/09/2026' -> date (or None)."""
    value = str(value or '').strip()
    if not value:
        return None
    for fmt in ('%d-%m-%Y', '%Y-%m-%d', '%d/%m/%Y'):
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            continue
    return None


def parse_header(form, parties, sites):
    """Validate and clean the header part of the form."""
    party_id = sid(form.get('party_id'))
    site_id = sid(form.get('site_id'))

    party = parties.get(party_id)
    if not party:
        raise ChallanError('Please select a valid party.')
    site = sites.get(site_id)
    if not site:
        raise ChallanError('Please select a valid site.')
    if sid(site.get('Party_ID')) != party_id:
        raise ChallanError('Selected site does not belong to the selected party.')

    return {
        'Party_ID': norm_id(party_id),
        'Site_ID': norm_id(site_id),
        'Party_Type': party.get('Party_Type', ''),
        'Challan_Date': normalize_date(form.get('date')),
        'Vehicle_No': (form.get('vehicle_no') or '').strip().upper(),
        'Driver_Name': (form.get('driver_name') or '').strip(),
        'Driver_Mob': (form.get('driver_mob') or '').strip(),
        'Driver_Licence': (form.get('driver_licence') or '').strip().upper(),
        'EWay_Bill_No': (form.get('eway_bill_no') or '').strip(),
    }


def build_lines(form, item_master):
    """Merge duplicate items, validate, and calculate weight / value."""
    merged = {}
    for iid, qty in zip(form.getlist('item_id[]'), form.getlist('quantity[]')):
        iid = sid(iid)
        if not iid:
            continue
        q = to_float(qty)
        if q <= 0:
            continue
        merged[iid] = merged.get(iid, 0) + q

    if not merged:
        raise ChallanError('Please add at least one item with quantity.')

    lines, total_wt, total_val = [], 0.0, 0.0
    for iid, q in merged.items():
        m = item_master.get(iid)
        if not m:
            raise ChallanError(f'Invalid item selected (Item ID {iid}).')
        unit_wt = to_float(m.get('Unit_Weight_KG'), DEFAULT_UNIT_WEIGHT) or DEFAULT_UNIT_WEIGHT
        rate = to_float(m.get('Goods_Value_Rate'), DEFAULT_GOODS_RATE) or DEFAULT_GOODS_RATE
        line_wt = unit_wt * q
        line_val = line_wt * rate
        total_wt += line_wt
        total_val += line_val
        lines.append({
            'Item_ID': norm_id(iid),
            'Quantity': q,
            'Unit_Weight_KG': unit_wt,
            'Total_Weight_KG': round(line_wt, 2),
            'Goods_Value_Rate': rate,
            'Approx_Goods_Value': round(line_val, 2),
        })
    return lines, round(total_wt, 2), round(total_val, 2)


def number_lines(lines, issue_id, start_line_id):
    """Attach Line_ID / Issue_ID in front of every line."""
    out, line_id = [], start_line_id
    for ln in lines:
        row = {'Line_ID': line_id, 'Issue_ID': issue_id}
        row.update(ln)
        out.append(row)
        line_id += 1
    return out


def safe_action(error_message):
    """Common error handling for create / edit / delete."""
    def decorator(fn):
        @wraps(fn)
        def wrapper(*args, **kwargs):
            try:
                return fn(*args, **kwargs)
            except ChallanError as e:
                flash(str(e), 'danger')
            except PermissionError:
                flash('Excel file is open in another program. Please close it and try again.', 'danger')
            except Exception:
                logger.exception(error_message)
                flash(error_message, 'danger')
            return redirect(url_for('issue_challan.index'))
        return wrapper
    return decorator


# --------------------------------------------------------------------------
# Routes
# --------------------------------------------------------------------------
@issue_challan_bp.route('/')
def index():
    search_query = request.args.get('search', '').strip()
    party_filter = request.args.get('party_id', '').strip()
    gst_filter = request.args.get('gst_filter', '').strip()  # <-- NEW: 'gst' or 'nogst'
    per_page = request.args.get('per_page', '10')
    if per_page not in ALLOWED_PER_PAGE:
        per_page = '10'
    page = max(to_int(request.args.get('page'), 1), 1)
    site_filter = request.args.get('site_id', '').strip()
    date_from = parse_any_date(request.args.get('date_from'))
    date_to = parse_any_date(request.args.get('date_to'))

    all_challans, parties, sites, items = [], [], [], []
    items_by_challan = {}

    try:
        book = Book.load()
        all_parties = load_parties(book)
        all_sites = load_sites(book)
        
        # --- NEW: GST / Without-GST Party Filtering ---
        filtered_parties = [p for p in all_parties if is_active(p)]
        if gst_filter == 'gst':
            filtered_parties = [p for p in filtered_parties if is_gst_registered(p)]
        elif gst_filter == 'nogst':
            filtered_parties = [p for p in filtered_parties if not is_gst_registered(p)]
        parties = filtered_parties
        # ---------------------------------------------

        sites = [s for s in all_sites if is_active(s)]

        for itm in load_items_master(book):
            if not is_active(itm):
                continue
            items.append({
                'Item_ID': itm.get('Item_ID', ''),
                'Item_Name': itm.get('Item_Name') or itm.get('Item_Code') or '',
                'Size': itm.get('Size') or '',
                'Unit_Weight_KG': to_float(itm.get('Unit_Weight_KG'), DEFAULT_UNIT_WEIGHT) or DEFAULT_UNIT_WEIGHT,
                'Goods_Value_Rate': to_float(itm.get('Goods_Value_Rate'), DEFAULT_GOODS_RATE) or DEFAULT_GOODS_RATE,
            })

        party_names = {sid(p.get('Party_ID')): p.get('Party_Name') or p.get('Party_ID') for p in all_parties}
        
        # Helper set for filtering challans based on selected GST type
        allowed_party_ids = {sid(p.get('Party_ID')) for p in parties}

        site_names = {sid(s.get('Site_ID')): s.get('Site_Name') or s.get('Site_Address') or s.get('Site_ID')
                      for s in all_sites}
        party_phones = {sid(p.get('Party_ID')): p.get('Phone') or p.get('Mobile_No') or p.get('Mobile') or ''
                        for p in all_parties}

        all_challans = records(book.df(challan_sheet(book)))
        for c in all_challans:
            c['Party_Name'] = party_names.get(sid(c.get('Party_ID')), c.get('Party_ID', ''))
            c['Site_Name'] = site_names.get(sid(c.get('Site_ID')), c.get('Site_ID', ''))
            c['Party_Mobile'] = party_phones.get(sid(c.get('Party_ID'))) or '-'

        # item lines per challan (for the edit popup)
        known_ids = {sid(c.get('Issue_ID')) for c in all_challans}
        for r in records(book.df(challan_items_sheet(book))):
            no = sid(r.get('Issue_ID'))                  # keyed by Issue_ID: the number can repeat
            if no in known_ids:
                items_by_challan.setdefault(no, []).append({
                    'item_id': sid(r.get('Item_ID')),
                    'qty': r.get('Quantity'),
                })
    except Exception:
        logger.exception('Error loading issue challan list')
        flash('Could not read the Excel data file. Please check that it exists and is not corrupted.', 'danger')

    try:
        numbers = preview_numbers(book, all_parties)
    except Exception:
        logger.exception('Could not compute next challan numbers')
        numbers = {'gst': '', 'non_gst': '', 'gst_party_ids': []}

    # newest first
    all_challans = list(reversed(all_challans))

    # filters
    filtered = all_challans

    # --- NEW: Filter challans by GST Type ---
    if gst_filter:
        filtered = [c for c in filtered if sid(c.get('Party_ID')) in allowed_party_ids]
    # ----------------------------------------

    if party_filter:
        filtered = [c for c in filtered if sid(c.get('Party_ID')) == sid(party_filter)]
    if site_filter:
        filtered = [c for c in filtered if sid(c.get('Site_ID')) == sid(site_filter)]
    if date_from or date_to:
        def in_range(c):
            d = parse_any_date(c.get('Challan_Date'))
            if d is None:
                return False
            if date_from and d < date_from:
                return False
            if date_to and d > date_to:
                return False
            return True
        filtered = [c for c in filtered if in_range(c)]
    if search_query:
        q = search_query.lower()
        fields = ('Challan_No', 'Party_Name', 'Site_Name', 'Party_Mobile', 'Vehicle_No', 'Driver_Name')
        filtered = [c for c in filtered if any(q in str(c.get(f, '')).lower() for f in fields)]

    total_records = len(filtered)

    # pagination
    if per_page == 'All':
        total_pages, page = 1, 1
        page_rows = filtered
    else:
        size = int(per_page)
        total_pages = max(math.ceil(total_records / size), 1)
        page = min(page, total_pages)
        page_rows = filtered[(page - 1) * size: page * size]

    date_from_iso = date_from.isoformat() if date_from else ''
    date_to_iso = date_to.isoformat() if date_to else ''
    qargs = {k: v for k, v in {
        'search': search_query, 'party_id': party_filter, 'site_id': site_filter,
        'gst_filter': gst_filter, 'date_from': date_from_iso, 'date_to': date_to_iso, 'per_page': per_page,
    }.items() if v}

    return render_template(
        'issue_challan/index.html',
        challans=page_rows,
        parties=parties,
        sites_json=to_js_json(sites),
        items_json=to_js_json(items),
        challans_json=to_js_json(all_challans),
        challan_items_json=to_js_json(items_by_challan),
        search_query=search_query,
        party_filter=party_filter,
        site_filter=site_filter,
        gst_filter=gst_filter,
        date_from=date_from_iso,
        date_to=date_to_iso,
        qargs=qargs,
        filters_active=bool(search_query or party_filter or site_filter or gst_filter or date_from or date_to),
        per_page=per_page,
        page=page,
        total_pages=total_pages,
        total_records=total_records,
        current_date=datetime.now().strftime(DATE_FMT),
        next_gst_challan_no=numbers['gst'],
        next_non_gst_challan_no=numbers['non_gst'],
        gst_party_ids_json=json.dumps(numbers['gst_party_ids']),
    )


@safe_action('Error creating challan.')
def _create_challan():
    f = request.form
    with _excel_lock:
        book = Book.load()
        if not book.sheets:
            raise ChallanError('Excel data file not found.')

        c_sheet = challan_sheet(book) or CHALLAN_SHEET
        i_sheet = challan_items_sheet(book) or CHALLAN_ITEMS_SHEET
        df = book.df(c_sheet)
        items_df = book.df(i_sheet)

        parties = {sid(p.get('Party_ID')): p for p in load_parties(book)}
        sites = {sid(s.get('Site_ID')): s for s in load_sites(book)}
        item_master = {sid(i.get('Item_ID')): i for i in load_items_master(book)}

        header = parse_header(f, parties, sites)
        lines, total_wt, total_val = build_lines(f, item_master)

        # always server generated, from the party's GST status (never trusted from the form)
        challan_no, counter_update = take_challan_no(book, ISSUE, parties.get(sid(header['Party_ID']), {}),
                                                     records(df))
        issue_id = next_id(df, 'Issue_ID')
        lines = number_lines(lines, issue_id, next_id(items_df, 'Line_ID'))

        new_row = {
            'Issue_ID': issue_id,
            'Challan_No': challan_no,
            'Party_ID': header['Party_ID'],
            'Site_ID': header['Site_ID'],
            'Challan_Date': header['Challan_Date'],
            'Party_Type': header['Party_Type'],
            'Transporter_Name': '',
            'Vehicle_No': header['Vehicle_No'],
            'EWay_Bill_No': header['EWay_Bill_No'],
            'Remarks': '',
            'Status': 'Active',
            'Created_At': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'Created_By': current_user(),
            'Cancel_Reason': '',
            'Driver_Name': header['Driver_Name'],
            'Driver_Mob': header['Driver_Mob'],
            'Weight': total_wt,
            'Goods_Value': total_val,
            'Driver_Licence': header['Driver_Licence'],
        }

        a_sheet, audit_df = audit_update(book, 'CREATE', challan_no,
                                         after={**new_row, 'Items': lines})
        save_sheets({
            c_sheet: append_rows(df, [new_row]),
            i_sheet: append_rows(items_df, lines),
            a_sheet: audit_df,
            **counter_update,                      # series counter + 1, saved together with the challan
        })

    flash(f'Challan {challan_no} created successfully!', 'success')
    return redirect(url_for('issue_challan.print_challan', challan_no=challan_no, id=issue_id))


@issue_challan_bp.route('/generate', methods=['GET', 'POST'])
def generate():
    # The create popup lives on the main page, so GET simply goes there.
    if request.method == 'POST':
        return _create_challan()
    return redirect(url_for('issue_challan.index'))


@issue_challan_bp.route('/create', methods=['POST'])
def create():
    return _create_challan()


@issue_challan_bp.route('/print/<path:challan_no>')
def print_challan(challan_no):
    challan_data = None
    challan_items = []
    party, site = {}, {}
    total_qty = 0.0

    try:
        book = Book.load()
        df = book.df(challan_sheet(book))
        if not df.empty and 'Challan_No' in df.columns:
            hit = pick_challan_row(df, challan_no)
            if hit is not None:
                challan_data = records(df.loc[[hit]])[0]

        if challan_data:
            party = next((p for p in load_parties(book)
                          if sid(p.get('Party_ID')) == sid(challan_data.get('Party_ID'))), {})
            site = next((s for s in load_sites(book)
                         if sid(s.get('Site_ID')) == sid(challan_data.get('Site_ID'))), {})

            idf = book.df(book.find(exact=('issue_challan_items',), keywords=('challan_item', 'issue_item')))
            rows = pd.DataFrame()
            issue_id = sid(challan_data.get('Issue_ID'))
            if not idf.empty and 'Issue_ID' in idf.columns and issue_id:
                rows = idf[idf['Issue_ID'].apply(sid) == issue_id]
            if rows.empty and not idf.empty and 'Challan_No' in idf.columns:      # older data
                rows = idf[idf['Challan_No'].astype(str) == str(challan_no)]

            master = {sid(i.get('Item_ID')): i for i in load_items_master(book)}
            for r in records(rows):
                m = master.get(sid(r.get('Item_ID')), {})
                name = str(m.get('Item_Name') or m.get('Item_Code')
                           or r.get('Item_Name') or r.get('Item_ID') or '').strip()
                size = str(m.get('Size') or '').strip()
                r['Description'] = (f"{name} {size}".strip()
                                    if size and size.lower() not in name.lower() else name)
                qty = to_float(r.get('Quantity'))
                total_qty += qty
                r['Qty_Display'] = fmt_qty(qty)
                challan_items.append(r)
    except Exception:
        logger.exception('Error loading challan print data')

    weight = challan_data.get('Weight') if challan_data else 0
    goods = challan_data.get('Goods_Value') if challan_data else 0

    html = render_template(
        'issue_challan/print.html',
        challan=challan_data,
        party=party,
        site=site,
        items=challan_items,
        challan_no=challan_no,
        total_qty=fmt_qty(total_qty),
        total_qty_words=number_to_words(total_qty),
        weight_display=f"{fmt_qty(weight)} Kg",
        goods_value_display=fmt_money(goods),
    )
    return html, (200 if challan_data else 404)


@issue_challan_bp.route('/edit/<path:challan_no>', methods=['POST'])
@safe_action('Error updating challan.')
def edit(challan_no):
    f = request.form
    with _excel_lock:
        book = Book.load()
        c_sheet = challan_sheet(book)
        df = book.df(c_sheet)
        if df.empty or 'Challan_No' not in df.columns:
            raise ChallanError('Challan not found.')
        idx = pick_challan_row(df, challan_no)
        if idx is None:
            raise ChallanError('Challan not found.')

        before = records(df.loc[[idx]])[0]
        if str(before.get('Status', '')).strip().lower() in ('cancelled', 'canceled'):
            raise ChallanError('A cancelled challan cannot be edited.')

        parties = {sid(p.get('Party_ID')): p for p in load_parties(book)}
        sites = {sid(s.get('Site_ID')): s for s in load_sites(book)}
        item_master = {sid(i.get('Item_ID')): i for i in load_items_master(book)}

        header = parse_header(f, parties, sites)
        lines, total_wt, total_val = build_lines(f, item_master)
        issue_id = norm_id(before.get('Issue_ID'))

        # make sure every column exists and can hold any value
        for col in ('Challan_Date', 'Party_ID', 'Site_ID', 'Party_Type', 'Vehicle_No', 'Driver_Name',
                    'Driver_Mob', 'Driver_Licence', 'EWay_Bill_No', 'Weight', 'Goods_Value'):
            if col not in df.columns:
                df[col] = ''
            df[col] = df[col].astype(object)

        for col, val in header.items():
            df.at[idx, col] = val
        df.at[idx, 'Weight'] = total_wt
        df.at[idx, 'Goods_Value'] = total_val

        # replace item lines (Line_IDs keep increasing, never reused)
        i_sheet = challan_items_sheet(book) or CHALLAN_ITEMS_SHEET
        items_df = book.df(i_sheet)
        new_lines = number_lines(lines, issue_id, next_id(items_df, 'Line_ID'))
        if not items_df.empty and 'Issue_ID' in items_df.columns:
            items_df = items_df[items_df['Issue_ID'].apply(sid) != sid(issue_id)]

        after = records(df.loc[[idx]])[0]
        after['Items'] = new_lines
        a_sheet, audit_df = audit_update(book, 'UPDATE', str(challan_no), before=before, after=after)

        save_sheets({
            c_sheet: df,
            i_sheet: append_rows(items_df, new_lines),
            a_sheet: audit_df,
        })

    flash(f'Challan No. {challan_no} updated successfully!', 'success')
    return redirect(url_for('issue_challan.index'))


@issue_challan_bp.route('/delete/<path:challan_no>', methods=['POST'])
@safe_action('Error deleting challan.')
def delete(challan_no):
    with _excel_lock:
        book = Book.load()
        c_sheet = challan_sheet(book)
        df = book.df(c_sheet)
        if df.empty or 'Challan_No' not in df.columns:
            raise ChallanError('Challan not found.')
        idx = pick_challan_row(df, challan_no)
        if idx is None:
            raise ChallanError('Challan not found.')

        before = records(df.loc[[idx]])[0]
        issue_id = sid(before.get('Issue_ID'))
        updates = {c_sheet: df.drop(index=idx)}      # only THIS row (the number can repeat)

        i_sheet = challan_items_sheet(book)
        if i_sheet:
            items_df = book.df(i_sheet)
            if not items_df.empty and 'Issue_ID' in items_df.columns and issue_id:
                gone = items_df['Issue_ID'].apply(sid) == issue_id
                before['Items'] = records(items_df[gone])
                updates[i_sheet] = items_df[~gone]

        a_sheet, audit_df = audit_update(book, 'DELETE', str(challan_no), before=before)
        updates[a_sheet] = audit_df
        save_sheets(updates)

    flash(f'Challan No. {challan_no} deleted successfully!', 'success')
    return redirect(url_for('issue_challan.index'))
