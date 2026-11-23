"""
challan_store.py - shared helpers for the Receive Challan module (OM Enterprise ERP)
-----------------------------------------------------------------------------------
Place this file at:  app/routes/challan_store.py

Same storage style as issue_challan.py (Excel via pandas, atomic save through a temp
file, server-side numbering, audit log).

    stock with party = opening stock + issued - already returned
"""
import json
import logging
import os
import re
import shutil
import tempfile
import threading
from datetime import datetime
from functools import wraps

import pandas as pd
from flask import flash, redirect, session, url_for

logger = logging.getLogger(__name__)

EXCEL_PATH = os.path.join('data', 'om_enterprise.xlsx')

RECEIVE_SHEET = 'Receive_Challan'
RECEIVE_ITEMS_SHEET = 'Receive_Challan_Items'
DAMAGE_SHEET = 'Damage_Charges'
AUDIT_SHEET = 'Audit_Log'

DATE_FMT = '%d-%m-%Y'
DEFAULT_UNIT_WEIGHT = 10
ALLOWED_PER_PAGE = ('10', '25', '50', 'All')

# Share the SAME lock as the issue challan module, so both modules never write the
# workbook at the same time. Falls back to its own lock if that import fails.
try:
    from app.utils.challan_helpers import excel_lock            # lock used by the issue challan module
except Exception:  # pragma: no cover
    try:
        from .issue_challan import _excel_lock as excel_lock
    except Exception:
        try:
            from app.routes.issue_challan import _excel_lock as excel_lock
        except Exception:
            excel_lock = threading.RLock()


class ChallanError(Exception):
    """User-facing validation error (message is shown with flash)."""


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
    try:
        return str(session.get('username') or session.get('user') or session.get('user_name') or '')
    except RuntimeError:        # outside a request
        return ''


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
    """Append rows while keeping the existing column order (new columns go at the end)."""
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


def is_cancelled(row):
    return str(row.get('Status', '')).strip().lower() in ('cancelled', 'canceled')


# --------------------------------------------------------------------------
# Display formatters
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


def item_label(m):
    """Items master row -> 'Item Name Size'."""
    name = str(m.get('Item_Name') or m.get('Item_Code') or m.get('Item_ID') or '').strip()
    size = str(m.get('Size') or '').strip()
    if size and size.lower() not in name.lower():
        return f"{name} {size}".strip()
    return name


# --------------------------------------------------------------------------
# Dates / numbering
# --------------------------------------------------------------------------
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


def next_number(prefix, rows, field):
    """'OM/RC/' + (highest trailing number in rows[field] + 1), zero padded to 4."""
    highest = 0
    for r in rows:
        m = re.search(r'(\d+)\s*$', str(r.get(field, '')))
        if m:
            highest = max(highest, int(m.group(1)))
    return f"{prefix}{highest + 1:04d}"


def next_challan_number(book, prefix, rows, field, is_gst_customer):
    """
    Reads next challan number dynamically from Settings sheet.
    GST -> OM/26-27/180, 181...
    Without-GST -> 01, 02, 03...
    """
    s_name = book.find(exact=('settings',)) if book else None
    settings_df = book.df(s_name) if s_name else pd.DataFrame()

    if is_gst_customer:
        key = 'GST_CHALLAN_NEXT_NO'
        default_start = 180
        pfx_key = 'GST_CHALLAN_PREFIX'
        actual_prefix = prefix
        
        if not settings_df.empty and 'Setting_Key' in settings_df.columns:
            pfx_match = settings_df[settings_df['Setting_Key'].astype(str).str.strip() == pfx_key]
            if not pfx_match.empty:
                actual_prefix = str(pfx_match.iloc[0]['Setting_Value']).strip()

        start_num = default_start
        if not settings_df.empty and 'Setting_Key' in settings_df.columns:
            match = settings_df[settings_df['Setting_Key'].astype(str).str.strip() == key]
            if not match.empty:
                try:
                    start_num = int(str(match.iloc[0]['Setting_Value']).strip())
                except ValueError:
                    pass
        return f"{actual_prefix}{start_num}"
    else:
        key = 'NON_GST_CHALLAN_NEXT_NO'
        start_num = 1
        if not settings_df.empty and 'Setting_Key' in settings_df.columns:
            match = settings_df[settings_df['Setting_Key'].astype(str).str.strip() == key]
            if not match.empty:
                try:
                    start_num = int(str(match.iloc[0]['Setting_Value']).strip())
                except ValueError:
                    pass
        return f"{start_num:02d}"

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


def receive_sheet(book):
    return book.find(exact=('receive_challan',))


def receive_items_sheet(book):
    return book.find(exact=('receive_challan_items',))


def damage_sheet(book):
    return book.find(exact=('damage_charges',))


def _issue_sheet(book):
    return book.find(exact=('issue_challan',))


def _issue_items_sheet(book):
    return book.find(exact=('issue_challan_items',))


def _opening_sheet(book):
    return book.find(exact=('openingstock',))


def _audit_sheet(book):
    return book.find(exact=('audit_log',))


# Same lookup rules as issue_challan.py, so both modules see the same master data.
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


def audit_update(book, action, entity_id, before=None, after=None, reason='',
                 entity_type='Issue_Challan'):
    """Returns (sheet_name, dataframe) with one audit row appended."""
    name = _audit_sheet(book) or AUDIT_SHEET
    df = book.df(name)
    row = {
        'Audit_ID': next_id(df, 'Audit_ID'),
        'Event_Time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'Username': current_user(),
        'Action': action,
        'Entity_Type': entity_type,
        'Entity_ID': entity_id,
        'Reason': reason,
        'Before_Data': json.dumps(before, default=str, ensure_ascii=False) if before else '',
        'After_Data': json.dumps(after, default=str, ensure_ascii=False) if after else '',
    }
    return name, append_rows(df, [row])


def safe_action(error_message, endpoint='issue_challan.index'):
    """Common error handling for create / edit / delete (redirects to `endpoint`)."""
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
            return redirect(url_for(endpoint))
        return wrapper
    return decorator


# --------------------------------------------------------------------------
# Stock lying with a party at a site
# --------------------------------------------------------------------------
_IGNORED_OPENING_STATUS = ('cancelled', 'canceled', 'deleted', 'inactive')


def stock_summary(book, party_id, site_id, exclude_receive_id=None):
    """
    {item_id(str): {'opening', 'issued', 'received', 'outstanding'}} for one party + site.

        opening     OpeningStock: Verified_Qty (if filled) else Estimated_Qty
        issued      Issue_Challan_Items.Quantity of non-cancelled issue challans
        received    Receive_Challan_Items.Returned_Qty (good + damaged) of non-cancelled
                    receive challans; `exclude_receive_id` is skipped (used while editing)
        outstanding opening + issued - received
    """
    p, s = sid(party_id), sid(site_id)
    exclude = sid(exclude_receive_id) if exclude_receive_id not in (None, '') else None
    summary = {}

    def slot(iid):
        return summary.setdefault(iid, {'opening': 0.0, 'issued': 0.0,
                                        'received': 0.0, 'outstanding': 0.0})

    # opening stock
    for r in records(book.df(_opening_sheet(book))):
        if sid(r.get('Party_ID')) != p or sid(r.get('Site_ID')) != s:
            continue
        if str(r.get('Status', '')).strip().lower() in _IGNORED_OPENING_STATUS:
            continue
        iid = sid(r.get('Item_ID'))
        if not iid:
            continue
        verified = r.get('Verified_Qty')
        qty = to_float(verified) if str(verified).strip() != '' else to_float(r.get('Estimated_Qty'))
        slot(iid)['opening'] += qty

    # issued
    valid_issue = {sid(c.get('Issue_ID')) for c in records(book.df(_issue_sheet(book)))
                   if sid(c.get('Party_ID')) == p and sid(c.get('Site_ID')) == s
                   and not is_cancelled(c)}
    for r in records(book.df(_issue_items_sheet(book))):
        if sid(r.get('Issue_ID')) in valid_issue:
            iid = sid(r.get('Item_ID'))
            if iid:
                slot(iid)['issued'] += to_float(r.get('Quantity'))

    # already returned
    valid_receive = {sid(c.get('Receive_ID')) for c in records(book.df(receive_sheet(book)))
                     if sid(c.get('Party_ID')) == p and sid(c.get('Site_ID')) == s
                     and not is_cancelled(c) and sid(c.get('Receive_ID')) != exclude}
    for r in records(book.df(receive_items_sheet(book))):
        if sid(r.get('Receive_ID')) in valid_receive:
            iid = sid(r.get('Item_ID'))
            if iid:
                returned = to_float(r.get('Returned_Qty'))
                if not returned:
                    returned = to_float(r.get('Accepted_Qty')) + to_float(r.get('Damaged_Qty'))
                slot(iid)['received'] += returned

    for v in summary.values():
        v['outstanding'] = v['opening'] + v['issued'] - v['received']
    return summary


# --------------------------------------------------------------------------
# Helpers used by the Stock Ledger page (stock_ledger.py)
# --------------------------------------------------------------------------
# Columns that link a challan header to its item lines
ISSUE_KEYS = ('Issue_ID', 'Challan_No', 'Issue_No')
RECEIVE_KEYS = ('Receive_ID', 'Receive_No', 'Receive_Challan_No')


def _doc_keys(row, keys):
    """Non-empty link keys of a header / item row, e.g. ['Issue_ID:5', 'Challan_No:IC-001']."""
    out = []
    for col in keys:
        val = sid(row.get(col))
        if val:
            out.append(f'{col}:{val}')
    return out


def _opening_qty(r):
    """Opening stock qty of one OpeningStock row: Verified_Qty if filled, else Estimated_Qty."""
    verified = r.get('Verified_Qty')
    if str(verified).strip().lower() not in ('', 'nan', 'none'):
        return to_float(verified)
    return to_float(r.get('Estimated_Qty'))


def _returned_split(r):
    """(total_returned, damaged, good) of one Receive_Challan_Items row."""
    damaged = to_float(r.get('Damaged_Qty'))
    total = to_float(r.get('Returned_Qty'))
    if not total:
        total = to_float(r.get('Accepted_Qty')) + damaged
    return total, damaged, total - damaged


# --------------------------------------------------------------------------
# Compatibility names used by app/services/billing_service.py (Monthly Billing)
# --------------------------------------------------------------------------
def challan_sheet(book):
    """Issue challan header sheet."""
    return _issue_sheet(book)


def challan_items_sheet(book):
    """Issue challan item lines sheet."""
    return _issue_items_sheet(book)


opening_qty = _opening_qty      # same function, name without the underscore



def log_audit(book, action, details, username="Admin"):
    """Appends an audit log entry to the Excel sheet."""
    try:
        from datetime import datetime
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # Check if 'Audit_Logs' sheet exists, if not create a default structure
        if 'Audit_Logs' not in book.sheets:
            # Create sheet data with headers
            book.sheets['Audit_Logs'] = [
                ['Timestamp', 'User', 'Action', 'Details']
            ]
        
        # Append log row
        book.sheets['Audit_Logs'].append([timestamp, username, action, details])
        
        # Save changes back to Excel
        save_sheets({'Audit_Logs': book.sheets['Audit_Logs']})
    except Exception as e:
        logger.exception("Error writing audit log")