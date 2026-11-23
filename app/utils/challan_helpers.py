"""
Shared helpers for the challan modules (Issue Challan / Receive Challan)
------------------------------------------------------------------------
- Excel (data/om_enterprise.xlsx) access: Book snapshot + atomic multi-sheet save
- Lock, audit log, error handling decorator
- Validation / formatting helpers
- STOCK LOGIC:  stock with party at a site
        = Opening stock + Issued (Issue_Challan_Items) - Returned (Receive_Challan_Items)
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

CHALLAN_SHEET = 'Issue_Challan'
CHALLAN_ITEMS_SHEET = 'Issue_Challan_Items'
RECEIVE_SHEET = 'Receive_Challan'
RECEIVE_ITEMS_SHEET = 'Receive_Challan_Items'
DAMAGE_SHEET = 'Damage_Charges'
AUDIT_SHEET = 'Audit_Log'

DATE_FMT = '%d-%m-%Y'
DEFAULT_GOODS_RATE = 59
DEFAULT_UNIT_WEIGHT = 10
ALLOWED_PER_PAGE = ('10', '25', '50', 'All')

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
    return str(session.get('username') or session.get('user') or session.get('user_name') or '')


def is_cancelled(row):
    return str(row.get('Status', '')).strip().lower() in ('cancelled', 'canceled')


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
    """Append rows while keeping the existing column order (new columns go to the end)."""
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


def next_number(prefix, rows, key):
    """OM/IC/0003 style numbers: highest trailing number + 1."""
    highest = 0
    for r in rows:
        m = re.search(r'(\d+)\s*$', str(r.get(key, '')))
        if m:
            highest = max(highest, int(m.group(1)))
    return f"{prefix}{highest + 1:04d}"


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


def normalize_date(value):
    value = (value or '').strip()
    if not value:
        return datetime.now().strftime(DATE_FMT)
    d = parse_any_date(value)
    if d is None:
        raise ChallanError('Invalid date. Please use DD-MM-YYYY.')
    return d.strftime(DATE_FMT)


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


def item_label(item):
    """{'Item_Name': 'PLATE', 'Size': '3 X 2'} -> 'PLATE 3 X 2'"""
    name = str(item.get('Item_Name') or item.get('Item_Code') or item.get('Item_ID') or '').strip()
    size = str(item.get('Size') or '').strip()
    if size and size.lower() not in name.lower():
        return f"{name} {size}".strip()
    return name


# --------------------------------------------------------------------------
# Excel access
# --------------------------------------------------------------------------
class Book:
    """Snapshot of all sheets (the file is closed immediately after reading)."""

    def __init__(self, sheets):
        self.sheets = sheets
        self._lower = {n.lower(): n for n in sheets}

    @classmethod
    def load(cls):
        if not os.path.exists(EXCEL_PATH):
            return cls({})
        with excel_lock:
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

    def challan_sheet(self):
        return self.find(exact=('issue_challan', 'issuechallans'), keywords=('issue_challan', 'challan', 'issue'), exclude=('item',))

    def challan_items_sheet(self):
        return self.find(exact=('issue_challan_items', 'issueitems'), keywords=('issue_challan_items',))

    def receive_sheet(self):
        return self.find(exact=('receive_challan', 'receivechallans'), keywords=('receive_challan', 'receive'))

    def receive_items_sheet(self):
        return self.find(exact=('receive_challan_items', 'receiveitems'), keywords=('receive_challan_items',))

    def load_parties(self):
        s = self.find(exact=('customers', 'customer', 'parties', 'party'),
                      keywords=('customer', 'client'),
                      exclude=('challan', 'issue', 'site', 'item'))
        return records(self.df(s))

    def load_sites(self):
        s = self.find(exact=('sitelocations', 'sitelocation', 'sites', 'site'),
                      keywords=('sitelocation',),
                      exclude=('challan', 'issue', 'item'))
        return records(self.df(s))

    def load_items_master(self):
        s = self.find(exact=('items', 'item'),
                      keywords=('product',),
                      exclude=('challan', 'issue'))
        return records(self.df(s))


def save_sheets(updates):
    """
    Save several sheets in ONE operation.
    Written to a temp copy first and swapped in only on success.
    """
    with excel_lock:
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


def audit_update(book, action, entity_id, before=None, after=None, reason='', entity_type='Issue_Challan'):
    name = book.find(exact=('audit_log',)) or AUDIT_SHEET
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