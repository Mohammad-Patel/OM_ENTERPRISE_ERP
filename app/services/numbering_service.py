"""
Numbering service - OM Enterprise ERP
-------------------------------------
Place this file at:  app/services/numbering_service.py

ONE place that decides every challan / invoice number, so Issue Challan, Receive Challan and
Monthly Billing can never disagree with the Settings page.

Required formats
    GST party      Issue challan      OM/26-27/180, OM/26-27/181 ...
    Non-GST party  Issue challan      01, 02, 03 ...
    Receive challan (GST + Non-GST)   01, 02, 03 ...   (one running series, see note)
    GST invoice                       130/26-27, 131/26-27 ...
    Non-GST invoice                   150/26-27, 151/26-27 ...

Rule for every series:   next number = max(start number in Settings, highest used number + 1)
    - Settings only holds the START number; it is never "incremented" (no duplicates, no skips).
    - Numbers of cancelled documents stay used (never reused).
    - Old-format rows (OM/IC/0001, OM/RC/0001) are ignored by the new series.

Note on receive challans: GST and non-GST receive challans are looked up by their number in
URLs (edit / delete / print). Two challans both numbered "01" would be ambiguous, so receive
challans use ONE shared 01, 02, 03 ... series.

Settings keys (case-insensitive)
    GST_ISSUE_CHALLAN_PREFIX         OM/26-27/
    GST_ISSUE_CHALLAN_START          180
    WITHOUT_GST_ISSUE_CHALLAN_START  1
    RECEIVE_CHALLAN_START            1
    GST_INVOICE_NEXT_NO              130
    NON_GST_INVOICE_NEXT_NO          150
    FINANCIAL_YEAR                   26-27   (blank = automatic, April - March)
"""
import re
from datetime import date


def records(df):
    """DataFrame -> list of dicts (NaN becomes ''). Self-contained so this file has no cycles."""
    if df is None or getattr(df, 'empty', True):
        return []
    return [{k: ('' if v != v else v) for k, v in r.items()} for r in df.to_dict('records')]


def to_float(value, default=0.0):
    try:
        f = float(str(value).replace(',', '').strip())
        return f if f == f else default
    except (TypeError, ValueError):
        return default

GSTIN_RE = re.compile(r'^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z]$')
_NO_GST_TYPES = {'unregistered', 'composition', 'consumer', 'b2c'}

DEFAULTS = {
    'gst_issue_challan_prefix': 'OM/26-27/',
    'gst_issue_challan_start': 180,
    'without_gst_issue_challan_start': 1,
    'receive_challan_start': 1,
    'gst_invoice_next_no': 130,
    'non_gst_invoice_next_no': 150,
}
# older key names that the first Settings page used
LEGACY_KEYS = {
    'gst_issue_challan_prefix': ('gst_challan_prefix',),
    'gst_issue_challan_start': ('gst_challan_next_no',),
    'without_gst_issue_challan_start': ('non_gst_challan_next_no',),
}


# --------------------------------------------------------------------------
# Settings / GST helpers
# --------------------------------------------------------------------------
def get_settings(book):
    """{lower_case_key: value_string} from the Settings sheet."""
    out = {}
    name = book.find(exact=('settings',))
    for r in records(book.df(name)):
        key = str(r.get('Setting_Key') or '').strip().lower()
        if key:
            val = r.get('Setting_Value')
            out[key] = '' if val is None else str(val).strip()
    return out


def _setting(settings, key):
    for k in (key,) + LEGACY_KEYS.get(key, ()):
        if settings.get(k, '') != '':
            return settings[k]
    return DEFAULTS[key]


def _int_setting(settings, key):
    try:
        n = int(float(_setting(settings, key)))
        return n if n > 0 else int(DEFAULTS[key])
    except (TypeError, ValueError):
        return int(DEFAULTS[key])


def clean_gstin(value):
    return re.sub(r'\s+', '', str(value or '')).upper()


def is_valid_gstin(value):
    return bool(GSTIN_RE.match(clean_gstin(value)))


def is_gst_registered(party):
    """GST party = valid GSTIN and not unregistered / composition. Used EVERYWHERE."""
    if not party or not is_valid_gstin(party.get('GSTIN')):
        return False
    return str(party.get('Party_Type') or '').strip().lower() not in _NO_GST_TYPES


def financial_year(on, settings):
    fy = str(settings.get('financial_year', '')).strip()
    if re.fullmatch(r'\d{2}-\d{2}', fy):
        return fy
    y = on.year if on.month >= 4 else on.year - 1
    return f'{y % 100:02d}-{(y + 1) % 100:02d}'


def _issue_sheet(book):
    return book.find(exact=('issue_challan',))


def _receive_sheet(book):
    return book.find(exact=('receive_challan',))


def _bills_sheet(book):
    return book.find(exact=('monthly_bills',)) or 'Monthly_Bills'


def _trailing_int(text, pattern):
    m = re.match(pattern, str(text or '').strip())
    return int(m.group(1)) if m else None


def _free(n, used, fmt):
    """First number >= n whose formatted text is not used yet (safety net)."""
    while fmt(n) in used:
        n += 1
    return n


# --------------------------------------------------------------------------
# Issue challan
# --------------------------------------------------------------------------
def next_issue_challan_no(book, party, sheet_rows=None):
    """
    GST party      -> <prefix><n>   e.g. OM/26-27/180
    Non-GST party  -> 01, 02, 03 ...
    """
    rows = sheet_rows if sheet_rows is not None else records(book.df(_issue_sheet(book)))
    settings = get_settings(book)
    used = {str(r.get('Challan_No') or '').strip() for r in rows}

    if is_gst_registered(party):
        prefix = str(_setting(settings, 'gst_issue_challan_prefix'))
        pattern = r'^' + re.escape(prefix) + r'(\d+)$'
        highest = max([n for n in (_trailing_int(u, pattern) for u in used) if n is not None] or [0])
        n = _free(max(_int_setting(settings, 'gst_issue_challan_start'), highest + 1), used,
                  lambda x: f'{prefix}{x}')
        return f'{prefix}{n}'

    highest = max([n for n in (_trailing_int(u, r'^(\d+)$') for u in used) if n is not None] or [0])
    n = _free(max(_int_setting(settings, 'without_gst_issue_challan_start'), highest + 1), used,
              lambda x: f'{x:02d}')
    return f'{n:02d}'


# --------------------------------------------------------------------------
# Receive challan  (one shared 01, 02, 03 ... series)
# --------------------------------------------------------------------------
def next_receive_challan_no(book, sheet_rows=None):
    rows = sheet_rows if sheet_rows is not None else records(book.df(_receive_sheet(book)))
    settings = get_settings(book)
    used = {str(r.get('Receive_Challan_No') or '').strip() for r in rows}
    highest = max([n for n in (_trailing_int(u, r'^(\d+)$') for u in used) if n is not None] or [0])
    n = _free(max(_int_setting(settings, 'receive_challan_start'), highest + 1), used, lambda x: f'{x:02d}')
    return f'{n:02d}'


# --------------------------------------------------------------------------
# Monthly invoice
# --------------------------------------------------------------------------
def bill_is_gst(row):
    flag = str(row.get('Apply_GST', '')).strip().lower()
    if flag in ('yes', 'no'):
        return flag == 'yes'
    return to_float(row.get('CGST_Amount')) > 0


def next_invoice_no(book, apply_gst, bill_date=None, bills=None):
    """GST -> 130/26-27 ...   Non-GST -> 150/26-27 ..."""
    settings = get_settings(book)
    fy = financial_year(bill_date or date.today(), settings)
    start = _int_setting(settings, 'gst_invoice_next_no' if apply_gst else 'non_gst_invoice_next_no')
    rows = bills if bills is not None else records(book.df(_bills_sheet(book)))
    pattern = r'^(\d+)/' + re.escape(fy) + r'$'
    used = {str(b.get('Bill_No') or '').strip() for b in rows}
    highest = max([n for n in (_trailing_int(str(b.get('Bill_No') or ''), pattern)
                               for b in rows if bill_is_gst(b) == bool(apply_gst)) if n is not None] or [0])
    n = _free(max(start, highest + 1), used, lambda x: f'{x}/{fy}')
    return f'{n}/{fy}'
