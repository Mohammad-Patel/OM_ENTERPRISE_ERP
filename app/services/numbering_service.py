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

Challan series (4 independent series)
    GST Issue          GST_ISSUE_CHALLAN_NEXT_NO            OM/26-27/180, 181 ...
    Without-GST Issue  WITHOUT_GST_ISSUE_CHALLAN_NEXT_NO    01, 02, 03 ...
    GST Receive        GST_RECEIVE_CHALLAN_NEXT_NO          01, 02, 03 ...
    Without-GST Receive WITHOUT_GST_RECEIVE_CHALLAN_NEXT_NO 01, 02, 03 ...

    The *_NEXT_NO value in Settings is the number the NEXT challan of that series gets.
    It is written back (value + 1) in the SAME save as the challan, so the series simply counts up
    from whatever was typed in Settings - 1 .. 10, then 89 .. 95, then 1 .. 10 again - without
    looking at old challans and without touching any other series.
    While a *_NEXT_NO key is not saved yet (first run / old workbook) the old rule is used so existing
    data keeps working:  max(old START value, highest used number + 1).

Invoices:   next number = max(start number in Settings, highest used number + 1)
    - Numbers of cancelled documents stay used (never reused).
    - Old-format rows (OM/IC/0001, OM/RC/0001) are ignored by the new series.

Settings keys (case-insensitive)
    GST_ISSUE_CHALLAN_PREFIX         OM/26-27/
    GST_ISSUE_CHALLAN_NEXT_NO        180      (old: GST_ISSUE_CHALLAN_START)
    WITHOUT_GST_ISSUE_CHALLAN_NEXT_NO 1       (old: WITHOUT_GST_ISSUE_CHALLAN_START)
    GST_RECEIVE_CHALLAN_NEXT_NO      1        (old: RECEIVE_CHALLAN_START)
    WITHOUT_GST_RECEIVE_CHALLAN_NEXT_NO 1     (old: RECEIVE_CHALLAN_START)
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
    'gst_issue_challan_next_no': 0,
    'without_gst_issue_challan_next_no': 0,
    'gst_receive_challan_next_no': 0,
    'without_gst_receive_challan_next_no': 0,
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
# Challan series  (4 independent series, explicit counter kept in Settings)
# --------------------------------------------------------------------------
ISSUE, RECEIVE = 'issue', 'receive'
NEXT_KEYS = {
    (ISSUE, True): 'GST_ISSUE_CHALLAN_NEXT_NO',
    (ISSUE, False): 'WITHOUT_GST_ISSUE_CHALLAN_NEXT_NO',
    (RECEIVE, True): 'GST_RECEIVE_CHALLAN_NEXT_NO',
    (RECEIVE, False): 'WITHOUT_GST_RECEIVE_CHALLAN_NEXT_NO',
}
NEXT_KEY_NAMES = tuple(NEXT_KEYS.values())


def series_key(kind, gst):
    return NEXT_KEYS[(kind, bool(gst))]


def _explicit_next(settings, key):
    """Counter saved in Settings for this series, or None when it was never set."""
    raw = settings.get(key.lower(), '')
    try:
        n = int(float(raw))
        return n if n > 0 else None
    except (TypeError, ValueError):
        return None


def _fmt_no(kind, gst, n, prefix):
    if kind == ISSUE and gst:
        return f'{prefix}{n}'
    return f'{n:02d}'


def _legacy_next(book, kind, gst, settings, rows):
    """Old rule, used until the series counter is saved for the first time."""
    if kind == ISSUE:
        col = 'Challan_No'
        rows = rows if rows is not None else records(book.df(_issue_sheet(book)))
    else:
        col = 'Receive_Challan_No'
        rows = rows if rows is not None else records(book.df(_receive_sheet(book)))
    used = {str(r.get(col) or '').strip() for r in rows}
    prefix = str(_setting(settings, 'gst_issue_challan_prefix'))
    if kind == ISSUE and gst:
        pattern = r'^' + re.escape(prefix) + r'(\d+)$'
        start = _int_setting(settings, 'gst_issue_challan_start')
    else:
        pattern = r'^(\d+)$'
        start = _int_setting(settings, 'without_gst_issue_challan_start' if kind == ISSUE
                             else 'receive_challan_start')
    highest = max([n for n in (_trailing_int(u, pattern) for u in used) if n is not None] or [0])
    return _free(max(start, highest + 1), used, lambda x: _fmt_no(kind, gst, x, prefix))


def peek_challan(book, kind, gst, rows=None, settings=None):
    """-> (challan_no_text, number_int, settings_key). Read only - nothing is saved."""
    settings = settings if settings is not None else get_settings(book)
    key = series_key(kind, gst)
    n = _explicit_next(settings, key)
    if n is None:
        n = _legacy_next(book, kind, gst, settings, rows)
    prefix = str(_setting(settings, 'gst_issue_challan_prefix'))
    return _fmt_no(kind, gst, n, prefix), n, key


def settings_sheet_name(book):
    return book.find(exact=('settings',)) or 'Settings'


def settings_with(book, values):
    """
    -> (sheet_name, DataFrame) : the Settings sheet with {KEY: value} set (added when missing).
    Put it in the same save_sheets() call as the challan so the counter and the challan are
    written together or not at all.
    """
    import pandas as pd
    name = settings_sheet_name(book)
    df = book.df(name)
    if df is None or df.empty or 'Setting_Key' not in df.columns:
        df = pd.DataFrame(columns=['Setting_Key', 'Setting_Value', 'Description'])
    df = df.copy()
    for col in ('Setting_Value', 'Description'):
        if col not in df.columns:
            df[col] = ''
        df[col] = df[col].astype(object)
    keys = df['Setting_Key'].astype(str).str.strip().str.upper()
    add = []
    for key, value in values.items():
        hit = df.index[keys == key.upper()]
        if len(hit):
            df.at[hit[0], 'Setting_Value'] = value
        else:
            add.append({'Setting_Key': key, 'Setting_Value': value,
                        'Description': 'Next challan number of this series (auto-increments)'})
    if add:
        df = pd.concat([df, pd.DataFrame(add)], ignore_index=True)
    return name, df


def take_challan_no(book, kind, party, rows=None):
    """
    Number for a challan that is being SAVED now.
    -> (challan_no_text, {sheet_name: settings_dataframe_with_counter_plus_one})
    Add the returned sheet to the save_sheets() call that stores the challan.
    """
    gst = is_gst_registered(party)
    text, n, key = peek_challan(book, kind, gst, rows)
    name, df = settings_with(book, {key: str(n + 1)})
    return text, {name: df}


def preview_challan_numbers(book, kind, rows=None):
    """{'gst': '...', 'non_gst': '...'} for the form (read only)."""
    settings = get_settings(book)
    return {
        'gst': peek_challan(book, kind, True, rows, settings)[0],
        'non_gst': peek_challan(book, kind, False, rows, settings)[0],
    }


def next_issue_challan_no(book, party, sheet_rows=None):
    """Kept for old callers: the number the next issue challan of this party's series would get."""
    return peek_challan(book, ISSUE, is_gst_registered(party), sheet_rows)[0]


def next_receive_challan_no(book, party=None, sheet_rows=None):
    """Kept for old callers: next receive challan number for this party's series."""
    return peek_challan(book, RECEIVE, is_gst_registered(party), sheet_rows)[0]


def current_series_values(book):
    """{KEY: value to show on the Settings page} for the 4 series (GST issue shows the full number)."""
    settings = get_settings(book)
    out = {}
    for (kind, gst), key in NEXT_KEYS.items():
        text, n, _ = peek_challan(book, kind, gst, None, settings)
        out[key] = text if (kind == ISSUE and gst) else str(n)
    return out


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
