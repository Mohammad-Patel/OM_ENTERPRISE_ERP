"""
Monthly Billing service - OM Enterprise ERP
-------------------------------------------
Place this file at:  app/services/billing_service.py

Reads the challan logs (Issue / Receive / OpeningStock / Damage_Charges) from
data/om_enterprise.xlsx and builds an editable monthly rental invoice.

Only names that really exist in app/routes/challan_store.py are imported.

Public API (used by app/routes/monthly_billing.py)
    bills_sheet(book)                         name of the Monthly_Bills sheet
    build_draft(book, party, site, from, to)  auto-calculated draft invoice (dict)
    parse_payload(form, book, bill_no=None)   validated invoice from the editor form
    draft_from_form(book, form)               draft to re-show the editor after an error
    create_bill(book, payload)                -> (sheets_to_save, bill_no)
    update_bill(book, bill_no, payload)       -> sheets_to_save
    cancel_bill(book, bill_no, reason)        -> sheets_to_save
    load_bill(book, bill_no)                  saved invoice as a draft dict
    print_context(book, bill_no)              everything print.html needs
    list_bills(book)                          rows for the index page
    compute_totals(rows, apply_gst, cgst, sgst), fmt_date(d), inr(x)

Business rules
    Days   = (To - From) + 1                      (both dates are billed days)
    Number = NOS x Days
    Amount = Number x Rate
    Damage / extra rows : Amount = NOS x Rate     (no dates)
    CGST + SGST (9% each by default) only when the party has a VALID GSTIN
    Net total = taxable + GST, rounded to the nearest rupee (Round_Off shown)
    Issue day counts as a rental day; return day counts too (reduced qty from next day)
    Invoice no.: GST bills 130/26-27, 131/26-27 ...  Non-GST bills 150/26-27, 151/26-27 ...
"""
import re
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

import pandas as pd

from app.services import numbering_service as _num

from app.routes.challan_store import (
    DATE_FMT, ChallanError, append_rows, audit_update, current_user, damage_sheet,
    is_cancelled, item_label, load_items_master, load_parties, load_sites, next_id,
    norm_id, number_to_words, parse_any_date, receive_items_sheet, receive_sheet,
    records, sid, to_float,
)

BILL_SHEET = 'Monthly_Bills'
LINE_SHEET = 'Monthly_Bill_Lines'
DEFAULT_GST_RATE = Decimal('9')
DEFAULT_GST_START = 130
DEFAULT_NON_GST_START = 150
MAX_PERIOD_DAYS = 400
MAX_ROWS = 500
MAX_AMOUNT = Decimal('100000000')          # sanity limit for qty / rate

ROW_TYPES = ('rental', 'damage', 'extra')
LINE_TYPE = {'rental': 'Rental', 'damage': 'Damage', 'extra': 'Extra'}
LINE_TYPE_REV = {'rental': 'rental', 'damage': 'damage', 'extra': 'extra', 'custom': 'extra'}

GSTIN_RE = re.compile(r'^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z]$')
_NO_GST_TYPES = {'unregistered', 'composition', 'consumer', 'b2c'}

DEFAULT_COMPANY = {
    'name': 'OM ENTERPRISE',
    'tagline': 'RCC FORM-WORK & SCAFFOLDING RENTALS',
    'address': 'PLOT NO: 1046, KIRC ROAD, NEAR SINDBAD HOTEL, KALOL, '
               'TALUKA-KALOL, GANDHINAGAR, GUJARAT - 382715',
    'phone': '+91 9998802167, 9824219138',
    'gstin': '', 'email': '', 'state': 'Gujarat', 'state_code': '24',
    'bank_name': '', 'bank_account': '', 'bank_ifsc': '', 'bank_branch': '',
    'terms': '', 'signatory': 'OM ENTERPRISE',
}


# ----------------------------------------------------------------------------
# Small helpers
# ----------------------------------------------------------------------------
def D(value, default='0'):
    s = str(value if value is not None else '').strip().replace(',', '')
    if s == '':
        s = default
    try:
        d = Decimal(s)
    except InvalidOperation:
        raise ChallanError(f'"{value}" is not a valid number.')
    if not d.is_finite():
        raise ChallanError(f'"{value}" is not a valid number.')
    return d


def q2(x):
    return Decimal(x).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)


def q0(x):
    return Decimal(x).quantize(Decimal('1'), rounding=ROUND_HALF_UP)


def num(x):
    """Decimal -> int when whole, else float (JSON / Excel friendly)."""
    x = Decimal(x)
    return int(x) if x == x.to_integral_value() else float(x)


def iso(dt):
    return dt.isoformat() if dt else ''


def fmt_date(dt):
    if isinstance(dt, str):
        dt = parse_any_date(dt)
    return dt.strftime(DATE_FMT) if dt else ''


def inr(value, places=2):
    """Indian digit grouping: 1234567.5 -> '12,34,567.50'."""
    d = q2(Decimal(str(value or 0))) if places == 2 else q0(Decimal(str(value or 0)))
    sign = '-' if d < 0 else ''
    s = f'{abs(d):.{places}f}'
    whole, _, frac = s.partition('.')
    if len(whole) > 3:
        head, tail = whole[:-3], whole[-3:]
        groups = []
        while len(head) > 2:
            groups.insert(0, head[-2:])
            head = head[:-2]
        if head:
            groups.insert(0, head)
        whole = ','.join(groups + [tail])
    return f'{sign}{whole}' + (f'.{frac}' if places else '')


def qty_str(value):
    d = Decimal(str(value or 0))
    return f'{int(d):,}' if d == d.to_integral_value() else f'{float(d):,.2f}'.rstrip('0').rstrip('.')


def consolidate_damage_rows(rows):
    """Same item/description ke damage charges ko group karke quantities aur amounts jod deta hai."""
    rental_rows = [r for r in rows if r['type'] == 'rental']
    damage_rows = [r for r in rows if r['type'] == 'damage']
    extra_rows = [r for r in rows if r['type'] == 'extra']
    
    merged_damage = {}
    for dr in damage_rows:
        key = (dr.get('item_id'), dr.get('desc'))
        if key in merged_damage:
            merged_damage[key]['qty'] = num(Decimal(str(merged_damage[key]['qty'])) + Decimal(str(dr['qty'])))
        else:
            merged_damage[key] = dict(dr)
            
    final_damage = list(merged_damage.values())
    return rental_rows + final_damage + extra_rows


# ----------------------------------------------------------------------------
# Sheets
# ----------------------------------------------------------------------------
def bills_sheet(book):
    return book.find(exact=('monthly_bills',)) or BILL_SHEET


def lines_sheet(book):
    return book.find(exact=('monthly_bill_lines',)) or LINE_SHEET


def _issue_sheet(book):
    return book.find(exact=('issue_challan',))


def _issue_items_sheet(book):
    return book.find(exact=('issue_challan_items',))


def _opening_sheet(book):
    return book.find(exact=('openingstock', 'opening_stock'))


# ----------------------------------------------------------------------------
# Settings (keys are case-insensitive: GST_INVOICE_NEXT_NO == gst_invoice_next_no)
# ----------------------------------------------------------------------------
def get_settings(book):
    out = {}
    name = book.find(exact=('settings',))
    for r in records(book.df(name)):
        key = str(r.get('Setting_Key') or '').strip().lower()
        if key:
            out[key] = str(r.get('Setting_Value') if r.get('Setting_Value') is not None else '').strip()
    return out


def company_info(book):
    s = get_settings(book)
    c = dict(DEFAULT_COMPANY)
    mapping = {
        'name': ('company_name',), 'tagline': ('company_tagline',),
        'address': ('company_address',), 'phone': ('company_phone', 'company_contact'),
        'gstin': ('company_gstin',), 'email': ('company_email',),
        'state': ('company_state',), 'state_code': ('company_state_code',),
        'bank_name': ('bank_name',), 'bank_account': ('bank_account_no', 'bank_account'),
        'bank_ifsc': ('bank_ifsc',), 'bank_branch': ('bank_branch',),
        'terms': ('invoice_terms',), 'signatory': ('invoice_signatory',),
    }
    for field, keys in mapping.items():
        for k in keys:
            if s.get(k):
                c[field] = s[k]
                break
    return c


def get_gst_rates(book):
    s = get_settings(book)
    out = []
    for key in ('cgst_rate', 'sgst_rate'):
        try:
            out.append(Decimal(s.get(key, '').rstrip('%').strip() or str(DEFAULT_GST_RATE)))
        except InvalidOperation:
            out.append(DEFAULT_GST_RATE)
    return out[0], out[1]


# ----------------------------------------------------------------------------
# GST
# ----------------------------------------------------------------------------
def clean_gstin(value):
    return re.sub(r'\s+', '', str(value or '')).upper()


def is_valid_gstin(value):
    return bool(GSTIN_RE.match(clean_gstin(value)))


def is_gst_registered(party):
    """GST only for a party with a VALID GSTIN that is not unregistered / composition."""
    return _num.is_gst_registered(party)


# ----------------------------------------------------------------------------
# Party / site / period
# ----------------------------------------------------------------------------
def _party_site(book, party_id, site_id):
    party = next((p for p in load_parties(book) if sid(p.get('Party_ID')) == sid(party_id)), None)
    site = next((s for s in load_sites(book) if sid(s.get('Site_ID')) == sid(site_id)), None)
    if not party:
        raise ChallanError('Please select a valid party.')
    if not site:
        raise ChallanError('Please select a valid site.')
    if sid(site.get('Party_ID')) != sid(party_id):
        raise ChallanError('Selected site does not belong to the selected party.')
    return party, site


def check_period(start, end):
    if not start or not end:
        raise ChallanError('Please select the billing period (From and To dates).')
    if start > end:
        raise ChallanError('Billing period: "From" date cannot be after the "To" date.')
    if (end - start).days + 1 > MAX_PERIOD_DAYS:
        raise ChallanError(f'Billing period cannot be longer than {MAX_PERIOD_DAYS} days.')


# ----------------------------------------------------------------------------
# Quantity at the party's site, day by day
# ----------------------------------------------------------------------------
def _opening_qty(row):
    for col in ('Verified_Qty', 'Estimated_Qty', 'Quantity', 'Qty'):
        v = row.get(col)
        if str(v).strip() not in ('', 'None', 'nan'):
            return to_float(v)
    return 0.0


def _returned_qty(row):
    total = to_float(row.get('Returned_Qty'))
    if not total:
        total = to_float(row.get('Accepted_Qty')) + to_float(row.get('Damaged_Qty'))
    return total


def _header_maps(headers, id_col, no_cols):
    by_id, by_no = {}, {}
    for h in headers:
        v = sid(h.get(id_col))
        if v:
            by_id[v] = h
        for c in no_cols:
            v = sid(h.get(c))
            if v:
                by_no[v] = h
    return by_id, by_no


def _resolve(item, maps, id_col, no_cols):
    """Find the header of an item row by ID first, else by challan number."""
    by_id, by_no = maps
    v = sid(item.get(id_col))
    if v and v in by_id:
        return by_id[v]
    for c in no_cols:
        v = sid(item.get(c))
        if v and v in by_no:
            return by_no[v]
    return None


ISSUE_NO_COLS = ('Challan_No', 'Issue_No')
RECEIVE_NO_COLS = ('Receive_No', 'Receive_Challan_No')


def quantity_events(book, party_id, site_id, warnings):
    """
    {item_id: [(effective_date, actual_date, delta_qty), ...]}
      opening stock / issue : counts from the day itself
      return                : the returned qty is still billed on the return day,
                              so it drops from the NEXT day (effective = date + 1)
    """
    p, s = sid(party_id), sid(site_id)
    ev = {}

    def add(item_id, eff, actual, delta):
        if delta:
            ev.setdefault(sid(item_id), []).append((eff, actual, Decimal(str(delta))))

    # opening stock lying with the party
    for r in records(book.df(_opening_sheet(book))):
        if (sid(r.get('Party_ID')) != p or sid(r.get('Site_ID')) != s or not sid(r.get('Item_ID'))
                or 'godown' in str(r.get('Location_Type', '')).lower()):
            continue
        if str(r.get('Status', '')).strip().lower() in ('cancelled', 'canceled', 'deleted', 'inactive'):
            continue
        dt = parse_any_date(r.get('Verified_Date')) or parse_any_date(r.get('Estimated_Date')) or date.min
        add(r.get('Item_ID'), dt, dt, _opening_qty(r))

    # issue challans
    issue_headers = [c for c in records(book.df(_issue_sheet(book)))
                     if sid(c.get('Party_ID')) == p and sid(c.get('Site_ID')) == s]
    maps = _header_maps(issue_headers, 'Issue_ID', ISSUE_NO_COLS)
    for c in issue_headers:
        if not is_cancelled(c) and parse_any_date(c.get('Challan_Date')) is None:
            warnings.append(f"Issue challan {c.get('Challan_No')} has an invalid date and was ignored.")
    for r in records(book.df(_issue_items_sheet(book))):
        h = _resolve(r, maps, 'Issue_ID', ISSUE_NO_COLS)
        if h is None or is_cancelled(h):
            continue
        dt = parse_any_date(h.get('Challan_Date'))
        if dt:
            add(r.get('Item_ID'), dt, dt, to_float(r.get('Quantity')))

    # receive challans
    recv_headers = [c for c in records(book.df(receive_sheet(book)))
                    if sid(c.get('Party_ID')) == p and sid(c.get('Site_ID')) == s]
    maps = _header_maps(recv_headers, 'Receive_ID', RECEIVE_NO_COLS)
    for c in recv_headers:
        if not is_cancelled(c) and parse_any_date(c.get('Receive_Date')) is None:
            warnings.append(f"Receive challan {c.get('Receive_Challan_No') or c.get('Receive_No')} "
                            f"has an invalid date and was ignored.")
    for r in records(book.df(receive_items_sheet(book))):
        h = _resolve(r, maps, 'Receive_ID', RECEIVE_NO_COLS)
        if h is None or is_cancelled(h):
            continue
        dt = parse_any_date(h.get('Receive_Date'))
        if dt:
            add(r.get('Item_ID'), dt + timedelta(days=1), dt, -_returned_qty(r))
    return ev


def item_intervals(events, start, end):
    """Split [start, end] into runs of constant quantity. Returns all runs (qty may be <= 0)."""
    cuts = sorted({eff for eff, _, _ in events if start < eff <= end})
    starts = [start] + cuts
    out = []
    for i, s in enumerate(starts):
        e = starts[i + 1] - timedelta(days=1) if i + 1 < len(starts) else end
        qty = sum((delta for eff, _, delta in events if eff <= s), Decimal('0'))
        if out and out[-1]['qty'] == qty and out[-1]['to'] + timedelta(days=1) == s:
            out[-1]['to'] = e
        else:
            out.append({'from': s, 'to': e, 'qty': qty})
    return out


def closing_balance(events, end):
    return sum((delta for _, actual, delta in events if actual <= end), Decimal('0'))


def balance_text(events, end, master):
    parts = []
    for iid, evs in events.items():
        qty = closing_balance(evs, end)
        if qty > 0:
            parts.append((item_label(master.get(iid, {'Item_ID': iid})), qty))
    parts.sort(key=lambda x: x[0].lower())
    return ', '.join(f'{label} = {num(qty)}' for label, qty in parts)


# ----------------------------------------------------------------------------
# Money
# ----------------------------------------------------------------------------
def row_amounts(row):
    """-> (number, amount). number is None for damage / extra rows."""
    qty, rate = D(row.get('qty')), D(row.get('rate'))
    if row.get('type') == 'rental':
        number = qty * D(row.get('days'))
        return number, q2(number * rate)
    return None, q2(qty * rate)


def compute_totals(rows, apply_gst, cgst_rate, sgst_rate):
    taxable = Decimal('0')
    out_rows = []
    for r in rows:
        number, amount = row_amounts(r)
        taxable += amount
        out_rows.append({'number': num(number) if number is not None else None, 'amount': float(amount)})
    taxable = q2(taxable)
    cgst = q2(taxable * D(cgst_rate) / 100) if apply_gst else Decimal('0.00')
    sgst = q2(taxable * D(sgst_rate) / 100) if apply_gst else Decimal('0.00')
    total = taxable + cgst + sgst
    net = q0(total)
    return {'rows': out_rows, 'taxable': float(taxable), 'cgst': float(cgst), 'sgst': float(sgst),
            'total': float(total), 'round_off': float(net - total), 'net': float(net)}


# ----------------------------------------------------------------------------
# Invoice numbering  (GST: 130/26-27 ...   Non-GST: 150/26-27 ...)
# ----------------------------------------------------------------------------
def financial_year(on, settings):
    fy = str(settings.get('financial_year', '')).strip()
    if re.fullmatch(r'\d{2}-\d{2}', fy):
        return fy
    y = on.year if on.month >= 4 else on.year - 1
    return f'{y % 100:02d}-{(y + 1) % 100:02d}'


def _bill_is_gst(row):
    flag = str(row.get('Apply_GST', '')).strip().lower()
    if flag in ('yes', 'no'):
        return flag == 'yes'
    return to_float(row.get('CGST_Amount')) > 0


def _start_no(settings, key, default):
    try:
        n = int(float(settings.get(key, '') or default))
        return n if n > 0 else default
    except ValueError:
        return default


def next_bill_no(book, apply_gst, bill_date=None):
    """GST 130/26-27 ... | Non-GST 150/26-27 ... (single source: numbering_service)."""
    return _num.next_invoice_no(book, apply_gst, bill_date)


# ----------------------------------------------------------------------------
# Overlaps & damage charges
# ----------------------------------------------------------------------------
def find_overlaps(book, party_id, site_id, start, end, exclude_bill_id=None):
    out = []
    for b in records(book.df(bills_sheet(book))):
        if (sid(b.get('Party_ID')) != sid(party_id) or sid(b.get('Site_ID')) != sid(site_id)
                or is_cancelled(b)
                or (exclude_bill_id is not None and sid(b.get('Bill_ID')) == sid(exclude_bill_id))):
            continue
        ps, pe = parse_any_date(b.get('Period_Start')), parse_any_date(b.get('Period_End'))
        if ps and pe and ps <= end and start <= pe:
            out.append({'bill_no': b.get('Bill_No'), 'from': fmt_date(ps), 'to': fmt_date(pe)})
    return out


def pending_damage_charges(book, party_id, site_id, end):
    recv = [c for c in records(book.df(receive_sheet(book)))
            if sid(c.get('Party_ID')) == sid(party_id) and sid(c.get('Site_ID')) == sid(site_id)]
    cancelled_ids = {sid(c.get('Receive_ID')) for c in recv if is_cancelled(c)} - {''}
    out = []
    for r in records(book.df(damage_sheet(book))):
        if sid(r.get('Party_ID')) != sid(party_id) or sid(r.get('Site_ID')) != sid(site_id):
            continue
        status = str(r.get('Status', '')).strip().lower()
        if status in ('cancelled', 'canceled', 'billed') or sid(r.get('Billed_Bill_ID')):
            continue
        if sid(r.get('Receive_ID')) in cancelled_ids:
            continue
        dt = parse_any_date(r.get('Charge_Date'))
        if dt and dt > end:
            continue
        out.append(r)
    return out


def damage_row(charge, master):
    iid = sid(charge.get('Item_ID'))
    m = master.get(iid)
    name = item_label(m) if m else (charge.get('Item_Name') or 'Item')
    desc = str(charge.get('Description') or '').strip()
    qty = D(charge.get('Damaged_Qty'))
    rate = D(charge.get('Damage_Rate'))
    if rate == 0 and qty > 0 and to_float(charge.get('Damage_Amount')):
        rate = D(charge.get('Damage_Amount')) / qty
    return {'type': 'damage', 'item_id': iid, 'desc': f'Damage {name}' + (f' - {desc}' if desc else ''),
            'from': '', 'to': '', 'qty': num(qty), 'days': '', 'rate': num(q2(rate)),
            'charge_id': sid(charge.get('Damage_Charge_ID'))}


# ----------------------------------------------------------------------------
# Draft (auto-calculated)
# ----------------------------------------------------------------------------
def build_draft(book, party_id, site_id, start, end):
    check_period(start, end)
    party, site = _party_site(book, party_id, site_id)
    master = {sid(i.get('Item_ID')): i for i in load_items_master(book)}
    warnings = []

    events = quantity_events(book, party_id, site_id, warnings)
    rows = []
    for iid in sorted(events, key=lambda i: item_label(master.get(i, {'Item_ID': i})).lower()):
        m = master.get(iid, {})
        label = item_label(m) if m else f'Item #{iid}'
        rate = D(m.get('Daily_Rental_Rate'))
        ivs = item_intervals(events[iid], start, end)
        if any(iv['qty'] < 0 for iv in ivs):
            warnings.append(f'{label}: returned quantity is more than issued - please verify the challans.')
        shown = [iv for iv in ivs if iv['qty'] > 0]
        if shown and rate == 0:
            warnings.append(f'{label}: Daily_Rental_Rate is 0 in the Items sheet - please enter the rate.')
        for iv in shown:
            rows.append({'type': 'rental', 'item_id': iid, 'desc': label, 'from': iso(iv['from']),
                         'to': iso(iv['to']), 'qty': num(iv['qty']), 'days': (iv['to'] - iv['from']).days + 1,
                         'rate': num(rate), 'charge_id': ''})

    for ch in pending_damage_charges(book, party_id, site_id, end):
        rows.append(damage_row(ch, master))
        
    # Damage rows ko consolidate (merge) karne ke liye yahan add kiya hai
    rows = consolidate_damage_rows(rows)

    if not rows:
        warnings.append('No rental quantity or damage charge found for this period. '
                        'You can still add rows manually.')

    registered = is_gst_registered(party)
    cgst, sgst = get_gst_rates(book)
    today = date.today()
    return {
        'mode': 'new', 'bill_no': '', 'status': 'Active', 'party': party, 'site': site,
        'party_id': sid(party_id), 'site_id': sid(site_id),
        'period_start': iso(start), 'period_end': iso(end), 'bill_date': iso(today),
        'rows': rows, 'apply_gst': registered, 'gst_registered': registered,
        'cgst_rate': num(cgst), 'sgst_rate': num(sgst),
        'balance_text': balance_text(events, end, master), 'remarks': '',
        'warnings': warnings, 'overlaps': find_overlaps(book, party_id, site_id, start, end),
        'totals': compute_totals(rows, registered, cgst, sgst),
        'bill_no_preview': {'gst': next_bill_no(book, True, today),
                            'non_gst': next_bill_no(book, False, today)},
    }


# ----------------------------------------------------------------------------
# Editor form -> validated payload
# ----------------------------------------------------------------------------
def _date_field(value, label, required=True):
    s = str(value or '').strip()
    if not s:
        if required:
            raise ChallanError(f'{label} is required.')
        return None
    dt = parse_any_date(s)
    if dt is None:
        raise ChallanError(f'{label}: invalid date "{s}".')
    return dt


_ARRAYS = ('row_type[]', 'row_item_id[]', 'row_desc[]', 'row_from[]', 'row_to[]',
           'row_qty[]', 'row_days[]', 'row_rate[]', 'row_charge_id[]')


def _form_arrays(form):
    return {n: form.getlist(n) for n in _ARRAYS}


def _at(arrays, name, i):
    lst = arrays[name]
    return lst[i] if i < len(lst) else ''


def parse_rows(form, master):
    a = _form_arrays(form)
    count = len(a['row_type[]'])
    if count == 0:
        raise ChallanError('The invoice has no rows. Please add at least one row.')
    if count > MAX_ROWS:
        raise ChallanError(f'An invoice can have at most {MAX_ROWS} rows.')

    rows = []
    for i in range(count):
        n = f'Row {i + 1}: '
        rtype = str(_at(a, 'row_type[]', i)).strip().lower()
        if rtype not in ROW_TYPES:
            rtype = 'rental'
        item_id = sid(_at(a, 'row_item_id[]', i))
        desc = str(_at(a, 'row_desc[]', i)).strip()[:150]
        qty, rate = D(_at(a, 'row_qty[]', i)), D(_at(a, 'row_rate[]', i))
        if qty <= 0:
            raise ChallanError(n + 'quantity (NOS) must be greater than zero.')
        if rate < 0:
            raise ChallanError(n + 'rate cannot be negative.')
        if qty > MAX_AMOUNT or rate > MAX_AMOUNT:
            raise ChallanError(n + 'quantity or rate is too large.')

        if rtype == 'rental':
            m = master.get(item_id)
            if not item_id or m is None:
                raise ChallanError(n + 'please select an item.')
            dfrom = _date_field(_at(a, 'row_from[]', i), n + 'From date')
            dto = _date_field(_at(a, 'row_to[]', i), n + 'To date')
            if dfrom > dto:
                raise ChallanError(n + '"From" date cannot be after the "To" date.')
            rows.append({'type': 'rental', 'item_id': item_id, 'desc': desc or item_label(m) or 'Item',
                         'from': iso(dfrom), 'to': iso(dto), 'qty': num(qty),
                         'days': (dto - dfrom).days + 1,            # always computed here
                         'rate': num(rate), 'charge_id': ''})
        else:
            if not desc:
                raise ChallanError(n + 'please enter the description.')
            rows.append({'type': rtype, 'item_id': item_id, 'desc': desc, 'from': '', 'to': '',
                         'qty': num(qty), 'days': '', 'rate': num(rate),
                         'charge_id': sid(_at(a, 'row_charge_id[]', i)) if rtype == 'damage' else ''})
    return rows


def parse_payload(form, book, bill_no=None):
    """Validate the editor form. With bill_no (edit) party / site / GST flag stay as saved."""
    existing = None
    if bill_no is not None:
        existing = load_bill(book, bill_no)
        party_id, site_id = existing['party_id'], existing['site_id']
    else:
        party_id, site_id = sid(form.get('party_id')), sid(form.get('site_id'))
    party, site = _party_site(book, party_id, site_id)
    start = _date_field(form.get('period_start'), 'Period From')
    end = _date_field(form.get('period_end'), 'Period To')
    check_period(start, end)
    bill_date = _date_field(form.get('bill_date'), 'Invoice date', required=False) or date.today()

    master = {sid(i.get('Item_ID')): i for i in load_items_master(book)}
    rows = parse_rows(form, master)

    cgst, sgst = D(form.get('cgst_rate'), str(DEFAULT_GST_RATE)), D(form.get('sgst_rate'), str(DEFAULT_GST_RATE))
    for label, val in (('CGST', cgst), ('SGST', sgst)):
        if val < 0 or val > 100:
            raise ChallanError(f'{label} rate must be between 0 and 100.')

    if existing is not None:
        apply_gst = existing['apply_gst']              # number series is fixed once saved
    else:
        wants = str(form.get('apply_gst', '')).strip().lower() in ('1', 'on', 'true', 'yes')
        apply_gst = wants and is_gst_registered(party)  # GST only with a valid GSTIN
    return {
        'party_id': party_id, 'site_id': site_id, 'party': party, 'site': site,
        'period_start': start, 'period_end': end, 'bill_date': bill_date, 'rows': rows,
        'apply_gst': apply_gst, 'cgst_rate': cgst, 'sgst_rate': sgst,
        'balance_text': str(form.get('balance_text') or '').strip()[:2000],
        'remarks': str(form.get('remarks') or '').strip()[:300],
    }


def draft_from_form(book, form, bill_no=None):
    """Draft dict rebuilt from the posted form (never raises) - re-shows the editor after an error."""
    def lenient(v, default=0):
        try:
            s = str(v).strip().replace(',', '')
            return num(Decimal(s)) if s else default
        except (InvalidOperation, ValueError):
            return default

    saved = None
    if bill_no is not None:
        try:
            saved = load_bill(book, bill_no)
        except Exception:
            saved = None
    party_id = saved['party_id'] if saved else sid(form.get('party_id'))
    site_id = saved['site_id'] if saved else sid(form.get('site_id'))
    party = next((p for p in load_parties(book) if sid(p.get('Party_ID')) == party_id), {})
    site = next((s for s in load_sites(book) if sid(s.get('Site_ID')) == site_id), {})

    a = _form_arrays(form)
    rows = []
    for i in range(len(a['row_type[]'])):
        rtype = str(_at(a, 'row_type[]', i)).strip().lower()
        if rtype not in ROW_TYPES:
            rtype = 'rental'
        d1, d2 = parse_any_date(_at(a, 'row_from[]', i)), parse_any_date(_at(a, 'row_to[]', i))
        rows.append({
            'type': rtype, 'item_id': sid(_at(a, 'row_item_id[]', i)), 'desc': str(_at(a, 'row_desc[]', i)),
            'from': iso(d1), 'to': iso(d2), 'qty': lenient(_at(a, 'row_qty[]', i)),
            'days': ((d2 - d1).days + 1 if d1 and d2 and d2 >= d1 else 0) if rtype == 'rental' else '',
            'rate': lenient(_at(a, 'row_rate[]', i)), 'charge_id': sid(_at(a, 'row_charge_id[]', i)),
        })

    registered = is_gst_registered(party) if party else False
    if saved:
        apply_gst = saved['apply_gst']
    else:
        apply_gst = registered and str(form.get('apply_gst', '')).strip().lower() in ('1', 'on', 'true', 'yes')
    cgst, sgst = lenient(form.get('cgst_rate'), 9), lenient(form.get('sgst_rate'), 9)
    try:
        totals = compute_totals(rows, apply_gst, cgst, sgst)
    except Exception:
        totals = {'rows': [], 'taxable': 0, 'cgst': 0, 'sgst': 0, 'total': 0, 'round_off': 0, 'net': 0}

    today = date.today()
    return {
        'mode': 'edit' if saved else 'new', 'bill_no': saved['bill_no'] if saved else '',
        'status': 'Active', 'party': party, 'site': site, 'party_id': party_id, 'site_id': site_id,
        'period_start': str(form.get('period_start') or ''), 'period_end': str(form.get('period_end') or ''),
        'bill_date': str(form.get('bill_date') or ''), 'rows': rows, 'apply_gst': apply_gst,
        'gst_registered': registered, 'cgst_rate': cgst, 'sgst_rate': sgst,
        'balance_text': str(form.get('balance_text') or ''), 'remarks': str(form.get('remarks') or ''),
        'warnings': [], 'overlaps': [], 'totals': totals,
        'bill_no_preview': {'gst': next_bill_no(book, True, today), 'non_gst': next_bill_no(book, False, today)},
    }


# ----------------------------------------------------------------------------
# Saving
# ----------------------------------------------------------------------------
def _s(v):
    """sid() for raw DataFrame cells (blank Excel cell = NaN -> '')."""
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return ''
    return sid(v)


def _sync_damage_charges(ddf, bill_id, keep_ids):
    """Damage charges of this bill that are no longer on it become Pending again;
    the ones on it become Billed."""
    if ddf.empty or 'Damage_Charge_ID' not in ddf.columns:
        return ddf
    for col in ('Billed_Bill_ID', 'Status'):
        if col not in ddf.columns:
            ddf[col] = ''
        ddf[col] = ddf[col].astype(object)

    ids = ddf['Damage_Charge_ID'].apply(_s)
    keep = [sid(k) for k in keep_ids]
    mine = (ddf['Billed_Bill_ID'].apply(_s) == sid(bill_id)) & (~ids.isin(keep))
    ddf.loc[mine, 'Billed_Bill_ID'] = ''
    ddf.loc[mine, 'Status'] = 'Pending'
    for cid in keep:
        match = ddf.index[ids == cid]
        if len(match) > 0:
            i = match[0]
            ddf.at[i, 'Billed_Bill_ID'] = norm_id(bill_id)
            ddf.at[i, 'Status'] = 'Billed'
    return ddf


def _check_charges_free(ddf, charge_ids, bill_id):
    """A damage charge can only sit on one active invoice."""
    if ddf.empty or 'Damage_Charge_ID' not in ddf.columns or not charge_ids:
        return
    ids = ddf['Damage_Charge_ID'].apply(_s)
    for cid in charge_ids:
        match = ddf.index[ids == sid(cid)]
        if len(match) == 0:
            raise ChallanError(f'Damage charge #{cid} no longer exists. Please remove that row.')
        row = ddf.loc[match[0]]
        billed_to = _s(row.get('Billed_Bill_ID')) if 'Billed_Bill_ID' in ddf.columns else ''
        if billed_to and billed_to != sid(bill_id):
            raise ChallanError(f'Damage charge #{cid} is already billed on another invoice.')


def _line_rows(rows, bill_id, start_line_id):
    out, n = [], start_line_id
    for r in rows:
        number, amount = row_amounts(r)
        out.append({
            'Bill_Line_ID': n, 'Bill_ID': bill_id,
            'Item_ID': norm_id(r['item_id']) if r['item_id'] else '',
            'Item_Name_Snapshot': r['desc'],
            'Date_From': fmt_date(parse_any_date(r['from'])) if r['from'] else '',
            'Date_To': fmt_date(parse_any_date(r['to'])) if r['to'] else '',
            'Quantity': num(D(r['qty'])),
            'Rental_Days': int(D(r['days'])) if r['type'] == 'rental' else '',
            'Daily_Rate': num(D(r['rate'])),
            'Line_Amount': float(amount),
            'Line_Type': LINE_TYPE[r['type']],
            'Damage_Charge_ID': norm_id(r['charge_id']) if r.get('charge_id') else '',
        })
        n += 1
    return out


def _bill_row(payload, totals, bill_id, bill_no, status='Active'):
    gst = payload['apply_gst']
    return {
        'Bill_ID': bill_id, 'Bill_No': bill_no,
        'Party_ID': norm_id(payload['party_id']), 'Site_ID': norm_id(payload['site_id']),
        'Period_Start': fmt_date(payload['period_start']), 'Period_End': fmt_date(payload['period_end']),
        'Bill_Date': fmt_date(payload['bill_date']),
        'Taxable_Value': totals['taxable'],
        'CGST_Rate': num(payload['cgst_rate']) if gst else 0, 'CGST_Amount': totals['cgst'],
        'SGST_Rate': num(payload['sgst_rate']) if gst else 0, 'SGST_Amount': totals['sgst'],
        'Total_Amount': totals['total'], 'Round_Off': totals['round_off'], 'Net_Total': totals['net'],
        'Status': status,
    }


def _locate(bdf, bill_no):
    if bdf is None or bdf.empty or 'Bill_No' not in bdf.columns:
        raise ChallanError('Invoice not found.')
    mask = bdf['Bill_No'].astype(str).str.strip() == str(bill_no).strip()
    if not mask.any():
        raise ChallanError('Invoice not found.')
    return bdf.index[mask][0]


def _set(bdf, idx, col, value):
    if col not in bdf.columns:
        bdf[col] = ''
    bdf[col] = bdf[col].astype(object)
    bdf.at[idx, col] = value


def _is_cancelled_value(v):
    return str(v).strip().lower() in ('cancelled', 'canceled')


def create_bill(book, payload):
    b_name, l_name, d_name = bills_sheet(book), lines_sheet(book), damage_sheet(book)
    bdf, ldf = book.df(b_name), book.df(l_name)

    overlaps = find_overlaps(book, payload['party_id'], payload['site_id'],
                             payload['period_start'], payload['period_end'])
    if overlaps:
        nos = ', '.join(f"{o['bill_no']} ({o['from']} to {o['to']})" for o in overlaps)
        raise ChallanError(f'An invoice already exists for this party, site and period: {nos}. '
                           f'Cancel it first or change the period.')

    charge_ids = [r['charge_id'] for r in payload['rows'] if r.get('charge_id')]
    ddf = book.df(d_name) if d_name else None
    bill_id = next_id(bdf, 'Bill_ID')
    if charge_ids and ddf is not None:
        _check_charges_free(ddf, charge_ids, bill_id)

    totals = compute_totals(payload['rows'], payload['apply_gst'], payload['cgst_rate'], payload['sgst_rate'])
    bill_no = next_bill_no(book, payload['apply_gst'], payload['bill_date'])

    row = _bill_row(payload, totals, bill_id, bill_no)
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    row.update({'Created_At': now, 'Created_By': current_user(), 'Cancel_Reason': '',
                'Balance_Text': payload['balance_text'], 'Apply_GST': 'Yes' if payload['apply_gst'] else 'No',
                'Remarks': payload['remarks']})
    lines = _line_rows(payload['rows'], bill_id, next_id(ldf, 'Bill_Line_ID'))

    audit_name, audit_df = audit_update(book, 'CREATE', bill_no, after=row, entity_type='Monthly_Bill')
    updates = {b_name: append_rows(bdf, [row]), l_name: append_rows(ldf, lines), audit_name: audit_df}
    if charge_ids and ddf is not None:
        updates[d_name] = _sync_damage_charges(ddf, bill_id, charge_ids)
    return updates, bill_no


def update_bill(book, bill_no, payload):
    """Edit an existing invoice in place (same Bill_No; lines replaced; damage charges re-linked)."""
    b_name, l_name, d_name = bills_sheet(book), lines_sheet(book), damage_sheet(book)
    bdf, ldf = book.df(b_name), book.df(l_name)

    idx = _locate(bdf, bill_no)
    if _is_cancelled_value(bdf.at[idx, 'Status']):
        raise ChallanError('A cancelled invoice cannot be edited.')
    bill_id = bdf.at[idx, 'Bill_ID']
    before = records(bdf.loc[[idx]])[0]

    overlaps = find_overlaps(book, payload['party_id'], payload['site_id'],
                             payload['period_start'], payload['period_end'], exclude_bill_id=bill_id)
    if overlaps:
        nos = ', '.join(f"{o['bill_no']} ({o['from']} to {o['to']})" for o in overlaps)
        raise ChallanError(f'Another invoice already exists for this party, site and period: {nos}.')

    charge_ids = [r['charge_id'] for r in payload['rows'] if r.get('charge_id')]
    ddf = book.df(d_name) if d_name else None
    if charge_ids and ddf is not None:
        _check_charges_free(ddf, charge_ids, bill_id)

    totals = compute_totals(payload['rows'], payload['apply_gst'], payload['cgst_rate'], payload['sgst_rate'])
    head = _bill_row(payload, totals, bill_id, str(bill_no))
    head.update({'Balance_Text': payload['balance_text'], 'Apply_GST': 'Yes' if payload['apply_gst'] else 'No',
                 'Remarks': payload['remarks'], 'Updated_At': datetime.now().strftime('%Y-%m-%d %H:%M:%S')})
    for col, val in head.items():
        if col in ('Bill_ID', 'Bill_No', 'Status'):
            continue
        _set(bdf, idx, col, val)

    next_line = next_id(ldf, 'Bill_Line_ID')             # line IDs are never reused
    if not ldf.empty and 'Bill_ID' in ldf.columns:
        ldf = ldf[ldf['Bill_ID'].apply(_s) != sid(bill_id)]
    ldf = append_rows(ldf, _line_rows(payload['rows'], bill_id, next_line))

    audit_name, audit_df = audit_update(book, 'UPDATE', str(bill_no), before=before, after=head,
                                        entity_type='Monthly_Bill')
    updates = {b_name: bdf, l_name: ldf, audit_name: audit_df}
    if ddf is not None:
        updates[d_name] = _sync_damage_charges(ddf, bill_id, charge_ids)
    return updates


def cancel_bill(book, bill_no, reason):
    reason = str(reason or '').strip()
    if len(reason) < 3:
        raise ChallanError('Please enter a reason (at least 3 characters) for cancelling the invoice.')
    b_name, d_name = bills_sheet(book), damage_sheet(book)
    bdf = book.df(b_name)
    idx = _locate(bdf, bill_no)
    if _is_cancelled_value(bdf.at[idx, 'Status']):
        raise ChallanError('This invoice is already cancelled.')
    bill_id = bdf.at[idx, 'Bill_ID']
    before = records(bdf.loc[[idx]])[0]

    _set(bdf, idx, 'Status', 'Cancelled')
    _set(bdf, idx, 'Cancel_Reason', reason[:300])
    _set(bdf, idx, 'Cancelled_At', datetime.now().strftime('%Y-%m-%d %H:%M:%S'))

    audit_name, audit_df = audit_update(book, 'CANCEL', str(bill_no), before=before, reason=reason[:300],
                                        entity_type='Monthly_Bill')
    updates = {b_name: bdf, audit_name: audit_df}
    if d_name:                                           # its damage charges become billable again
        updates[d_name] = _sync_damage_charges(book.df(d_name), bill_id, [])
    return updates


# ----------------------------------------------------------------------------
# Load a saved invoice / list / print
# ----------------------------------------------------------------------------
def load_bill(book, bill_no):
    bdf = book.df(bills_sheet(book))
    idx = _locate(bdf, bill_no)
    head = records(bdf.loc[[idx]])[0]
    bill_id = sid(head.get('Bill_ID'))
    party = next((p for p in load_parties(book) if sid(p.get('Party_ID')) == sid(head.get('Party_ID'))), {})
    site = next((s for s in load_sites(book) if sid(s.get('Site_ID')) == sid(head.get('Site_ID'))), {})

    rows = []
    for r in records(book.df(lines_sheet(book))):
        if sid(r.get('Bill_ID')) != bill_id:
            continue
        rtype = LINE_TYPE_REV.get(str(r.get('Line_Type', '')).strip().lower(), 'extra')
        qty = D(r.get('Quantity'))
        rows.append({
            'type': rtype, 'item_id': sid(r.get('Item_ID')), 'desc': r.get('Item_Name_Snapshot') or '',
            'from': iso(parse_any_date(r.get('Date_From'))), 'to': iso(parse_any_date(r.get('Date_To'))),
            'qty': num(qty), 'days': int(D(r.get('Rental_Days'))) if rtype == 'rental' else '',
            'rate': num(D(r.get('Daily_Rate'))), 'charge_id': sid(r.get('Damage_Charge_ID')),
            '_line_id': to_float(r.get('Bill_Line_ID')),
        })
    rows.sort(key=lambda x: x['_line_id'])
    for r in rows:
        r.pop('_line_id', None)
        
    # Saved invoice load hote waqt bhi damage rows consolidate karne ke liye yahan add kiya hai
    rows = consolidate_damage_rows(rows)

    apply_gst = _bill_is_gst(head)
    cgst_default, sgst_default = get_gst_rates(book)
    cgst = D(head.get('CGST_Rate')) if apply_gst else cgst_default
    sgst = D(head.get('SGST_Rate')) if apply_gst else sgst_default
    totals = compute_totals(rows, apply_gst, cgst, sgst)
    return {
        'mode': 'edit', 'bill_no': str(head.get('Bill_No')), 'status': head.get('Status') or 'Active',
        'cancel_reason': head.get('Cancel_Reason') or '', 'party': party, 'site': site,
        'party_id': sid(head.get('Party_ID')), 'site_id': sid(head.get('Site_ID')),
        'period_start': iso(parse_any_date(head.get('Period_Start'))),
        'period_end': iso(parse_any_date(head.get('Period_End'))),
        'bill_date': iso(parse_any_date(head.get('Bill_Date'))), 'rows': rows, 'apply_gst': apply_gst,
        'gst_registered': is_gst_registered(party) if party else False,
        'cgst_rate': num(cgst), 'sgst_rate': num(sgst),
        'balance_text': head.get('Balance_Text') or head.get('Balance_Qty_Text') or '',
        'remarks': head.get('Remarks') or '',
        'warnings': [], 'overlaps': [], 'totals': totals, 'head': head,
        'bill_no_preview': {},
    }


def list_bills(book):
    """Rows for the index page (newest first) + summary cards."""
    names = {sid(p.get('Party_ID')): p.get('Party_Name') or p.get('Party_ID') for p in load_parties(book)}
    sites = {sid(s.get('Site_ID')): s.get('Site_Name') or s.get('Site_Address') or s.get('Site_ID')
             for s in load_sites(book)}
    bills = records(book.df(bills_sheet(book)))
    active = cancelled = 0
    billed = Decimal('0')
    for b in bills:
        b['Party_Name'] = names.get(sid(b.get('Party_ID')), b.get('Party_ID', ''))
        b['Site_Name'] = sites.get(sid(b.get('Site_ID')), b.get('Site_ID', ''))
        b['Cancelled'] = is_cancelled(b)
        b['Has_GST'] = _bill_is_gst(b)
        b['Net_Display'] = inr(b.get('Net_Total'))
        if b['Cancelled']:
            cancelled += 1
        else:
            active += 1
            billed += Decimal(str(to_float(b.get('Net_Total'))))
    bills.reverse()
    return bills, {'active': active, 'cancelled': cancelled, 'billed': inr(billed)}


# ----------------------------------------------------------------------------
# Print pagination  (A4 pages are filled completely before the next page starts)
# Row height is fixed in print.html (one unit = one single-line row).
# PRINT_CAP_MID  = units a NON-last page can hold  (page-total box at the bottom)
# PRINT_CAP_LAST = units the LAST page can hold    (full summary + terms + signatures)
# ----------------------------------------------------------------------------
PRINT_CAP_MID = 26      # units a non-last page can hold (measured)
PRINT_CAP_LAST = 16     # units the last page can hold (measured)
PRINT_LONG_NAME = 26          # item names longer than this wrap to 2 lines (= 2 units)


def _row_units(name, shows_name):
    return 2 if (shows_name and len(name) > PRINT_LONG_NAME) else 1


def _qn(x):
    d = Decimal(str(x or 0))
    return f'{int(d)}' if d == d.to_integral_value() else f'{float(d):.2f}'.rstrip('0').rstrip('.')


def paginate_rows(items, cap_mid=None, cap_last=None):
    """items: list of dicts {group, desc, type, qty, amount(Decimal), ...}.  Returns list of pages (lists of items)."""
    cap_mid = cap_mid or PRINT_CAP_MID
    cap_last = cap_last or PRINT_CAP_LAST

    def cost_of(chunk, first_page):
        """Units used by this chunk when placed on a page (continuation rows show the name again)."""
        total, prev = (0 if first_page else 1), None          # +1 = "brought forward" row
        for k, it in enumerate(chunk):
            shows = k == 0 or it['group'] != prev
            nm = it['desc'] + (' (contd.)' if k == 0 and it.get('cont') else '')
            total += _row_units(nm, shows)
            prev = it['group']
        return total

    pages, i, n = [], 0, len(items)
    while i < n:
        first = not pages
        rest = items[i:]
        if cost_of(rest, first) <= cap_last:
            pages.append(rest)
            break
        chunk = []
        for it in rest:
            trial = chunk + [it]
            if cost_of(trial, first) > cap_mid:
                break
            chunk = trial
        if not chunk:                                          # safety: always move forward
            chunk = rest[:1]
        if len(chunk) == len(rest):                            # leave at least one row for the last page
            chunk = chunk[:-1] or chunk
        pages.append(chunk)
        i += len(chunk)
        if i < n:
            items[i]['cont'] = items[i]['group'] == items[i - 1]['group']
    return pages


def build_print_pages(rows_raw):
    """rows_raw: list of (row, calc) -> list of ready-to-render pages with running totals."""
    items, prev_group, gid, prev_type, prev_desc, prev_qty = [], None, 0, None, None, None
    sno = 0
    for r, calc in rows_raw:
        rental = r['type'] == 'rental'
        same = rental and prev_type == 'rental' and prev_desc == r['desc']
        if not same:
            gid += 1
            sno += 1
        qty = Decimal(str(r['qty'] or 0))
        delta = ''
        if same and prev_qty is not None:
            diff = qty - prev_qty
            if diff < 0:
                delta = f'({_qn(prev_qty)} - {_qn(-diff)}) \u2192 {_qn(qty)}'
            elif diff > 0:
                delta = f'({_qn(prev_qty)} + {_qn(diff)}) \u2192 {_qn(qty)}'
            else:
                delta = f'({_qn(qty)})'
        items.append({
            'group': gid, 'sno': sno, 'new_group': not same, 'desc': r['desc'], 'type': r['type'],
            'qty': qty_str(r['qty']), 'from': fmt_date(r['from']) if rental else '',
            'to': fmt_date(r['to']) if rental else '', 'days': r['days'] if rental else '',
            'number': qty_str(calc['number']) if rental else '', 'rate': inr(r['rate']),
            'amount_d': q2(Decimal(str(calc['amount'] or 0))), 'amount': inr(calc['amount']),
            'delta': delta, 'cont': False,
        })
        prev_type, prev_desc, prev_qty = r['type'], r['desc'], qty

    chunks = paginate_rows(items)
    total_pages, pages, cum = len(chunks) or 1, [], Decimal('0')
    for pi, chunk in enumerate(chunks or [[]]):
        used, run_start = (0 if pi == 0 else 1), None
        out = []
        for k, it in enumerate(chunk):
            first_on_page = k == 0
            shows_name = first_on_page or it['new_group']
            cont = first_on_page and not it['new_group']
            row = dict(it)
            row.update({'shows_name': shows_name, 'contd': cont, 'run_start': shows_name,
                        'span': 1, 'units': _row_units(it['desc'] + (' (contd.)' if cont else ''), shows_name)})
            used += row['units']
            out.append(row)
        # S.No cell spans every row of the same item on this page
        for k, row in enumerate(out):
            if row['run_start']:
                span, j = 1, k + 1
                while j < len(out) and out[j]['group'] == row['group']:
                    span += 1
                    j += 1
                row['span'] = span
            nxt = out[k + 1] if k + 1 < len(out) else None
            row['group_end'] = nxt is None or nxt['group'] != row['group']
        page_total = sum((r['amount_d'] for r in out), Decimal('0'))
        before = cum
        cum += page_total
        last = pi == total_pages - 1
        pages.append({
            'no': pi + 1, 'of': total_pages, 'last': last, 'rows': out, 'used': used,
            'capacity': PRINT_CAP_LAST if last else PRINT_CAP_MID,
            'blank': max(0, (PRINT_CAP_LAST if last else PRINT_CAP_MID) - used),
            'bf': inr(before) if pi > 0 else '',
            'page_total': inr(page_total), 'cf': inr(cum),
            'page_total_words': number_to_words(q2(page_total)).upper() if not last else '',
        })
    return pages


def print_context(book, bill_no):
    """Everything print.html needs, already formatted and split into A4 pages."""
    d = load_bill(book, bill_no)
    t = d['totals']
    rows = []
    for r, calc in zip(d['rows'], t['rows']):
        rental = r['type'] == 'rental'
        rows.append({
            'desc': r['desc'], 'type': r['type'],
            'from': fmt_date(r['from']) if rental else '', 'to': fmt_date(r['to']) if rental else '',
            'qty': qty_str(r['qty']), 'qty_num': float(r['qty'] or 0), 'days': r['days'] if rental else '',
            'number': qty_str(calc['number']) if rental else '',
            'rate': inr(r['rate']), 'amount': inr(calc['amount']),
        })
    pages = build_print_pages(list(zip(d['rows'], t['rows'])))
    party, site, gst = d['party'], d['site'], d['apply_gst']
    return {
        'bill_no': d['bill_no'], 'status': d['status'], 'cancel_reason': d['cancel_reason'],
        'bill_date': fmt_date(d['bill_date']), 'period_start': fmt_date(d['period_start']),
        'period_end': fmt_date(d['period_end']), 'party': party, 'site': site,
        'party_gstin': clean_gstin(party.get('GSTIN')) if is_valid_gstin(party.get('GSTIN')) else '',
        'rows': rows, 'pages': pages, 'apply_gst': gst, 'title': 'TAX INVOICE' if gst else 'INVOICE',
        'cgst_rate': qty_str(d['cgst_rate']), 'sgst_rate': qty_str(d['sgst_rate']),
        'taxable': inr(t['taxable']), 'cgst': inr(t['cgst']), 'sgst': inr(t['sgst']),
        'total': inr(t['total']), 'round_off': ('+' if t['round_off'] >= 0 else '-') + inr(abs(t['round_off'])),
        'net': inr(t['net']), 'net_words': 'Rupees ' + number_to_words(t['net']),
        'balance_text': d['balance_text'], 'remarks': d['remarks'],
        'site_full': site.get('Site_Name') or site.get('Site_Address') or '-',
        'party_address': site.get('Site_Address') or party.get('Address') or '-',
        'state_code': party.get('State_Code') or '',
    }


def delete_bill_permanently(book, bill_no):
    b_name, l_name, d_name = bills_sheet(book), lines_sheet(book), damage_sheet(book)
    bdf = book.df(b_name)
    idx = _locate(bdf, bill_no)
    bill_id = bdf.at[idx, 'Bill_ID']
    
    # DataFrame se row uda do
    bdf = bdf.drop(idx)
    
    # Associated lines bhi hata do
    ldf = book.df(l_name)
    if not ldf.empty and 'Bill_ID' in ldf.columns:
        ldf = ldf[ldf['Bill_ID'].apply(_s) != sid(bill_id)]
        
    return {b_name: bdf, l_name: ldf}