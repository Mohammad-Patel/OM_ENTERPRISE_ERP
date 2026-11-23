"""
Receive Challan module - OM Enterprise ERP
------------------------------------------
Goods coming back from a party.  No issue challan has to be selected: the form shows
the STOCK LYING WITH THE PARTY at the chosen site and the receive happens from that stock.

    stock with party = opening stock + issued - already returned     (see challan_store.stock_summary)

Rules
    * received qty + breakage qty (stock items) of an item can never exceed the stock with the party
    * Breakage rows can be a stock item OR a manual item (typed by hand, no stock effect)
    * breakage of a stock item is returned damaged -> it also reduces the party's stock
    * a receive challan whose damage charges are already billed cannot be edited / deleted

Sheets written (one atomic save):
    Receive_Challan, Receive_Challan_Items, Damage_Charges, Audit_Log
"""
import logging
import math
from datetime import datetime

from flask import (Blueprint, flash, jsonify, redirect, render_template, request,
                   url_for)

try:
    from .challan_store import (  # noqa: F401
        ALLOWED_PER_PAGE, DAMAGE_SHEET, DATE_FMT, DEFAULT_UNIT_WEIGHT, RECEIVE_ITEMS_SHEET,
        RECEIVE_SHEET, Book, ChallanError, append_rows, audit_update, damage_sheet,
        excel_lock, fmt_money, fmt_qty, is_active, is_cancelled, item_label,
        load_items_master, load_parties, load_sites, next_id, next_number, norm_id,
        normalize_date, number_to_words, parse_any_date, receive_items_sheet,
        receive_sheet, records, safe_action, save_sheets, sid, stock_summary,
        to_float, to_int, to_js_json, current_user)
except ImportError:  # pragma: no cover
    from app.routes.challan_store import (  # noqa: F401
        ALLOWED_PER_PAGE, DAMAGE_SHEET, DATE_FMT, DEFAULT_UNIT_WEIGHT, RECEIVE_ITEMS_SHEET,
        RECEIVE_SHEET, Book, ChallanError, append_rows, audit_update, damage_sheet,
        excel_lock, fmt_money, fmt_qty, is_active, is_cancelled, item_label,
        load_items_master, load_parties, load_sites, next_id, next_number, norm_id,
        normalize_date, number_to_words, parse_any_date, receive_items_sheet,
        receive_sheet, records, safe_action, save_sheets, sid, stock_summary,
        to_float, to_int, to_js_json, current_user)

logger = logging.getLogger(__name__)

receive_challan_bp = Blueprint('receive_challan', __name__, url_prefix='/receive-challan')

RECEIVE_PREFIX = 'OM/RC/'
MANUAL = '__manual__'          # value of the breakage item dropdown for a hand-typed item
ENDPOINT_INDEX = 'receive_challan.index'


# GST rule + numbering live in ONE place (app/services/numbering_service.py)
from app.services.numbering_service import is_gst_registered, next_receive_challan_no  # noqa: E402



# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------
def _num(v):
    """5.0 -> 5 (for JSON)."""
    f = to_float(v)
    return int(f) if f == int(f) else round(f, 2)


def _whole(value, label):
    """Quantity must be a whole number >= 0 ('' -> 0)."""
    s = str(value if value is not None else '').strip()
    if s == '':
        return 0
    try:
        f = float(s)
    except ValueError:
        raise ChallanError(f'{label}: please enter a valid number.')
    if f < 0:
        raise ChallanError(f'{label}: quantity cannot be negative.')
    if f != int(f):
        raise ChallanError(f'{label}: quantity must be a whole number.')
    return int(f)


def _money(value, label):
    s = str(value if value is not None else '').strip()
    if s == '':
        raise ChallanError(f'{label}: please enter the rate (0 if no charge).')
    try:
        f = float(s)
    except ValueError:
        raise ChallanError(f'{label}: invalid rate.')
    if f < 0:
        raise ChallanError(f'{label}: rate cannot be negative.')
    return round(f, 2)


def _master_label(master, iid):
    return item_label(master.get(sid(iid), {'Item_ID': iid}))


def is_billed(charge):
    return sid(charge.get('Billed_Bill_ID')) != '' or str(charge.get('Status', '')).strip().lower() == 'billed'


def parse_header(form, parties, sites):
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
        'Receive_Date': normalize_date(form.get('date')),
        'Vehicle_No': (form.get('vehicle_no') or '').strip().upper(),
        'Driver_Name': (form.get('driver_name') or '').strip(),
        'Driver_Mob': (form.get('driver_mob') or '').strip(),
        'Remarks': (form.get('remarks') or '').strip(),
    }


def parse_received(form, master):
    """{item_id: qty} of the items being received (qty > 0)."""
    received = {}
    for iid, qty in zip(form.getlist('recv_item_id[]'), form.getlist('recv_qty[]')):
        iid = sid(iid)
        if not iid:
            continue
        label = _master_label(master, iid)
        q = _whole(qty, label)
        if q == 0:
            continue
        if iid in received:
            raise ChallanError(f'{label} is listed more than once.')
        received[iid] = q
    return received


def parse_breakage(form, master):
    """List of breakage rows. Item_ID is '' for a manual (hand typed) item."""
    items = form.getlist('breakage_item[]')
    names = form.getlist('breakage_name[]')
    qtys = form.getlist('breakage_qty[]')
    rates = form.getlist('breakage_rate[]')
    descs = form.getlist('breakage_desc[]')

    def at(lst, i):
        return lst[i] if i < len(lst) else ''

    rows = []
    for i in range(len(items)):
        item, name = sid(at(items, i)), (at(names, i) or '').strip()
        qty_raw, rate_raw, desc = at(qtys, i), at(rates, i), (at(descs, i) or '').strip()
        if not item and not name and not str(qty_raw).strip() and not str(rate_raw).strip() and not desc:
            continue                                          # completely blank row

        if not item:
            raise ChallanError('Breakage: please select an item (or choose "Other / Manual item").')
        if item == MANUAL:
            if not name:
                raise ChallanError('Breakage: please type the item name for the manual entry.')
            iid, label = '', name
        else:
            if item not in master:
                raise ChallanError(f'Breakage: invalid item selected (Item ID {item}).')
            iid, label = item, _master_label(master, item)

        qty = _whole(qty_raw, f'Breakage - {label}')
        if qty <= 0:
            raise ChallanError(f'Breakage - {label}: quantity is required.')
        rate = _money(rate_raw, f'Breakage - {label}')
        rows.append({'Item_ID': iid, 'Item_Name': label, 'Qty': qty, 'Rate': rate,
                     'Amount': round(qty * rate, 2), 'Description': desc})
    return rows


def validate_against_stock(received, breakage, stock, master):
    """received + breakage (stock items) of an item must stay within the stock with the party."""
    need = {}
    for iid, q in received.items():
        need.setdefault(iid, {'recv': 0, 'dmg': 0})['recv'] += q
    for b in breakage:
        if b['Item_ID']:
            need.setdefault(sid(b['Item_ID']), {'recv': 0, 'dmg': 0})['dmg'] += b['Qty']

    for iid, n in need.items():
        label = _master_label(master, iid)
        available = stock.get(iid, {}).get('outstanding', 0)
        total = n['recv'] + n['dmg']
        if available <= 0:
            raise ChallanError(f'{label}: no stock of this item is lying with this party at this site.')
        if total > available + 1e-9:
            raise ChallanError(
                f'{label}: only {fmt_qty(available)} lying with the party at this site, '
                f'but {fmt_qty(total)} entered (received {n["recv"]} + breakage {n["dmg"]}).')


def build_item_lines(received, breakage):
    """One Receive_Challan_Items row per stock item (good + damaged together)."""
    order = list(received.keys())
    for b in breakage:
        iid = sid(b['Item_ID'])
        if iid and iid not in order:
            order.append(iid)

    lines = []
    for iid in order:
        accepted = received.get(iid, 0)
        dmg_rows = [b for b in breakage if sid(b['Item_ID']) == iid]
        dmg_qty = sum(b['Qty'] for b in dmg_rows)
        dmg_amt = round(sum(b['Amount'] for b in dmg_rows), 2)
        lines.append({
            'Item_ID': norm_id(iid),
            'Returned_Qty': accepted + dmg_qty,
            'Damaged_Qty': dmg_qty,
            'Accepted_Qty': accepted,
            'Damage_Rate': round(dmg_amt / dmg_qty, 2) if dmg_qty else 0,
            'Damage_Amount': dmg_amt,
            'Damage_Status': 'Pending' if dmg_qty else '',
        })
    return lines


def number_rows(rows, id_col, start, extra):
    out, n = [], start
    for r in rows:
        row = {id_col: n}
        row.update(extra)
        row.update(r)
        out.append(row)
        n += 1
    return out


def damage_rows(breakage, start_id, receive_id, header):
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    rows, n = [], start_id
    for b in breakage:
        rows.append({
            'Damage_Charge_ID': n,
            'Receive_ID': receive_id,
            'Party_ID': header['Party_ID'],
            'Site_ID': header['Site_ID'],
            'Item_ID': norm_id(b['Item_ID']) if b['Item_ID'] else '',
            'Damaged_Qty': b['Qty'],
            'Damage_Rate': b['Rate'],
            'Damage_Amount': b['Amount'],
            'Charge_Date': header['Receive_Date'],
            'Status': 'Pending',
            'Billed_Bill_ID': '',
            'Created_At': now,
            'Created_By': current_user(),
            'Item_Name': b['Item_Name'],
            'Description': b['Description'],
        })
        n += 1
    return rows


def find_receive(book, receive_no):
    """-> (sheet_name, dataframe, row_index). Raises ChallanError if not found."""
    r_sheet = receive_sheet(book)
    df = book.df(r_sheet)
    if df.empty or 'Receive_Challan_No' not in df.columns:
        raise ChallanError('Receive challan not found.')
    mask = df['Receive_Challan_No'].astype(str) == str(receive_no)
    if not mask.any():
        raise ChallanError('Receive challan not found.')
    return r_sheet, df, df.index[mask][0]


def ensure_columns(df, cols):
    for col in cols:
        if col not in df.columns:
            df[col] = ''
        df[col] = df[col].astype(object)


# --------------------------------------------------------------------------
# Routes
# --------------------------------------------------------------------------
@receive_challan_bp.route('/')
def index():
    search_query = request.args.get('search', '').strip()
    party_filter = request.args.get('party_id', '').strip()
    gst_filter = request.args.get('gst_filter', '').strip()  # <-- NEW: 'gst' or 'nogst'
    site_filter = request.args.get('site_id', '').strip()
    date_from = parse_any_date(request.args.get('date_from'))
    date_to = parse_any_date(request.args.get('date_to'))
    per_page = request.args.get('per_page', '10')
    if per_page not in ALLOWED_PER_PAGE:
        per_page = '10'
    page = max(to_int(request.args.get('page'), 1), 1)

    all_challans, parties, sites = [], [], []
    receive_data = {}
    next_no = ''

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

        party_names = {sid(p.get('Party_ID')): p.get('Party_Name') or p.get('Party_ID') for p in all_parties}
        allowed_party_ids = {sid(p.get('Party_ID')) for p in parties}

        party_phones = {sid(p.get('Party_ID')): p.get('Phone') or p.get('Mobile_No') or p.get('Mobile') or ''
                        for p in all_parties}
        site_names = {sid(s.get('Site_ID')): s.get('Site_Name') or s.get('Site_Address') or s.get('Site_ID')
                      for s in all_sites}

        all_challans = records(book.df(receive_sheet(book)))
        id_to_no = {sid(c.get('Receive_ID')): str(c.get('Receive_Challan_No')) for c in all_challans}
        for c in all_challans:
            c['Party_Name'] = party_names.get(sid(c.get('Party_ID')), c.get('Party_ID', ''))
            c['Site_Name'] = site_names.get(sid(c.get('Site_ID')), c.get('Site_ID', ''))
            c['Party_Mobile'] = party_phones.get(sid(c.get('Party_ID'))) or '-'
            c['Received_Qty'] = 0
            c['Breakage_Qty'] = 0
            c['Breakage_Amount'] = 0

        by_no = {str(c.get('Receive_Challan_No')): c for c in all_challans}
        for r in records(book.df(receive_items_sheet(book))):
            no = id_to_no.get(sid(r.get('Receive_ID')))
            if not no:
                continue
            by_no[no]['Received_Qty'] += to_float(r.get('Accepted_Qty'))
            receive_data.setdefault(no, {'items': [], 'breakage': []})['items'].append(
                {'item_id': sid(r.get('Item_ID')), 'qty': _num(r.get('Accepted_Qty'))})
        for r in records(book.df(damage_sheet(book))):
            no = id_to_no.get(sid(r.get('Receive_ID')))
            if not no:
                continue
            by_no[no]['Breakage_Qty'] += to_float(r.get('Damaged_Qty'))
            by_no[no]['Breakage_Amount'] += to_float(r.get('Damage_Amount'))
            receive_data.setdefault(no, {'items': [], 'breakage': []})['breakage'].append({
                'item_id': sid(r.get('Item_ID')),
                'name': r.get('Item_Name') or '',
                'qty': _num(r.get('Damaged_Qty')),
                'rate': _num(r.get('Damage_Rate')),
                'desc': r.get('Description') or '',
            })
        for c in all_challans:
            c['Received_Qty'] = fmt_qty(c['Received_Qty'])
            c['Breakage_Qty'] = fmt_qty(c['Breakage_Qty'])
            c['Breakage_Amount'] = fmt_money(c['Breakage_Amount'])
    except Exception:
        logger.exception('Error loading receive challan list')
        flash('Could not read the Excel data file. Please check that it exists and is not corrupted.', 'danger')

    try:
        next_no = next_receive_challan_no(book)          # preview: 01, 02, 03 ...
    except Exception:
        logger.exception('Could not compute next receive challan number')

    all_challans = list(reversed(all_challans))               # newest first

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
            d = parse_any_date(c.get('Receive_Date'))
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
        fields = ('Receive_Challan_No', 'Party_Name', 'Site_Name', 'Party_Mobile', 'Vehicle_No', 'Driver_Name')
        filtered = [c for c in filtered if any(q in str(c.get(f, '')).lower() for f in fields)]

    total_records = len(filtered)
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
        'receive_challan/index.html',
        challans=page_rows,
        parties=parties,
        sites_json=to_js_json(sites),
        challans_json=to_js_json(all_challans),
        receive_data_json=to_js_json(receive_data),
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
        next_challan_no=next_no,
    )


@receive_challan_bp.route('/stock')
def stock():
    """Stock lying with a party at a site (JSON) - used by the create / edit popup."""
    party_id = request.args.get('party_id', '').strip()
    site_id = request.args.get('site_id', '').strip()
    exclude_no = request.args.get('exclude', '').strip()
    if not party_id or not site_id:
        return jsonify({'items': [], 'total_stock': 0})
    try:
        book = Book.load()
        exclude_id = None
        if exclude_no:
            df = book.df(receive_sheet(book))
            if not df.empty and 'Receive_Challan_No' in df.columns:
                match = df[df['Receive_Challan_No'].astype(str) == exclude_no]
                if not match.empty:
                    exclude_id = sid(match.iloc[0]['Receive_ID'])

        summary = stock_summary(book, party_id, site_id, exclude_receive_id=exclude_id)
        master = {sid(i.get('Item_ID')): i for i in load_items_master(book)}
        items = []
        for iid, v in summary.items():
            if v['outstanding'] <= 0:
                continue
            m = master.get(iid, {})
            items.append({
                'item_id': iid,
                'name': item_label(m) if m else f'Item #{iid}',
                'unit_weight': to_float(m.get('Unit_Weight_KG'), DEFAULT_UNIT_WEIGHT) or DEFAULT_UNIT_WEIGHT,
                'issued': _num(v['opening'] + v['issued']),
                'received': _num(v['received']),
                'outstanding': _num(v['outstanding']),
            })
        items.sort(key=lambda x: x['name'].lower())
        return jsonify({'items': items, 'total_stock': _num(sum(i['outstanding'] for i in items))})
    except Exception:
        logger.exception('Error loading party stock')
        return jsonify({'items': [], 'total_stock': 0, 'error': 'Could not load stock.'}), 500


@receive_challan_bp.route('/create', methods=['POST'])
@safe_action('Error creating receive challan.', ENDPOINT_INDEX)
def create():
    f = request.form
    with excel_lock:
        book = Book.load()
        if not book.sheets:
            raise ChallanError('Excel data file not found.')

        r_sheet = receive_sheet(book) or RECEIVE_SHEET
        ri_sheet = receive_items_sheet(book) or RECEIVE_ITEMS_SHEET
        d_sheet = damage_sheet(book) or DAMAGE_SHEET
        rdf, ridf, ddf = book.df(r_sheet), book.df(ri_sheet), book.df(d_sheet)

        parties = {sid(p.get('Party_ID')): p for p in load_parties(book)}
        sites = {sid(s.get('Site_ID')): s for s in load_sites(book)}
        master = {sid(i.get('Item_ID')): i for i in load_items_master(book)}

        header = parse_header(f, parties, sites)
        received = parse_received(f, master)
        breakage = parse_breakage(f, master)
        if not received and not breakage:
            raise ChallanError('Please enter the received quantity (or a breakage item) before saving.')

        stock = stock_summary(book, header['Party_ID'], header['Site_ID'])
        validate_against_stock(received, breakage, stock, master)

        receive_no = next_receive_challan_no(book)       # server generated: 01, 02, 03 ...
        receive_id = next_id(rdf, 'Receive_ID')

        item_lines = number_rows(build_item_lines(received, breakage), 'Line_ID',
                                 next_id(ridf, 'Line_ID'), {'Receive_ID': receive_id, 'Issue_Line_ID': ''})
        # keep the column order of the sheet:  Line_ID, Receive_ID, Issue_Line_ID, Item_ID ...
        charge_rows = damage_rows(breakage, next_id(ddf, 'Damage_Charge_ID'), receive_id, header)

        new_row = {
            'Receive_ID': receive_id,
            'Receive_Challan_No': receive_no,
            'Issue_ID': '',
            'Party_ID': header['Party_ID'],
            'Site_ID': header['Site_ID'],
            'Receive_Date': header['Receive_Date'],
            'Remarks': header['Remarks'],
            'Status': 'Active',
            'Created_At': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'Created_By': current_user(),
            'Cancel_Reason': '',
            'Vehicle_No': header['Vehicle_No'],
            'Driver_Name': header['Driver_Name'],
            'Driver_Mob': header['Driver_Mob'],
        }

        a_sheet, audit_df = audit_update(book, 'CREATE', receive_no, entity_type='Receive_Challan',
                                         after={**new_row, 'Items': item_lines, 'Breakage': charge_rows})
        save_sheets({
            r_sheet: append_rows(rdf, [new_row]),
            ri_sheet: append_rows(ridf, item_lines),
            d_sheet: append_rows(ddf, charge_rows),
            a_sheet: audit_df,
        })

    flash(f'Receive challan {receive_no} created successfully!', 'success')
    return redirect(url_for('receive_challan.print_challan', challan_no=receive_no))


@receive_challan_bp.route('/edit/<path:challan_no>', methods=['POST'])
@safe_action('Error updating receive challan.', ENDPOINT_INDEX)
def edit(challan_no):
    f = request.form
    with excel_lock:
        book = Book.load()
        r_sheet, rdf, idx = find_receive(book, challan_no)
        before = records(rdf.loc[[idx]])[0]
        if is_cancelled(before):
            raise ChallanError('A cancelled receive challan cannot be edited.')
        receive_id = norm_id(before.get('Receive_ID'))

        ri_sheet = receive_items_sheet(book) or RECEIVE_ITEMS_SHEET
        d_sheet = damage_sheet(book) or DAMAGE_SHEET
        ridf, ddf = book.df(ri_sheet), book.df(d_sheet)

        mine_charges = (ddf['Receive_ID'].apply(sid) == sid(receive_id)) if (not ddf.empty and 'Receive_ID' in ddf.columns) else None
        old_charges = records(ddf[mine_charges]) if mine_charges is not None else []
        if any(is_billed(c) for c in old_charges):
            raise ChallanError('Breakage charges of this receive challan are already billed, so it cannot be edited.')

        parties = {sid(p.get('Party_ID')): p for p in load_parties(book)}
        sites = {sid(s.get('Site_ID')): s for s in load_sites(book)}
        master = {sid(i.get('Item_ID')): i for i in load_items_master(book)}

        header = parse_header(f, parties, sites)
        received = parse_received(f, master)
        breakage = parse_breakage(f, master)
        if not received and not breakage:
            raise ChallanError('Please enter the received quantity (or a breakage item) before saving.')

        # this challan's own earlier returns are excluded, so they can be re-entered
        stock = stock_summary(book, header['Party_ID'], header['Site_ID'], exclude_receive_id=receive_id)
        validate_against_stock(received, breakage, stock, master)

        ensure_columns(rdf, ('Receive_Date', 'Party_ID', 'Site_ID', 'Remarks', 'Vehicle_No', 'Driver_Name', 'Driver_Mob'))
        for col, val in header.items():
            rdf.at[idx, col] = val

        # replace item lines + breakage rows (IDs keep increasing, never reused)
        new_lines = number_rows(build_item_lines(received, breakage), 'Line_ID',
                                next_id(ridf, 'Line_ID'), {'Receive_ID': receive_id, 'Issue_Line_ID': ''})
        new_charges = damage_rows(breakage, next_id(ddf, 'Damage_Charge_ID'), receive_id, header)
        if not ridf.empty and 'Receive_ID' in ridf.columns:
            ridf = ridf[ridf['Receive_ID'].apply(sid) != sid(receive_id)]
        if mine_charges is not None:
            ddf = ddf[~mine_charges]

        after = records(rdf.loc[[idx]])[0]
        after['Items'], after['Breakage'] = new_lines, new_charges
        a_sheet, audit_df = audit_update(book, 'UPDATE', str(challan_no), before=before, after=after,
                                         entity_type='Receive_Challan')
        save_sheets({
            r_sheet: rdf,
            ri_sheet: append_rows(ridf, new_lines),
            d_sheet: append_rows(ddf, new_charges),
            a_sheet: audit_df,
        })

    flash(f'Receive challan {challan_no} updated successfully!', 'success')
    return redirect(url_for('receive_challan.index'))


@receive_challan_bp.route('/delete/<path:challan_no>', methods=['POST'])
@safe_action('Error deleting receive challan.', ENDPOINT_INDEX)
def delete(challan_no):
    with excel_lock:
        book = Book.load()
        r_sheet, rdf, idx = find_receive(book, challan_no)
        before = records(rdf.loc[[idx]])[0]
        receive_id = sid(before.get('Receive_ID'))

        updates = {r_sheet: rdf.drop(index=idx)}

        ri_sheet = receive_items_sheet(book)
        if ri_sheet:
            ridf = book.df(ri_sheet)
            if not ridf.empty and 'Receive_ID' in ridf.columns:
                gone = ridf['Receive_ID'].apply(sid) == receive_id
                before['Items'] = records(ridf[gone])
                updates[ri_sheet] = ridf[~gone]

        d_sheet = damage_sheet(book)
        if d_sheet:
            ddf = book.df(d_sheet)
            if not ddf.empty and 'Receive_ID' in ddf.columns:
                gone = ddf['Receive_ID'].apply(sid) == receive_id
                charges = records(ddf[gone])
                if any(is_billed(c) for c in charges):
                    raise ChallanError('Breakage charges of this receive challan are already billed, '
                                       'so it cannot be deleted.')
                before['Breakage'] = charges
                updates[d_sheet] = ddf[~gone]

        a_sheet, audit_df = audit_update(book, 'DELETE', str(challan_no), before=before,
                                         entity_type='Receive_Challan')
        updates[a_sheet] = audit_df
        save_sheets(updates)

    flash(f'Receive challan {challan_no} deleted successfully!', 'success')
    return redirect(url_for('receive_challan.index'))


@receive_challan_bp.route('/print/<path:challan_no>')
def print_challan(challan_no):
    challan = None
    party, site = {}, {}
    items, breakage = [], []
    total_qty = 0.0
    total_weight = 0.0
    breakage_qty = 0.0
    breakage_amt = 0.0

    try:
        book = Book.load()
        df = book.df(receive_sheet(book))
        if not df.empty and 'Receive_Challan_No' in df.columns:
            match = df[df['Receive_Challan_No'].astype(str) == str(challan_no)]
            if not match.empty:
                challan = records(match)[0]

        if challan:
            rid = sid(challan.get('Receive_ID'))
            party = next((p for p in load_parties(book)
                          if sid(p.get('Party_ID')) == sid(challan.get('Party_ID'))), {})
            site = next((s for s in load_sites(book)
                         if sid(s.get('Site_ID')) == sid(challan.get('Site_ID'))), {})
            master = {sid(i.get('Item_ID')): i for i in load_items_master(book)}

            for r in records(book.df(receive_items_sheet(book))):
                if sid(r.get('Receive_ID')) != rid:
                    continue
                m = master.get(sid(r.get('Item_ID')), {})
                accepted = to_float(r.get('Accepted_Qty'))
                returned = to_float(r.get('Returned_Qty')) or accepted + to_float(r.get('Damaged_Qty'))
                total_weight += returned * (to_float(m.get('Unit_Weight_KG'), DEFAULT_UNIT_WEIGHT) or DEFAULT_UNIT_WEIGHT)
                if accepted > 0:
                    total_qty += accepted
                    items.append({'Description': item_label(m) if m else f"Item #{sid(r.get('Item_ID'))}",
                                  'Qty_Display': fmt_qty(accepted)})

            for r in records(book.df(damage_sheet(book))):
                if sid(r.get('Receive_ID')) != rid:
                    continue
                m = master.get(sid(r.get('Item_ID')), {})
                qty = to_float(r.get('Damaged_Qty'))
                amt = to_float(r.get('Damage_Amount'))
                breakage_qty += qty
                breakage_amt += amt
                breakage.append({
                    'Name': (item_label(m) if m else '') or r.get('Item_Name') or '-',
                    'Qty': fmt_qty(qty),
                    'Rate': fmt_money(r.get('Damage_Rate')),
                    'Amount': fmt_money(amt),
                    'Description': r.get('Description') or '',
                })
    except Exception:
        logger.exception('Error loading receive challan print data')

    html = render_template(
        'receive_challan/print.html',
        challan=challan,
        party=party,
        site=site,
        items=items,
        breakage=breakage,
        challan_no=challan_no,
        total_qty=fmt_qty(total_qty),
        total_qty_words=number_to_words(total_qty),
        breakage_qty=fmt_qty(breakage_qty),
        breakage_amount=fmt_money(breakage_amt),
        weight_display=f"{fmt_qty(round(total_weight, 2))} Kg",
    )
    return html, (200 if challan else 404)
