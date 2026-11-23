"""
Stock module - OM Enterprise ERP
--------------------------------
/stock/godown       : stock in OUR godown (what can be issued) + what is out with parties
/stock/entries      : godown stock IN / OUT entries (purchase, sale, scrap, adjustment)
/stock/statement    : item-wise godown statement with running balance (printable)
/stock/party-site   : party & site wise stock -> handled by the Stock Ledger page

Godown formula (per item):
    opening         OpeningStock rows whose Location_Type contains "godown"
    + purchases     Godown_Entries IN  (Purchase, Adjustment (+))
    - sales         Godown_Entries OUT (Sale, Scrap / Write-off, Adjustment (-))
    - issued        Issue_Challan_Items.Quantity (non-cancelled)
    + returned      (Returned_Qty - Damaged_Qty) of non-cancelled receive challans
    = available

    with_parties = party opening + issued - returned (good + damaged)
"""
import logging
from collections import defaultdict
from datetime import date, datetime

from flask import Blueprint, flash, redirect, render_template, request, url_for

from app.routes.challan_store import (
    Book, ChallanError, _IGNORED_OPENING_STATUS, _issue_items_sheet,
    _issue_sheet, _opening_qty, _opening_sheet, append_rows, audit_update,
    current_user, excel_lock, fmt_money, fmt_qty, is_active, is_cancelled,
    item_label, load_items_master, load_parties, next_id, next_number,
    normalize_date, parse_any_date, receive_items_sheet, receive_sheet,
    records, safe_action, save_sheets, sid, to_float, to_int,
)

logger = logging.getLogger(__name__)

stock_bp = Blueprint('stock', __name__, url_prefix='/stock')

ENTRY_SHEET = 'Godown_Entries'
IN_TYPES = ('Purchase', 'Adjustment (+)')
OUT_TYPES = ('Sale', 'Scrap / Write-off', 'Adjustment (-)')
ENTRY_TYPES = IN_TYPES + OUT_TYPES
REMARKS_REQUIRED = ('Adjustment (+)', 'Adjustment (-)', 'Scrap / Write-off')


# ----------------------------------------------------------------------------
# Data helpers
# ----------------------------------------------------------------------------
def _entries_sheet(book):
    return book.find(exact=('godown_entries',))


def _active_entries(book):
    return [e for e in records(book.df(_entries_sheet(book))) if not is_cancelled(e)]


def _new_stat():
    return {'g_open': 0.0, 'g_in': 0.0, 'g_out': 0.0, 'p_open': 0.0,
            'issued': 0.0, 'ret_total': 0.0, 'damaged': 0.0}


def _available(st):
    returned_good = st['ret_total'] - st['damaged']
    return st['g_open'] + st['g_in'] - st['g_out'] - st['issued'] + returned_good


def _collect_stats(book):
    """Item-wise numbers across ALL parties and sites."""
    stats = defaultdict(_new_stat)

    # opening stock (godown vs. already lying with parties)
    for r in records(book.df(_opening_sheet(book))):
        if str(r.get('Status', '')).strip().lower() in _IGNORED_OPENING_STATUS:
            continue
        iid = sid(r.get('Item_ID'))
        if not iid:
            continue
        qty = _opening_qty(r)
        if 'godown' in str(r.get('Location_Type', '')).strip().lower():
            stats[iid]['g_open'] += qty
        else:
            stats[iid]['p_open'] += qty

    # purchases / sales / adjustments
    for e in _active_entries(book):
        iid = sid(e.get('Item_ID'))
        if not iid:
            continue
        qty = to_float(e.get('Quantity'))
        if e.get('Entry_Type') in IN_TYPES:
            stats[iid]['g_in'] += qty
        elif e.get('Entry_Type') in OUT_TYPES:
            stats[iid]['g_out'] += qty

    # issued from godown
    valid_issue = {sid(c.get('Issue_ID')) for c in records(book.df(_issue_sheet(book)))
                   if not is_cancelled(c)}
    for r in records(book.df(_issue_items_sheet(book))):
        if sid(r.get('Issue_ID')) in valid_issue:
            iid = sid(r.get('Item_ID'))
            if iid:
                stats[iid]['issued'] += to_float(r.get('Quantity'))

    # returned to godown
    valid_receive = {sid(c.get('Receive_ID')) for c in records(book.df(receive_sheet(book)))
                     if not is_cancelled(c)}
    for r in records(book.df(receive_items_sheet(book))):
        if sid(r.get('Receive_ID')) in valid_receive:
            iid = sid(r.get('Item_ID'))
            if not iid:
                continue
            damaged = to_float(r.get('Damaged_Qty'))
            total = to_float(r.get('Returned_Qty'))
            if not total:
                total = to_float(r.get('Accepted_Qty')) + damaged
            stats[iid]['ret_total'] += total
            stats[iid]['damaged'] += damaged

    return stats


def _active_items(book):
    items = [(sid(i.get('Item_ID')), item_label(i)) for i in load_items_master(book)
             if sid(i.get('Item_ID')) and is_active(i)]
    return sorted(items, key=lambda x: x[1].lower())


# ----------------------------------------------------------------------------
# Godown page
# ----------------------------------------------------------------------------
@stock_bp.route('/godown', methods=['GET'])
def godown_stock():
    """Current stock in the godown, item-wise."""
    search = request.args.get('search', '').strip().lower()
    f_category = request.args.get('category', '').strip()
    f_stock = request.args.get('stock', '').strip()
    show_all = request.args.get('show') == 'all'
    
    per_page = request.args.get('per_page', '10')
    if per_page not in PER_PAGE_OPTIONS:
        per_page = '10'

    rows = []
    categories_set = set()
    totals = {k: '0' for k in ('opening', 'issued', 'returned_good', 'available',
                               'with_parties', 'damaged')}
    negative = False
    
    try:
        book = Book.load()
        master = {sid(i.get('Item_ID')): i for i in load_items_master(book)}
        stats = _collect_stats(book)

        for m in master.values():
            cat = str(m.get('Category', '')).strip()
            if cat:
                categories_set.add(cat)

        item_ids = [iid for iid, m in master.items() if iid and is_active(m)]
        item_ids += [iid for iid in stats if iid not in master]

        sums = defaultdict(float)
        for iid in item_ids:
            m = master.get(iid)
            st = stats.get(iid) or _new_stat()
            label = item_label(m) if m else f'Item #{iid}'
            category = (m.get('Category', '') if m else '') or '-'

            if search and search not in label.lower() and search not in str(category).lower():
                continue
            if f_category and category != f_category:
                continue

            opening = st['g_open'] + st['g_in'] - st['g_out']
            returned_good = st['ret_total'] - st['damaged']
            available = _available(st)
            with_parties = st['p_open'] + st['issued'] - st['ret_total']

            if f_stock == 'available' and available <= 0:
                continue
            elif f_stock == 'zero' and abs(available) > 1e-9:
                continue
            elif f_stock == 'negative' and available >= 0:
                continue
            elif f_stock == 'with_parties' and with_parties <= 0:
                continue

            if not any(abs(v) > 0 for v in st.values()) and not show_all:
                continue

            if available < 0:
                negative = True

            rows.append({
                'item_description': label,
                'category': category,
                'item_id': iid,
                'opening': fmt_qty(opening),
                'issued': fmt_qty(st['issued']),
                'returned_good': fmt_qty(returned_good),
                'damaged': fmt_qty(st['damaged']),
                'available': fmt_qty(available),
                'available_raw': available,
                'with_parties': fmt_qty(with_parties),
                'with_parties_raw': with_parties,
            })
            sums['opening'] += opening
            sums['issued'] += st['issued']
            sums['returned_good'] += returned_good
            sums['damaged'] += st['damaged']
            sums['available'] += available
            sums['with_parties'] += with_parties

        rows.sort(key=lambda r: r['item_description'].lower())
        for n, r in enumerate(rows, 1):
            r['s_no'] = n
        totals = {k: fmt_qty(sums[k]) for k in totals}

    except Exception:
        logger.exception('Error loading godown stock')
        flash('Could not read the Excel data file. Please check that it exists and is not corrupted.', 'danger')

    # ---- Pagination Logic --------------------------------------------------
    total = len(rows)
    if per_page == 'All':
        size, pages = max(total, 1), 1
    else:
        size = int(per_page)
        pages = max(1, -(-total // size))
    
    page = min(max(to_int(request.args.get('page'), 1), 1), pages)
    start = (page - 1) * size
    page_rows = rows[start:start + size]
    
    pagination = {
        'page': page, 'pages': pages, 'total': total,
        'first': start + 1 if total else 0, 'last': start + len(page_rows),
        'prev': page - 1 if page > 1 else None,
        'next': page + 1 if page < pages else None,
        'numbers': _page_numbers(page, pages),
    }

    qargs = {k: v for k, v in {
        'search': search, 'category': f_category, 'stock': f_stock,
        'show': 'all' if show_all else '', 'per_page': per_page,
    }.items() if v}

    return render_template(
        'stock/godown.html', 
        rows=page_rows, 
        totals=totals, 
        negative=negative,
        search=search, 
        show_all=show_all, 
        pagination=pagination, 
        qargs=qargs, 
        categories=sorted(list(categories_set)),
        f_category=f_category,
        f_stock=f_stock,
        per_page=per_page,
        per_page_options=PER_PAGE_OPTIONS
    )


# ----------------------------------------------------------------------------
# Godown entries: purchase / sale / scrap / adjustment
# ----------------------------------------------------------------------------
PER_PAGE_OPTIONS = ('10', '25', '50', 'All')


def _page_numbers(page, pages):
    """[1, None, 4, 5, 6, None, 20] - None means '...'."""
    wanted = sorted({1, pages, page - 2, page - 1, page, page + 1, page + 2})
    wanted = [n for n in wanted if 1 <= n <= pages]
    out, prev = [], 0
    for n in wanted:
        if prev and n - prev > 1:
            out.append(None)
        out.append(n)
        prev = n
    return out


@stock_bp.route('/entries', methods=['GET'])
@safe_action('Could not load godown entries.', 'stock.godown_stock')
def entries():
    book = Book.load()
    items = _active_items(book)
    labels = dict(items)

    # ---- filters -----------------------------------------------------------
    f_item = sid(request.args.get('item_id'))
    f_type = request.args.get('entry_type', '').strip()
    if f_type not in ENTRY_TYPES:
        f_type = ''
    f_status = request.args.get('status', '').strip().lower()
    if f_status not in ('active', 'cancelled'):
        f_status = ''
    f_from_raw = request.args.get('from', '').strip()
    f_to_raw = request.args.get('to', '').strip()
    f_from, f_to = parse_any_date(f_from_raw), parse_any_date(f_to_raw)
    f_q = request.args.get('q', '').strip()
    per_page = request.args.get('per_page', '10')
    if per_page not in PER_PAGE_OPTIONS:
        per_page = '10'

    all_entries = records(book.df(_entries_sheet(book)))
    rows = []
    for e in all_entries:
        iid = sid(e.get('Item_ID'))
        e['item_label'] = labels.get(iid) or f'Item #{iid}'
        e['cancelled'] = is_cancelled(e)
        d = parse_any_date(e.get('Entry_Date'))

        if f_item and iid != f_item:
            continue
        if f_type and e.get('Entry_Type') != f_type:
            continue
        if f_status == 'active' and e['cancelled']:
            continue
        if f_status == 'cancelled' and not e['cancelled']:
            continue
        if f_from and (not d or d < f_from):
            continue
        if f_to and (not d or d > f_to):
            continue
        if f_q:
            hay = ' '.join(str(e.get(k, '')) for k in
                           ('Entry_No', 'Party_Name', 'Reference_No', 'Remarks', 'item_label')).lower()
            if f_q.lower() not in hay:
                continue

        e['qty_fmt'] = fmt_qty(e.get('Quantity'))
        e['rate_fmt'] = fmt_money(e.get('Rate')) if to_float(e.get('Rate')) else '-'
        e['direction'] = 'in' if e.get('Entry_Type') in IN_TYPES else 'out'
        e['_date'] = d or date.min
        rows.append(e)

    rows.sort(key=lambda e: (e['_date'], to_float(e.get('Entry_ID'))), reverse=True)

    total_in = sum(to_float(e.get('Quantity')) for e in rows if not e['cancelled'] and e['direction'] == 'in')
    total_out = sum(to_float(e.get('Quantity')) for e in rows if not e['cancelled'] and e['direction'] == 'out')
    summary = {'in': fmt_qty(total_in), 'out': fmt_qty(total_out), 'net': fmt_qty(total_in - total_out)}

    total = len(rows)
    if per_page == 'All':
        size, pages = max(total, 1), 1
    else:
        size = int(per_page)
        pages = max(1, -(-total // size))
    page = min(max(to_int(request.args.get('page'), 1), 1), pages)
    start = (page - 1) * size
    page_rows = rows[start:start + size]
    pagination = {
        'page': page, 'pages': pages, 'total': total,
        'first': start + 1 if total else 0, 'last': start + len(page_rows),
        'prev': page - 1 if page > 1 else None,
        'next': page + 1 if page < pages else None,
        'numbers': _page_numbers(page, pages),
    }

    qargs = {k: v for k, v in {
        'item_id': f_item, 'entry_type': f_type, 'status': f_status,
        'from': f_from_raw, 'to': f_to_raw, 'q': f_q, 'per_page': per_page,
    }.items() if v}

    return render_template(
        'stock/entries.html', rows=page_rows, items=items, entry_types=ENTRY_TYPES,
        today=datetime.now().strftime('%d-%m-%Y'), pagination=pagination, summary=summary,
        qargs=qargs, per_page=per_page, per_page_options=PER_PAGE_OPTIONS,
        f_item=f_item, f_type=f_type, f_status=f_status,
        f_from=f_from_raw, f_to=f_to_raw, f_q=f_q, all_count=len(all_entries),
    )


@stock_bp.route('/entries/add', methods=['POST'])
@safe_action('Could not save the stock entry.', 'stock.entries')
def add_entry():
    f = request.form
    entry_type = f.get('entry_type', '').strip()
    if entry_type not in ENTRY_TYPES:
        raise ChallanError('Please select a valid entry type.')
    item_id = sid(f.get('item_id'))
    qty = to_float(f.get('quantity'))
    if qty <= 0:
        raise ChallanError('Quantity must be greater than zero.')
    entry_date = normalize_date(f.get('entry_date'))
    remarks = f.get('remarks', '').strip()
    if entry_type in REMARKS_REQUIRED and not remarks:
        raise ChallanError('Remarks (reason) are required for scrap / adjustment entries.')

    with excel_lock:
        book = Book.load()
        master = {sid(i.get('Item_ID')): i for i in load_items_master(book)}
        m = master.get(item_id)
        if not m or not is_active(m):
            raise ChallanError('Please select a valid item.')

        if entry_type in OUT_TYPES:
            avail = _available(_collect_stats(book).get(item_id) or _new_stat())
            if qty > avail + 1e-9:
                raise ChallanError(f'Only {fmt_qty(avail)} of {item_label(m)} is available in the godown.')

        sheet = _entries_sheet(book) or ENTRY_SHEET
        df = book.df(sheet)
        existing = records(df)
        entry_id = next_id(df, 'Entry_ID')
        row = {
            'Entry_ID': entry_id,
            'Entry_No': next_number('OM/GS/', existing, 'Entry_No'),
            'Entry_Date': entry_date,
            'Entry_Type': entry_type,
            'Item_ID': int(item_id) if item_id.isdigit() else item_id,
            'Quantity': qty,
            'Rate': to_float(f.get('rate')),
            'Party_Name': f.get('party_name', '').strip(),
            'Reference_No': f.get('reference_no', '').strip(),
            'Remarks': remarks,
            'Status': 'Active',
            'Created_At': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'Created_By': current_user(),
            'Cancel_Reason': '',
            'Cancelled_At': '',
        }
        audit_name, audit_df = audit_update(book, 'CREATE', entry_id, after=row,
                                            entity_type='Godown_Entry')
        save_sheets({sheet: append_rows(df, [row]), audit_name: audit_df})

    flash(f"{entry_type} entry {row['Entry_No']} saved.", 'success')
    return redirect(url_for('stock.entries'))


@stock_bp.route('/entries/<entry_id>/cancel', methods=['POST'])
@safe_action('Could not cancel the stock entry.', 'stock.entries')
def cancel_entry(entry_id):
    reason = request.form.get('reason', '').strip()
    if not reason:
        raise ChallanError('Please give a reason for cancelling the entry.')

    with excel_lock:
        book = Book.load()
        sheet = _entries_sheet(book)
        df = book.df(sheet)
        if df.empty or 'Entry_ID' not in df.columns:
            raise ChallanError('Entry not found.')
        mask = df['Entry_ID'].map(sid) == sid(entry_id)
        if not mask.any():
            raise ChallanError('Entry not found.')
        before = records(df[mask])[0]
        if is_cancelled(before):
            raise ChallanError('This entry is already cancelled.')

        iid = sid(before.get('Item_ID'))
        if before.get('Entry_Type') in IN_TYPES:
            avail = _available(_collect_stats(book).get(iid) or _new_stat())
            if avail - to_float(before.get('Quantity')) < -1e-9:
                raise ChallanError('Cannot cancel: godown stock of this item would become negative '
                                   '(it has already been issued or sold).')

        for col in ('Cancel_Reason', 'Cancelled_At'):
            if col not in df.columns:
                df[col] = ''
            df[col] = df[col].astype(object)
        df.loc[mask, 'Status'] = 'Cancelled'
        df.loc[mask, 'Cancel_Reason'] = reason
        df.loc[mask, 'Cancelled_At'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        audit_name, audit_df = audit_update(book, 'CANCEL', sid(entry_id), before=before,
                                            reason=reason, entity_type='Godown_Entry')
        save_sheets({sheet: df, audit_name: audit_df})

    flash('Entry cancelled.', 'success')
    return redirect(url_for('stock.entries'))


# ----------------------------------------------------------------------------
# Godown statement (running balance)
# ----------------------------------------------------------------------------
def _build_statement(book, item_id, from_d, to_d):
    parties = {sid(p.get('Party_ID')): p.get('Party_Name') or '' for p in load_parties(book)}
    early = date(1900, 1, 1)
    events = []       # dict(date, order, type, ref, note, in, out)

    for r in records(book.df(_opening_sheet(book))):
        if sid(r.get('Item_ID')) != item_id:
            continue
        if str(r.get('Status', '')).strip().lower() in _IGNORED_OPENING_STATUS:
            continue
        if 'godown' not in str(r.get('Location_Type', '')).strip().lower():
            continue
        q = _opening_qty(r)
        if q:
            d = parse_any_date(r.get('Verified_Date') or r.get('Estimated_Date')) or early
            events.append({'date': d, 'order': 0, 'type': 'Opening Stock', 'ref': '-',
                           'note': '', 'in': q, 'out': 0.0})

    for e in _active_entries(book):
        if sid(e.get('Item_ID')) != item_id:
            continue
        q = to_float(e.get('Quantity'))
        is_in = e.get('Entry_Type') in IN_TYPES
        note = ' | '.join(x for x in (str(e.get('Party_Name') or ''),
                                     f"Ref: {e.get('Reference_No')}" if e.get('Reference_No') else '',
                                     str(e.get('Remarks') or '')) if x)
        events.append({'date': parse_any_date(e.get('Entry_Date')) or early, 'order': 1,
                       'type': e.get('Entry_Type'), 'ref': e.get('Entry_No') or '-',
                       'note': note, 'in': q if is_in else 0.0, 'out': 0.0 if is_in else q})

    rec_head = {sid(c.get('Receive_ID')): c for c in records(book.df(receive_sheet(book)))
                if not is_cancelled(c)}
    for r in records(book.df(receive_items_sheet(book))):
        h = rec_head.get(sid(r.get('Receive_ID')))
        if not h or sid(r.get('Item_ID')) != item_id:
            continue
        damaged = to_float(r.get('Damaged_Qty'))
        total = to_float(r.get('Returned_Qty')) or (to_float(r.get('Accepted_Qty')) + damaged)
        good = total - damaged
        if not good and not damaged:
            continue
        note = parties.get(sid(h.get('Party_ID')), '')
        if damaged:
            note += f" | Damaged: {fmt_qty(damaged)} (not added to godown)"
        events.append({'date': parse_any_date(h.get('Receive_Date')) or early, 'order': 2,
                       'type': 'Receive Challan',
                       'ref': h.get('Receive_No') or h.get('Receive_Challan_No') or '-',
                       'note': note.strip(' |'), 'in': good, 'out': 0.0})

    iss_head = {sid(c.get('Issue_ID')): c for c in records(book.df(_issue_sheet(book)))
                if not is_cancelled(c)}
    for r in records(book.df(_issue_items_sheet(book))):
        h = iss_head.get(sid(r.get('Issue_ID')))
        if not h or sid(r.get('Item_ID')) != item_id:
            continue
        q = to_float(r.get('Quantity'))
        if q:
            events.append({'date': parse_any_date(h.get('Challan_Date')) or early, 'order': 3,
                           'type': 'Issue Challan', 'ref': h.get('Challan_No') or '-',
                           'note': parties.get(sid(h.get('Party_ID')), ''), 'in': 0.0, 'out': q})

    events.sort(key=lambda x: (x['date'], x['order']))

    balance = 0.0
    if from_d:
        balance = sum(e['in'] - e['out'] for e in events if e['date'] < from_d)
        events = [e for e in events if e['date'] >= from_d]
    if to_d:
        events = [e for e in events if e['date'] <= to_d]
    bf = balance

    rows, total_in, total_out = [], 0.0, 0.0
    for e in events:
        balance += e['in'] - e['out']
        total_in += e['in']
        total_out += e['out']
        rows.append({
            'date': e['date'].strftime('%d-%m-%Y') if e['date'] != early else '-',
            'type': e['type'], 'ref': e['ref'], 'note': e['note'],
            'in_q': fmt_qty(e['in']) if e['in'] else '',
            'out_q': fmt_qty(e['out']) if e['out'] else '',
            'balance': fmt_qty(balance), 'negative': balance < 0,
        })
    return rows, {'bf': fmt_qty(bf), 'in': fmt_qty(total_in), 'out': fmt_qty(total_out),
                  'closing': fmt_qty(balance)}


@stock_bp.route('/statement', methods=['GET'])
@safe_action('Could not prepare the godown statement.', 'stock.godown_stock')
def statement():
    book = Book.load()
    items = _active_items(book)
    item_id = sid(request.args.get('item_id'))
    from_raw = request.args.get('from', '').strip()
    to_raw = request.args.get('to', '').strip()

    rows, totals, item_name = [], None, ''
    if item_id:
        item_name = dict(items).get(item_id) or f'Item #{item_id}'
        rows, totals = _build_statement(book, item_id, parse_any_date(from_raw), parse_any_date(to_raw))

    return render_template('stock/statement.html', items=items, item_id=item_id, item_name=item_name,
                           from_raw=from_raw, to_raw=to_raw, rows=rows, totals=totals,
                           printed_on=datetime.now().strftime('%d-%m-%Y %H:%M'))


@stock_bp.route('/party-site', methods=['GET'])
def party_site_stock():
    """Party & site wise outstanding stock lives on the Stock Ledger page."""
    return redirect(url_for('stock_ledger.index', **request.args))


@stock_bp.route('/statement/print')
@safe_action('Could not prepare the print statement.', 'stock.godown_stock')
def statement_print():
    book = Book.load()
    items = _active_items(book)
    item_id = sid(request.args.get('item_id'))
    from_raw = request.args.get('from', '').strip()
    to_raw = request.args.get('to', '').strip()

    rows, totals, item_name = [], None, ''
    if item_id:
        item_name = dict(items).get(item_id) or f'Item #{item_id}'
        rows, totals = _build_statement(book, item_id, parse_any_date(from_raw), parse_any_date(to_raw))

    return render_template(
        'stock/statement_print.html',
        item_id=item_id,
        item_name=item_name,
        from_raw=from_raw,
        to_raw=to_raw,
        rows=rows,
        totals=totals,
        printed_on=datetime.now().strftime('%d-%m-%Y %H:%M')
    )


@stock_bp.route('/godown/print', methods=['GET'])
@safe_action('Could not prepare the godown print stock.', 'stock.godown_stock')
def godown_print():
    """Printable view for current stock in the godown."""
    search = request.args.get('search', '').strip().lower()
    f_category = request.args.get('category', '').strip()
    f_stock = request.args.get('stock', '').strip()
    show_all = request.args.get('show') == 'all'

    rows = []
    totals = {k: '0' for k in ('opening', 'issued', 'returned_good', 'available',
                               'with_parties', 'damaged')}
    try:
        book = Book.load()
        master = {sid(i.get('Item_ID')): i for i in load_items_master(book)}
        stats = _collect_stats(book)

        item_ids = [iid for iid, m in master.items() if iid and is_active(m)]
        item_ids += [iid for iid in stats if iid not in master]

        sums = defaultdict(float)
        for iid in item_ids:
            m = master.get(iid)
            st = stats.get(iid) or _new_stat()
            label = item_label(m) if m else f'Item #{iid}'
            category = (m.get('Category', '') if m else '') or '-'

            if search and search not in label.lower() and search not in str(category).lower():
                continue
            if f_category and category != f_category:
                continue

            opening = st['g_open'] + st['g_in'] - st['g_out']
            returned_good = st['ret_total'] - st['damaged']
            available = _available(st)
            with_parties = st['p_open'] + st['issued'] - st['ret_total']

            if f_stock == 'available' and available <= 0:
                continue
            elif f_stock == 'zero' and abs(available) > 1e-9:
                continue
            elif f_stock == 'negative' and available >= 0:
                continue
            elif f_stock == 'with_parties' and with_parties <= 0:
                continue

            if not any(abs(v) > 0 for v in st.values()) and not show_all:
                continue

            rows.append({
                'item_description': label,
                'category': category,
                'item_id': iid,
                'opening': fmt_qty(opening),
                'issued': fmt_qty(st['issued']),
                'returned_good': fmt_qty(returned_good),
                'damaged': fmt_qty(st['damaged']),
                'available': fmt_qty(available),
                'available_raw': available,
                'with_parties': fmt_qty(with_parties),
                'with_parties_raw': with_parties,
            })
            sums['opening'] += opening
            sums['issued'] += st['issued']
            sums['returned_good'] += returned_good
            sums['damaged'] += st['damaged']
            sums['available'] += available
            sums['with_parties'] += with_parties

        rows.sort(key=lambda r: r['item_description'].lower())
        for n, r in enumerate(rows, 1):
            r['s_no'] = n
        totals = {k: fmt_qty(sums[k]) for k in totals}

    except Exception:
        logger.exception('Error loading godown stock print')

    return render_template(
        'stock/godown_print.html',
        rows=rows,
        totals=totals,
        search=search,
        f_category=f_category,
        f_stock=f_stock,
        printed_on=datetime.now().strftime('%d-%m-%Y %H:%M')
    )