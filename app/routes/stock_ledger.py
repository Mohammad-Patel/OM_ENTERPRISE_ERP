"""
Stock Ledger - party & site wise stock.

    outstanding (stock with party) = opening + issued - returned

Pages:
  /stock-ledger/        item-wise balance for one party + site, plus movement history
  /stock-ledger/print   printable stock statement
Numbers come from challan_store.stock_summary(), the same function the godown page uses.
"""
import logging
from datetime import date

from flask import Blueprint, flash, redirect, render_template, request, url_for
from werkzeug.routing import BuildError

from app.routes.challan_store import (
    Book,
    _IGNORED_OPENING_STATUS,
    _doc_keys,
    _opening_qty,
    _returned_split,
    ISSUE_KEYS,
    RECEIVE_KEYS,
    fmt_qty,
    is_cancelled,
    item_label,
    load_items_master,
    load_parties,
    load_sites,
    parse_any_date,
    receive_items_sheet,
    receive_sheet,
    records,
    safe_action,
    sid,
    stock_summary,
    to_float,
    to_js_json,
)

logger = logging.getLogger(__name__)

stock_ledger_bp = Blueprint('stock_ledger', __name__, url_prefix='/stock-ledger')

HOME_ENDPOINT = 'dashboard.index'      # page to show when something fails
EPS = 1e-9
TYPE_ORDER = {'Opening': 0, 'Issue': 1, 'Receive': 2}


# ----------------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------------
def _first(row, *keys):
    for k in keys:
        s = str(row.get(k, '') if row.get(k) is not None else '').strip()
        if s and s.lower() != 'nan':
            return s
    return ''


def _url(endpoint, **values):
    try:
        return url_for(endpoint, **values)
    except BuildError:
        return None


def _fmt_date(d):
    return d.strftime('%d-%m-%Y') if d else '-'


def _party_name(p):
    return p.get('Party_Name') or p.get('Name') or ''


def _stock_rows(book, party_id, site_id, items):
    """Item-wise rows + totals for one party/site."""
    rows = []
    for iid, d in stock_summary(book, party_id, site_id).items():
        if not any(abs(d[k]) > EPS for k in ('opening', 'issued', 'received', 'outstanding')):
            continue
        meta = items.get(iid)
        rows.append({
            'item_id': iid,
            'item_description': item_label(meta) if meta else f'Item #{iid}',
            'opening': fmt_qty(d['opening']),
            'issued': fmt_qty(d['issued']),
            'received': fmt_qty(d['received']),
            'outstanding': fmt_qty(d['outstanding']),
            'outstanding_raw': round(d['outstanding'], 3),
            '_d': d,
        })
    rows.sort(key=lambda r: r['item_description'].lower())
    for n, r in enumerate(rows, 1):
        r['s_no'] = n
    sums = {k: sum(r['_d'][k] for r in rows) for k in ('opening', 'issued', 'received', 'outstanding')}
    for r in rows:
        del r['_d']
    totals = {k: fmt_qty(v) for k, v in sums.items()}
    return rows, totals, sums['outstanding']


def _collect_docs(headers, item_rows, keys, no_cols, date_cols, kind, party_id, site_id):
    """Group item lines under their (non-cancelled) challan for one party + site."""
    docs, by_key = [], {}
    for h in headers:
        if is_cancelled(h) or sid(h.get('Party_ID')) != party_id or sid(h.get('Site_ID')) != site_id:
            continue
        doc = {'type': kind, 'no': _first(h, *no_cols), 'date': parse_any_date(_first(h, *date_cols)),
               'lines': [], 'header': h}
        docs.append(doc)
        for k in _doc_keys(h, keys):
            by_key[k] = doc
    for r in item_rows:
        for k in _doc_keys(r, keys):
            if k in by_key:
                by_key[k]['lines'].append(r)
                break
    return docs


def _movements(book, party_id, site_id, items):
    """Opening + every issue / return of this party/site with a running balance."""
    def label(iid):
        meta = items.get(iid)
        return item_label(meta) if meta else f'Item #{iid}'

    entries = []

    # opening stock (one combined line)
    op_lines, op_dates = [], []
    for r in records(book.df(book.find(exact=('openingstock',)))):
        if sid(r.get('Party_ID')) != party_id or sid(r.get('Site_ID')) != site_id:
            continue
        if str(r.get('Status', '')).strip().lower() in _IGNORED_OPENING_STATUS:
            continue
        q = _opening_qty(r)
        if q > 0:
            op_lines.append((label(sid(r.get('Item_ID'))), q))
        d = parse_any_date(_first(r, 'Verified_Date', 'Estimated_Date'))
        if d:
            op_dates.append(d)
    if op_lines:
        entries.append({'type': 'Opening', 'no': 'OPENING', 'date': min(op_dates) if op_dates else None,
                        'lines': op_lines, 'sign': 1, 'url': None, 'damaged': 0})

    # issues
    issue_docs = _collect_docs(
        records(book.df(book.find(exact=('issue_challan',)))),
        records(book.df(book.find(exact=('issue_challan_items',)))),
        ISSUE_KEYS, ('Challan_No', 'Issue_No'), ('Challan_Date', 'Issue_Date'), 'Issue', party_id, site_id)
    for d in issue_docs:
        lines = [(label(sid(r.get('Item_ID'))), to_float(r.get('Quantity'))) for r in d['lines']]
        lines = [x for x in lines if x[1] > 0]
        if lines:
            entries.append({'type': 'Issue', 'no': d['no'], 'date': d['date'], 'lines': lines, 'sign': 1,
                            'url': _url('issue_challan.print_challan', issue_id=sid(d['header'].get('Issue_ID'))),
                            'damaged': 0})

    # returns
    recv_docs = _collect_docs(
        records(book.df(receive_sheet(book))), records(book.df(receive_items_sheet(book))),
        RECEIVE_KEYS, ('Receive_No', 'Receive_Challan_No'), ('Receive_Date', 'Date'), 'Receive',
        party_id, site_id)
    for d in recv_docs:
        lines, damaged = [], 0.0
        for r in d['lines']:
            total, dmg, _good = _returned_split(r)
            if total > 0:
                lines.append((label(sid(r.get('Item_ID'))), total))
                damaged += dmg
        if lines:
            entries.append({'type': 'Receive', 'no': d['no'], 'date': d['date'], 'lines': lines, 'sign': -1,
                            'url': _url('receive_challan.print_challan', receive_no=d['no']), 'damaged': damaged})

    entries.sort(key=lambda e: (e['date'] or date.min, TYPE_ORDER[e['type']], e['no']))

    out, running = [], 0.0
    for e in entries:
        qty = sum(q for _l, q in e['lines'])
        running += e['sign'] * qty
        detail = ', '.join(f'{lbl} \u00d7 {fmt_qty(q)}' for lbl, q in e['lines'])
        if e['damaged'] > EPS:
            detail += f'  (damaged: {fmt_qty(e["damaged"])})'
        out.append({
            'date_display': _fmt_date(e['date']), 'type': e['type'], 'no': e['no'], 'url': e['url'],
            'detail': detail, 'sign': e['sign'],
            'qty_display': ('+' if e['sign'] > 0 else '-') + fmt_qty(qty),
            'balance': fmt_qty(running),
        })
    return out


def _validated_selection(parties, sites):
    """Read party/site from the URL and drop anything that does not exist / belong together."""
    party_id = sid(request.args.get('party_id', '').strip())
    site_id = sid(request.args.get('site_id', '').strip())
    p_map = {sid(p.get('Party_ID') or p.get('ID')): p for p in parties if p}
    s_map = {sid(s.get('Site_ID') or s.get('site_id')): s for s in sites if s}

    if party_id and party_id not in p_map:
        flash('Selected party was not found.', 'warning')
        party_id = site_id = ''
    if site_id:
        s = s_map.get(site_id)
        if not party_id or not s or sid(s.get('Party_ID') or s.get('party_id')) != party_id:
            flash('Selected site does not belong to this party.', 'warning')
            site_id = ''
    return party_id, site_id, p_map, s_map


# ----------------------------------------------------------------------------
# Routes
# ----------------------------------------------------------------------------
@stock_ledger_bp.route('/', methods=['GET'])
@safe_action('Could not load the stock ledger.', HOME_ENDPOINT)
def index():
    book = Book.load()
    parties, sites = load_parties(book), load_sites(book)
    items = {sid(i.get('Item_ID')): i for i in load_items_master(book)}
    party_id, site_id, p_map, s_map = _validated_selection(parties, sites)

    stock_rows, movements = [], []
    totals = {'opening': '0', 'issued': '0', 'received': '0', 'outstanding': '0'}
    if party_id and site_id:
        stock_rows, totals, _raw = _stock_rows(book, party_id, site_id, items)
        movements = _movements(book, party_id, site_id, items)

    selected_party_name = _party_name(p_map[party_id]) if party_id and site_id else ''
    selected_site_name = ''
    if party_id and site_id:
        s = s_map[site_id]
        selected_site_name = s.get('Site_Name') or s.get('Site_Address') or s.get('Address') or site_id

    return render_template(
        'stock_ledger/index.html',
        parties=parties, sites_json=to_js_json(sites),
        stock_rows=stock_rows, totals=totals, movements=movements,
        party_id=party_id, site_id=site_id,
        selected_party_name=selected_party_name, selected_site_name=selected_site_name,
    )


@stock_ledger_bp.route('/print', methods=['GET'])
@safe_action('Could not prepare the stock statement.', 'stock_ledger.index')
def print_ledger():
    book = Book.load()
    parties, sites = load_parties(book), load_sites(book)
    party_id, site_id, p_map, s_map = _validated_selection(parties, sites)
    if not party_id or not site_id:
        flash('Please select both Party and Site Location.', 'warning')
        return redirect(url_for('stock_ledger.index'))

    items = {sid(i.get('Item_ID')): i for i in load_items_master(book)}
    rows, totals, outstanding = _stock_rows(book, party_id, site_id, items)
    p, s = p_map[party_id], s_map[site_id]

    ledger = {
        'party_name': _party_name(p) or party_id,
        'party_address': p.get('Address') or p.get('Billing_Address') or 'Ahmedabad, Gujarat',
        'party_mobile': p.get('Mobile_No') or p.get('Phone') or '-',
        'party_gstin': p.get('GSTIN') or '-',
        'site_name': s.get('Site_Name') or s.get('Site_Address') or s.get('Address') or site_id,
        'total_outstanding': fmt_qty(outstanding),
        'totals': totals,
    }
    return render_template('stock_ledger/print.html', ledger=ledger, rows=rows)
