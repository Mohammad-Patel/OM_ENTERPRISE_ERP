"""
Audit Logs - OM Enterprise ERP
------------------------------
Place this file at:  app/routes/audit_logs.py
Register it in app/__init__.py:   register_module('audit_logs', 'audit_logs_bp')

    /audit-logs/          list (search, action / entity / user / date filters, pages)
    /audit-logs/export    the same filtered list as a CSV file (opens in Excel)

Reads the "Audit_Log" sheet of data/om_enterprise.xlsx. Columns written by the ERP:
    Audit_ID, Event_Time, Username, Action, Entity_Type, Entity_ID, Reason, Before_Data, After_Data
Older / other column names (Timestamp, User, Details ...) are understood too.
"""
import csv
import io
import json
import logging
import math
import re
from datetime import date, datetime

from flask import Blueprint, Response, render_template, request, url_for
from werkzeug.routing import BuildError

from app.routes.challan_store import ALLOWED_PER_PAGE, Book, parse_any_date, records, to_int

logger = logging.getLogger(__name__)

audit_logs_bp = Blueprint('audit_logs', __name__, url_prefix='/audit-logs')

SHEET_NAMES = ('audit_log', 'audit_logs', 'auditlog', 'auditlogs')
TIME_FORMATS = ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M:%S.%f', '%Y-%m-%d %H:%M', '%d-%m-%Y %H:%M:%S',
                '%d-%m-%Y %H:%M', '%d/%m/%Y %H:%M:%S', '%Y-%m-%dT%H:%M:%S', '%d-%m-%Y', '%Y-%m-%d', '%d/%m/%Y')
IMPORTANT_KEYS = ('Bill_No', 'Challan_No', 'Receive_Challan_No', 'Entry_No', 'Party_ID', 'Net_Total',
                  'Quantity', 'Status', 'Taxable_Value')
MAX_VALUE_CHARS = 220


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------
def _home_url():
    for endpoint in ('dashboard.dashboard', 'dashboard.index'):
        try:
            return url_for(endpoint)
        except BuildError:
            continue
    return '/'


def _key(name):
    return re.sub(r'[^a-z0-9]', '', str(name).lower())


def _pick(low, *names):
    for n in names:
        v = low.get(_key(n))
        if v not in (None, ''):
            return v
    return ''


def _parse_dt(value):
    s = str(value or '').strip()
    if not s:
        return None
    for fmt in TIME_FORMATS:
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    d = parse_any_date(s)
    return datetime(d.year, d.month, d.day) if d else None


def _json(value):
    s = str(value if value is not None else '').strip()
    if not s:
        return None
    try:
        return json.loads(s)
    except (ValueError, TypeError):
        return s


def _short(v, limit=MAX_VALUE_CHARS):
    s = '' if v is None else (json.dumps(v, ensure_ascii=False) if isinstance(v, (dict, list)) else str(v))
    return s if len(s) <= limit else s[:limit] + '...'


def _changes(before, after):
    """[{'field', 'before', 'after', 'status'}] - what this audit entry recorded."""
    b = before if isinstance(before, dict) else None
    a = after if isinstance(after, dict) else None
    out = []
    if b is not None and a is not None:                              # UPDATE: only what changed
        for k in list(a.keys()) + [k for k in b if k not in a]:
            bv, av = b.get(k), a.get(k)
            if k not in b:
                out.append({'field': k, 'before': '', 'after': _short(av), 'status': 'new'})
            elif k not in a:
                out.append({'field': k, 'before': _short(bv), 'after': '', 'status': 'removed'})
            elif str(bv) != str(av):
                out.append({'field': k, 'before': _short(bv), 'after': _short(av), 'status': 'changed'})
    elif a is not None:                                              # CREATE: the new values
        out = [{'field': k, 'before': '', 'after': _short(v), 'status': 'new'} for k, v in a.items()]
    elif b is not None:                                              # CANCEL / DELETE: values before
        out = [{'field': k, 'before': _short(v), 'after': '', 'status': 'old'} for k, v in b.items()]
    else:                                                            # plain text
        if before not in (None, ''):
            out.append({'field': 'Before', 'before': _short(before), 'after': '', 'status': 'old'})
        if after not in (None, ''):
            out.append({'field': 'After', 'before': '', 'after': _short(after), 'status': 'new'})
    return out


def _summary(changes):
    changed = [c for c in changes if c['status'] == 'changed']
    if changed:
        parts = [f"{c['field']}: {c['before'][:40]} \u2192 {c['after'][:40]}" for c in changed[:2]]
        more = len(changed) - 2
        return '; '.join(parts) + (f'  (+{more} more)' if more > 0 else '')
    picked = [c for c in changes if c['field'] in IMPORTANT_KEYS]
    picked.sort(key=lambda c: IMPORTANT_KEYS.index(c['field']))
    if picked:
        return ' | '.join(f"{c['field']}: {(c['after'] or c['before'])[:40]}" for c in picked[:4])
    return ''


def _action_class(action):
    a = action.upper()
    if any(w in a for w in ('ERROR', 'FAIL', 'CANCEL', 'DELETE', 'REMOVE')):
        return 'b-danger'
    if any(w in a for w in ('WARN',)):
        return 'b-warn'
    if any(w in a for w in ('CREATE', 'ADD', 'NEW', 'INSERT', 'SAVE')):
        return 'b-create'
    if any(w in a for w in ('UPDATE', 'EDIT', 'MODIFY', 'CHANGE')):
        return 'b-update'
    return 'b-other'


def _normalise(row, pos):
    low = {_key(k): v for k, v in row.items()}
    dt = _parse_dt(_pick(low, 'Event_Time', 'Timestamp', 'Created_At', 'Date_Time', 'DateTime', 'Date', 'Time'))
    before = _json(_pick(low, 'Before_Data', 'Before', 'Old_Value', 'Old_Data'))
    after = _json(_pick(low, 'After_Data', 'After', 'New_Value', 'New_Data'))
    action = str(_pick(low, 'Action', 'Event', 'Event_Type', 'Type') or '').strip() or '-'
    changes = _changes(before, after)
    entry = {
        'id': str(_pick(low, 'Audit_ID', 'ID', 'Log_ID') or pos),
        'pos': pos,
        'dt': dt,
        'time': dt.strftime('%d-%m-%Y %H:%M:%S') if dt else str(_pick(low, 'Event_Time', 'Timestamp') or '-'),
        'user': str(_pick(low, 'Username', 'User', 'User_Name', 'Created_By') or '').strip(),
        'action': action,
        'action_class': _action_class(action),
        'entity_type': str(_pick(low, 'Entity_Type', 'Module', 'Entity', 'Table') or '').strip(),
        'entity_id': str(_pick(low, 'Entity_ID', 'Reference', 'Reference_ID', 'Record_ID') or '').strip(),
        'reason': str(_pick(low, 'Reason', 'Details', 'Description', 'Message', 'Remarks') or '').strip(),
        'changes': changes,
        'summary': _summary(changes),
        'before_raw': '' if before is None else _short(before, 4000),
        'after_raw': '' if after is None else _short(after, 4000),
    }
    entry['hay'] = ' '.join([entry['time'], entry['user'], entry['action'], entry['entity_type'], entry['entity_id'],
                             entry['reason'], entry['summary'], entry['before_raw'], entry['after_raw']]).lower()
    return entry


def _load_all():
    """-> (entries newest first, diag dict)."""
    diag = {'state': 'ok', 'rows': 0, 'message': ''}
    try:
        book = Book.load()
    except Exception:
        logger.exception('Could not read the Excel file for audit logs')
        return [], {'state': 'error', 'rows': 0,
                    'message': 'The Excel data file could not be read. Please check that it exists and is not corrupted.'}
    if not book.sheets:
        return [], {'state': 'nofile', 'rows': 0, 'message': 'The Excel data file (data/om_enterprise.xlsx) was not found.'}
    name = book.find(exact=SHEET_NAMES)
    if not name:
        return [], {'state': 'nosheet', 'rows': 0,
                    'message': 'The Excel file has no "Audit_Log" sheet yet. It is created automatically the first time '
                               'a module saves an audit entry (e.g. creating or cancelling an invoice).'}
    rows = records(book.df(name))
    diag['rows'] = len(rows)
    if not rows:
        diag.update(state='empty', message='The Audit_Log sheet exists but has no rows. No module has written an audit '
                                           'entry yet - make a change (create / edit / cancel a record) and check again.')
    entries = [_normalise(r, i + 1) for i, r in enumerate(rows)]
    entries.sort(key=lambda e: (e['dt'] or datetime.min, to_int(e['id'], 0), e['pos']), reverse=True)
    return entries, diag


def _args():
    per_page = request.args.get('per_page', '25')
    if per_page not in ALLOWED_PER_PAGE:
        per_page = '25'
    return {
        'search': request.args.get('search', '').strip(),
        'action': request.args.get('action', '').strip(),
        'entity': request.args.get('entity', '').strip(),
        'user': request.args.get('user', '').strip(),
        'date_from': parse_any_date(request.args.get('date_from')),
        'date_to': parse_any_date(request.args.get('date_to')),
        'per_page': per_page,
        'page': max(to_int(request.args.get('page'), 1), 1),
    }


def _filter(entries, a):
    q = a['search'].lower()
    out = []
    for e in entries:
        if a['action'] and e['action'] != a['action']:
            continue
        if a['entity'] and e['entity_type'] != a['entity']:
            continue
        if a['user'] and (e['user'] or '-') != a['user']:
            continue
        if a['date_from'] or a['date_to']:
            d = e['dt'].date() if e['dt'] else None
            if d is None or (a['date_from'] and d < a['date_from']) or (a['date_to'] and d > a['date_to']):
                continue
        if q and q not in e['hay']:
            continue
        out.append(e)
    return out


def _page_numbers(page, pages):
    wanted = sorted({n for n in (1, pages, page - 2, page - 1, page, page + 1, page + 2) if 1 <= n <= pages})
    out, prev = [], 0
    for n in wanted:
        if prev and n - prev > 1:
            out.append(None)
        out.append(n)
        prev = n
    return out


def _qargs(a):
    return {k: v for k, v in {
        'search': a['search'], 'action': a['action'], 'entity': a['entity'], 'user': a['user'],
        'date_from': a['date_from'].isoformat() if a['date_from'] else '',
        'date_to': a['date_to'].isoformat() if a['date_to'] else '',
        'per_page': a['per_page'],
    }.items() if v}


# --------------------------------------------------------------------------
# Routes
# --------------------------------------------------------------------------
@audit_logs_bp.route('/', methods=['GET'])
def index():
    a = _args()
    entries, diag = _load_all()
    filtered = _filter(entries, a)

    total = len(filtered)
    if a['per_page'] == 'All':
        pages, page, rows = 1, 1, filtered
    else:
        size = int(a['per_page'])
        pages = max(math.ceil(total / size), 1)
        page = min(a['page'], pages)
        rows = filtered[(page - 1) * size: page * size]

    today = date.today()
    cards = {
        'total': len(entries),
        'today': sum(1 for e in entries if e['dt'] and e['dt'].date() == today),
        'last': entries[0]['time'] if entries else '-',
    }
    qargs = _qargs(a)
    return render_template(
        'audit_logs/index.html', logs=rows, diag=diag, cards=cards, total_records=total,
        actions=sorted({e['action'] for e in entries}),
        entities=sorted({e['entity_type'] for e in entries if e['entity_type']}),
        users=sorted({(e['user'] or '-') for e in entries}),
        search_query=a['search'], action_filter=a['action'], entity_filter=a['entity'], user_filter=a['user'],
        date_from=qargs.get('date_from', ''), date_to=qargs.get('date_to', ''),
        qargs=qargs, per_page=a['per_page'], per_page_options=ALLOWED_PER_PAGE,
        page=page, total_pages=pages, page_numbers=_page_numbers(page, pages),
        filters_active=bool(a['search'] or a['action'] or a['entity'] or a['user'] or a['date_from'] or a['date_to']),
        home_url=_home_url())


def _csv_cell(v):
    s = '' if v is None else str(v)
    return "'" + s if s[:1] in ('=', '+', '-', '@') else s        # stop Excel from running formulas


@audit_logs_bp.route('/export', methods=['GET'])
def export():
    a = _args()
    entries, _diag = _load_all()
    rows = _filter(entries, a)
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(['Audit ID', 'Time', 'User', 'Action', 'Entity Type', 'Entity ID', 'Reason', 'Before', 'After'])
    for e in rows:
        w.writerow([_csv_cell(x) for x in (e['id'], e['time'], e['user'] or '-', e['action'], e['entity_type'],
                                          e['entity_id'], e['reason'], e['before_raw'], e['after_raw'])])
    name = f"audit_logs_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    return Response('\ufeff' + buf.getvalue(), mimetype='text/csv; charset=utf-8',
                    headers={'Content-Disposition': f'attachment; filename={name}'})
