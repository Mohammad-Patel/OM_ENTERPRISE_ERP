"""
Monthly Billing routes - OM Enterprise ERP
------------------------------------------
Place this file at:  app/routes/monthly_billing.py

    /monthly-billing/                 list of saved invoices (search / filter / pages)
    /monthly-billing/select           choose party, site and period
    /monthly-billing/process  (POST)  -> redirects to /new
    /monthly-billing/new              auto-calculated draft in the editor
    /monthly-billing/save     (POST)  save a new invoice
    /monthly-billing/edit/<no>        edit a saved invoice
    /monthly-billing/update/<no> (POST)
    /monthly-billing/cancel/<no> (POST)
    /monthly-billing/print/<no>       A4 portrait print view
"""
import logging
import math
from datetime import date, timedelta

from flask import Blueprint, flash, redirect, render_template, request, url_for
from werkzeug.routing import BuildError

from app.routes.challan_store import (
    ALLOWED_PER_PAGE, Book, ChallanError, excel_lock, is_active, item_label, load_items_master,
    load_parties, load_sites, parse_any_date, save_sheets, sid, to_float, to_int, to_js_json,
    log_audit
)
from app.services import billing_service as svc

logger = logging.getLogger(__name__)

monthly_billing_bp = Blueprint('monthly_billing', __name__, url_prefix='/monthly-billing')


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------
def _home_url():
    for endpoint in ('dashboard.dashboard', 'dashboard.index'):
        try:
            return url_for(endpoint)
        except BuildError:
            continue
    return '/dashboard'


def _items_for_editor(book):
    out = [{'id': sid(i.get('Item_ID')), 'name': item_label(i),
            'rate': to_float(i.get('Daily_Rental_Rate'))}
           for i in load_items_master(book) if is_active(i)]
    out.sort(key=lambda x: x['name'].lower())
    return out


def _render_editor(book, draft, error=None, status=200):
    return render_template(
        'monthly_billing/invoice_edit.html', draft=draft, error=error,
        items_json=to_js_json(_items_for_editor(book)),
        rows_json=to_js_json(draft['rows']),
        company=svc.company_info(book), home_url=_home_url()), status


def _month_bounds(today=None):
    today = today or date.today()
    first = today.replace(day=1)
    next_first = date(first.year + 1, 1, 1) if first.month == 12 else date(first.year, first.month + 1, 1)
    return first, next_first - timedelta(days=1)


# --------------------------------------------------------------------------
# List
# --------------------------------------------------------------------------
@monthly_billing_bp.route('/', methods=['GET'])
def index():
    search_query = request.args.get('search', '').strip()
    party_filter = request.args.get('party_id', '').strip()
    site_filter = request.args.get('site_id', '').strip()  # <-- Yeh add kiya hai
    status_filter = request.args.get('status', '').strip()
    gst_filter = request.args.get('gst_filter', '').strip()
    date_from = parse_any_date(request.args.get('date_from'))
    date_to = parse_any_date(request.args.get('date_to'))
    per_page = request.args.get('per_page', '10')
    if per_page not in ALLOWED_PER_PAGE:
        per_page = '10'
    page = max(to_int(request.args.get('page'), 1), 1)

    parties, sites, bills = [], [], []
    cards = {'active': 0, 'cancelled': 0, 'billed': '0.00'}
    try:
        book = Book.load()
        if not book.sheets:
            raise ChallanError('Excel database not found or empty.')
        parties = [p for p in load_parties(book) if is_active(p)]
        
        # --- Sites load karne ke liye yeh add kiya hai ---
        sites = [{'id': sid(s.get('Site_ID')), 'party_id': sid(s.get('Party_ID')),
                  'name': s.get('Site_Name') or s.get('Site_Address') or sid(s.get('Site_ID'))}
                 for s in load_sites(book) if is_active(s)]
                 
        bills, cards = svc.list_bills(book)
    except Exception:
        logger.exception('Error loading monthly bills')
        flash('Could not read the Excel data file. Please check that it exists and is not corrupted.', 'danger')

    filtered = bills
    if party_filter:
        filtered = [b for b in filtered if sid(b.get('Party_ID')) == sid(party_filter)]
    if site_filter:  # <-- Site filtering logic add ki hai
        filtered = [b for b in filtered if sid(b.get('Site_ID')) == sid(site_filter)]
    if status_filter == 'Active':
        filtered = [b for b in filtered if not b['Cancelled']]
    elif status_filter == 'Cancelled':
        filtered = [b for b in filtered if b['Cancelled']]
    if gst_filter == 'gst':
        filtered = [b for b in filtered if b['Has_GST']]
    elif gst_filter == 'nogst':
        filtered = [b for b in filtered if not b['Has_GST']]
    if date_from or date_to:
        def in_range(b):
            d = parse_any_date(b.get('Bill_Date'))
            return d is not None and (not date_from or d >= date_from) and (not date_to or d <= date_to)
        filtered = [b for b in filtered if in_range(b)]
    if search_query:
        q = search_query.lower()
        filtered = [b for b in filtered
                    if any(q in str(b.get(f, '')).lower() for f in ('Bill_No', 'Party_Name', 'Site_Name'))]

    total_records = len(filtered)
    if per_page == 'All':
        total_pages, page, rows = 1, 1, filtered
    else:
        size = int(per_page)
        total_pages = max(math.ceil(total_records / size), 1)
        page = min(page, total_pages)
        rows = filtered[(page - 1) * size: page * size]

    df_iso = date_from.isoformat() if date_from else ''
    dt_iso = date_to.isoformat() if date_to else ''
    qargs = {k: v for k, v in {'search': search_query, 'party_id': party_filter, 'site_id': site_filter, 'status': status_filter,
                               'gst_filter': gst_filter, 'date_from': df_iso, 'date_to': dt_iso,
                               'per_page': per_page}.items() if v}
    return render_template(
        'monthly_billing/index.html', bills=rows, parties=parties, sites_json=to_js_json(sites), cards=cards,
        search_query=search_query, party_filter=party_filter, site_filter=site_filter, status_filter=status_filter,
        gst_filter=gst_filter, date_from=df_iso, date_to=dt_iso, qargs=qargs,
        filters_active=bool(search_query or party_filter or site_filter or status_filter or gst_filter or date_from or date_to),
        per_page=per_page, per_page_options=ALLOWED_PER_PAGE, page=page, total_pages=total_pages,
        total_records=total_records, home_url=_home_url())

# --------------------------------------------------------------------------
# Selection -> draft
# --------------------------------------------------------------------------
@monthly_billing_bp.route('/select', methods=['GET'])
def generate():
    parties, sites = [], []
    try:
        book = Book.load()
        parties = [p for p in load_parties(book) if is_active(p)]
        sites = [{'id': sid(s.get('Site_ID')), 'party_id': sid(s.get('Party_ID')),
                  'name': s.get('Site_Name') or s.get('Site_Address') or sid(s.get('Site_ID'))}
                 for s in load_sites(book) if is_active(s)]
    except Exception:
        logger.exception('Error loading parties / sites')
        flash('Could not read the Excel data file.', 'danger')

    first, last = _month_bounds()
    date_from = request.args.get('date_from') or first.isoformat()
    return render_template(
        'monthly_billing/generate.html', parties=parties, sites_json=to_js_json(sites),
        party_id=request.args.get('party_id', ''), site_id=request.args.get('site_id', ''),
        date_from=date_from, date_to=request.args.get('date_to') or last.isoformat(),
        month=date_from[:7], home_url=_home_url())


@monthly_billing_bp.route('/process', methods=['POST'])
def process_billing():
    """Forwards the posted selection to the draft editor."""
    return redirect(url_for('monthly_billing.new',
                            party_id=request.form.get('party_id', ''),
                            site_id=request.form.get('site_id', ''),
                            date_from=request.form.get('date_from', ''),
                            date_to=request.form.get('date_to', '')), code=303)


@monthly_billing_bp.route('/new', methods=['GET'])
def new():
    party_id = request.args.get('party_id', '').strip()
    site_id = request.args.get('site_id', '').strip()
    raw_from, raw_to = request.args.get('date_from', ''), request.args.get('date_to', '')
    back = dict(party_id=party_id, site_id=site_id, date_from=raw_from, date_to=raw_to)
    try:
        if not party_id or not site_id:
            raise ChallanError('Please select a party and a site.')
        start, end = parse_any_date(raw_from), parse_any_date(raw_to)
        book = Book.load()
        if not book.sheets:
            raise ChallanError('Excel data file not found.')
        draft = svc.build_draft(book, party_id, site_id, start, end)
    except ChallanError as e:
        flash(str(e), 'danger')
        return redirect(url_for('monthly_billing.generate', **back))
    except Exception:
        logger.exception('Error building monthly bill draft')
        flash('Could not calculate the bill. Please check the data file and try again.', 'danger')
        return redirect(url_for('monthly_billing.generate', **back))
    return _render_editor(book, draft)[0]


# --------------------------------------------------------------------------
# Save / update / cancel
# --------------------------------------------------------------------------
def _persist(mode, bill_no=None):
    form = request.form
    error = None
    try:
        with excel_lock:
            book = Book.load()
            if not book.sheets:
                raise ChallanError('Excel data file not found.')
            payload = svc.parse_payload(form, book, bill_no if mode == 'edit' else None)
            if mode == 'new':
                updates, bill_no = svc.create_bill(book, payload)
                save_sheets(updates)
                log_audit(book, 'CREATE_INVOICE', f'Created monthly invoice {bill_no}')
            else:
                updates = svc.update_bill(book, bill_no, payload)
                save_sheets(updates)
                log_audit(book, 'UPDATE_INVOICE', f'Updated monthly invoice {bill_no}')
    except ChallanError as e:
        error = str(e)
        with excel_lock:
            book = Book.load()
            log_audit(book, 'ERROR', f'Challan Validation Error: {error}')
    except PermissionError:
        error = 'Excel file is open in another program. Please close it and press Save again.'
        with excel_lock:
            book = Book.load()
            log_audit(book, 'ERROR', 'PermissionError: Excel file locked/open in another program')
    except Exception as ex:
        logger.exception('Error saving monthly bill')
        error = 'Could not save the invoice. Please try again.'
        with excel_lock:
            book = Book.load()
            log_audit(book, 'ERROR', f'System Exception during save: {str(ex)}')

    if error:
        try:
            book = Book.load()
            draft = svc.draft_from_form(book, form, bill_no if mode == 'edit' else None)
            return _render_editor(book, draft, error=error, status=400)
        except Exception:
            logger.exception('Error re-showing the invoice editor')
            flash(error, 'danger')
            return redirect(url_for('monthly_billing.generate'))

    flash(f'Invoice {bill_no} saved successfully.', 'success')
    if form.get('after') == 'print':
        return redirect(url_for('monthly_billing.print_invoice', bill_no=bill_no))
    return redirect(url_for('monthly_billing.index'))


@monthly_billing_bp.route('/save', methods=['POST'])
def save_invoice():
    return _persist('new')


@monthly_billing_bp.route('/edit/<path:bill_no>', methods=['GET'])
def edit_invoice(bill_no):
    try:
        book = Book.load()
        draft = svc.load_bill(book, bill_no)
    except ChallanError as e:
        flash(str(e), 'danger')
        return redirect(url_for('monthly_billing.index'))
    except Exception:
        logger.exception('Error loading invoice')
        flash('Could not load the invoice.', 'danger')
        return redirect(url_for('monthly_billing.index'))
    if str(draft.get('status', '')).lower() in ('cancelled', 'canceled'):
        flash('A cancelled invoice cannot be edited.', 'warning')
        return redirect(url_for('monthly_billing.index'))
    return _render_editor(book, draft)[0]


@monthly_billing_bp.route('/update/<path:bill_no>', methods=['POST'])
def update_invoice(bill_no):
    return _persist('edit', bill_no)


@monthly_billing_bp.route('/cancel/<path:bill_no>', methods=['POST'])
def cancel_invoice(bill_no):
    try:
        with excel_lock:
            book = Book.load()
            save_sheets(svc.cancel_bill(book, bill_no, request.form.get('reason')))
            log_audit(book, 'CANCEL_INVOICE', f'Cancelled monthly invoice {bill_no}')
        flash(f'Invoice {bill_no} cancelled.', 'success')
    except ChallanError as e:
        flash(str(e), 'danger')
    except PermissionError:
        flash('Excel file is open in another program. Please close it and try again.', 'danger')
    except Exception:
        logger.exception('Error cancelling invoice')
        flash('Could not cancel the invoice.', 'danger')
    return redirect(url_for('monthly_billing.index'))


# --------------------------------------------------------------------------
# Print
# --------------------------------------------------------------------------
@monthly_billing_bp.route('/print/<path:bill_no>', methods=['GET'])
def print_invoice(bill_no):
    inv, company = None, svc.DEFAULT_COMPANY
    try:
        book = Book.load()
        company = svc.company_info(book)
        inv = svc.print_context(book, bill_no)
    except ChallanError:
        inv = None
    except Exception:
        logger.exception('Error loading invoice for print')
        inv = None
    return (render_template('monthly_billing/print.html', inv=inv, company=company, bill_no=bill_no),
            200 if inv else 404)


@monthly_billing_bp.route('/delete/<path:bill_no>', methods=['POST'])
def delete_invoice(bill_no):
    try:
        with excel_lock:
            book = Book.load()
            if not book.sheets:
                raise ChallanError('Excel data file not found.')
            updates = svc.delete_bill_permanently(book, bill_no)
            save_sheets(updates)
            log_audit(book, 'DELETE_INVOICE', f'Permanently deleted monthly invoice {bill_no}')
        flash(f'Invoice {bill_no} deleted permanently.', 'success')
    except ChallanError as e:
        flash(str(e), 'danger')
    except PermissionError:
        flash('Excel file is open in another program. Please close it and try again.', 'danger')
    except Exception:
        logger.exception('Error deleting invoice permanently')
        flash('Could not delete the invoice.', 'danger')
    return redirect(url_for('monthly_billing.index'))