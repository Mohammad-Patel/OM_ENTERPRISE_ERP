"""
Settings Route - OM Enterprise ERP
----------------------------------
Place this file at:  app/routes/settings.py

Reads / saves the system settings (invoice numbering, company profile, GST rates, bank
details) in the "Settings" sheet of data/om_enterprise.xlsx.

Fixes over the old version
    * only the known fields are saved (the hidden csrf_token is no longer written to Excel)
    * values are validated (numbers, GSTIN, financial year ...)
    * saved through save_sheets() under excel_lock -> atomic, same as every other module
    * keys are matched case-insensitively (GST_INVOICE_NEXT_NO == gst_invoice_next_no)
"""
import logging
import re
from datetime import datetime

import pandas as pd
from flask import Blueprint, flash, redirect, render_template, request, url_for
from werkzeug.routing import BuildError

from app.services.numbering_service import current_series_values
from app.routes.challan_store import (
    Book, ChallanError, append_rows, audit_update, excel_lock, next_id, records, save_sheets,
)

logger = logging.getLogger(__name__)

settings_bp = Blueprint('settings', __name__, url_prefix='/settings')

SETTINGS_SHEET = 'Settings'
GSTIN_RE = re.compile(r'^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z]$')

# key -> (default, kind, label, description)
#   kind: 'int' positive whole number | 'rate' 0-100 | 'gstin' | 'fy' (26-27 or blank) | 'text'
FIELDS = {
    'GST_INVOICE_NEXT_NO': ('130', 'int', 'GST invoice next number', 'First number of the GST invoice series'),
    'NON_GST_INVOICE_NEXT_NO': ('150', 'int', 'Non-GST invoice next number', 'First number of the non-GST invoice series'),
    'FINANCIAL_YEAR': ('', 'fy', 'Financial year', 'Invoice suffix like 26-27. Blank = automatic (April - March)'),
    'GST_ISSUE_CHALLAN_PREFIX': ('OM/26-27/', 'text', 'GST issue challan prefix', 'GST issue challan = prefix + number'),
    # the 4 challan series: the value is the number the NEXT challan gets (counts up by itself)
    'GST_ISSUE_CHALLAN_NEXT_NO': ('', 'series', 'GST issue challan next no', 'Next GST issue challan number, e.g. OM/26-27/89'),
    'WITHOUT_GST_ISSUE_CHALLAN_NEXT_NO': ('', 'series', 'Without-GST issue challan next no', 'Next without-GST issue challan number'),
    'GST_RECEIVE_CHALLAN_NEXT_NO': ('', 'series', 'GST receive challan next no', 'Next GST receive challan number'),
    'WITHOUT_GST_RECEIVE_CHALLAN_NEXT_NO': ('', 'series', 'Without-GST receive challan next no', 'Next without-GST receive challan number'),
    'CGST_RATE': ('9', 'rate', 'CGST %', 'Applied on GST invoices'),
    'SGST_RATE': ('9', 'rate', 'SGST %', 'Applied on GST invoices'),
    'COMPANY_NAME': ('OM Enterprise', 'text', 'Company name', 'Printed on invoices'),
    'COMPANY_GSTIN': ('', 'gstin', 'Company GSTIN', 'Printed on tax invoices'),
    'COMPANY_ADDRESS': ('Plot No. 1046, KIRC Road, Near Sindbad Hotel, Kalol, Gujarat', 'text', 'Company address', ''),
    'COMPANY_PHONE': ('+91 9824219138', 'text', 'Company phone', ''),
    'COMPANY_EMAIL': ('', 'text', 'Company e-mail', ''),
    'COMPANY_STATE': ('Gujarat', 'text', 'State', ''),
    'COMPANY_STATE_CODE': ('24', 'text', 'State code', ''),
    'DEFAULT_RENTAL_RATE': ('1.25', 'text', 'Default rental rate (fallback)', ''),
    'BANK_NAME': ('', 'text', 'Bank name', 'Printed on invoices'),
    'BANK_ACCOUNT_NO': ('', 'text', 'Bank account no.', 'Printed on invoices'),
    'BANK_IFSC': ('', 'text', 'IFSC code', 'Printed on invoices'),
    'BANK_BRANCH': ('', 'text', 'Bank branch', 'Printed on invoices'),
    'INVOICE_TERMS': ('', 'text', 'Terms & conditions', 'Printed at the bottom of invoices'),
}
MAX_TEXT = 500


def _home_url():
    for endpoint in ('dashboard.dashboard', 'dashboard.index'):
        try:
            return url_for(endpoint)
        except BuildError:
            continue
    return '/dashboard'


def _sheet_name(book):
    return book.find(exact=('settings',)) or SETTINGS_SHEET


def _current_values(book):
    """{KEY: value} for every known field (saved value, else default)."""
    saved = {}
    for r in records(book.df(_sheet_name(book))):
        key = str(r.get('Setting_Key') or '').strip().upper()
        if key:
            saved[key] = str(r.get('Setting_Value') if r.get('Setting_Value') is not None else '').strip()
    values = {k: saved.get(k, spec[0]) for k, spec in FIELDS.items()}
    values.update(current_series_values(book))      # the number each series will give next, right now
    return values


def _validate(form):
    """-> (clean_values, errors)"""
    clean, errors = {}, []
    for key, (default, kind, label, _desc) in FIELDS.items():
        if key not in form:                       # field not on the form -> leave unchanged
            continue
        val = str(form.get(key, '')).strip()[:MAX_TEXT]
        if kind == 'series':
            # untouched field (same as what the page showed) -> leave the live counter alone, so a stale
            # tab can never move a series backwards
            if val == str(form.get('ORIG_' + key, '')).strip():
                continue
            m = re.fullmatch(r'(.*?)(\d{1,9})', val)
            if not m or int(m.group(2)) < 1:
                errors.append(f'{label}: enter the number the next challan should get (1 or more).')
                continue
            if key == 'GST_ISSUE_CHALLAN_NEXT_NO' and m.group(1).strip():
                clean['GST_ISSUE_CHALLAN_PREFIX'] = m.group(1).strip()   # OM/26-27/89 -> prefix + 89
            clean[key] = str(int(m.group(2)))
            continue
        if kind == 'int':
            if not re.fullmatch(r'\d{1,9}', val) or int(val) < 1:
                errors.append(f'{label}: enter a whole number of 1 or more.')
            else:
                val = str(int(val))
        elif kind == 'rate':
            try:
                if not 0 <= float(val) <= 100:
                    raise ValueError
                val = str(float(val)).rstrip('0').rstrip('.')
            except ValueError:
                errors.append(f'{label}: enter a number between 0 and 100.')
        elif kind == 'gstin':
            val = re.sub(r'\s+', '', val).upper()
            if val and not GSTIN_RE.match(val):
                errors.append(f'{label}: "{val}" is not a valid 15-character GSTIN.')
        elif kind == 'fy':
            if val and not re.fullmatch(r'\d{2}-\d{2}', val):
                errors.append(f'{label}: use the format 26-27 (or leave blank).')
        elif key == 'COMPANY_NAME' and not val:
            errors.append('Company name is required.')
        clean[key] = val
    return clean, errors


def _apply(book, clean):
    """Merge `clean` into the Settings sheet. Returns the updates dict for save_sheets()."""
    name = _sheet_name(book)
    df = book.df(name)
    if df.empty or 'Setting_Key' not in df.columns:
        df = pd.DataFrame(columns=['Setting_Key', 'Setting_Value', 'Description'])
    for col in ('Setting_Value', 'Description'):
        if col not in df.columns:
            df[col] = ''
        df[col] = df[col].astype(object)

    keys_lower = df['Setting_Key'].astype(str).str.strip().str.upper()
    new_rows = []
    for key, value in clean.items():
        match = df.index[keys_lower == key]
        if len(match):
            df.at[match[0], 'Setting_Value'] = value
        else:
            new_rows.append({'Setting_Key': key, 'Setting_Value': value, 'Description': FIELDS[key][3]})
    if new_rows:
        df = append_rows(df, new_rows)

    audit_name, audit_df = audit_update(book, 'UPDATE', 'Settings', after=clean, entity_type='Settings')
    return {name: df, audit_name: audit_df}


@settings_bp.route('/', methods=['GET', 'POST'])
def manage_settings():
    try:
        book = Book.load()
        if not book.sheets:
            raise ChallanError('Database file (data/om_enterprise.xlsx) not found.')
    except ChallanError as e:
        flash(str(e), 'danger')
        return redirect(_home_url())
    except Exception:
        logger.exception('Error reading settings')
        flash('Could not read the Excel data file.', 'danger')
        return redirect(_home_url())

    if request.method == 'POST':
        clean, errors = _validate(request.form)
        if not errors:
            try:
                with excel_lock:
                    book = Book.load()
                    save_sheets(_apply(book, clean))
                flash('Settings updated successfully.', 'success')
                return redirect(url_for('settings.manage_settings'))
            except PermissionError:
                errors.append('Excel file is open in another program. Please close it and try again.')
            except Exception:
                logger.exception('Error saving settings')
                errors.append('Could not save the settings. Please try again.')
        values = _current_values(book)
        values.update({k: v for k, v in request.form.items() if k in FIELDS})
        return render_template('settings/index.html', settings=values, errors=errors,
                               home_url=_home_url()), 400

    return render_template('settings/index.html', settings=_current_values(book), errors=[],
                           home_url=_home_url())
