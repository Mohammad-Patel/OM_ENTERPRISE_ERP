from flask import Flask, redirect, url_for, Blueprint, request, session, render_template, flash
import os
import json
import shutil
from datetime import datetime

def create_app():
    app = Flask(__name__)
    app.config['SECRET_KEY'] = 'om_enterprise_secret_key_2026'

    # Config file path for password management
    CONFIG_FILE = os.path.join(app.root_path, '..', 'config.json')

    def load_config():
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r') as f:
                return json.load(f)
        return {"password": "om123", "security_question": "What is your business city?", "security_answer": "kalol"}


    def save_config(config_data):
        with open(CONFIG_FILE, 'w') as f:
            json.dump(config_data, f, indent=4)

    # --- AUTOMATED EXCEL BACKUP ON STARTUP ---
    try:
        backup_dir = os.path.join(app.root_path, '..', 'backups')
        if not os.path.exists(backup_dir):
            os.makedirs(backup_dir, exist_ok=True)
        
        root_dir = os.path.join(app.root_path, '..')
        for file in os.listdir(root_dir):
            if file.endswith('.xlsx') and 'backup' not in file:
                src_path = os.path.join(root_dir, file)
                timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
                backup_filename = f"{os.path.splitext(file)[0]}_backup_{timestamp}.xlsx"
                shutil.copy(src_path, os.path.join(backup_dir, backup_filename))
        print("Excel backup created successfully on startup.")
    except Exception as e:
        print(f"Backup warning: {e}")

    # --- AUTHENTICATION & PASSWORD WORKFLOW ---
    @app.before_request
    def require_login():
        allowed_routes = ['login', 'forgot_password', 'static']
        if request.endpoint in allowed_routes or (request.endpoint and 'static' in request.endpoint):
            return
        if not session.get('authenticated'):
            return redirect(url_for('login'))

    @app.route('/login', methods=['GET', 'POST'])
    def login():
        error = None
        config = load_config()
        if request.method == 'POST':
            if request.form.get('password', '') == config.get('password'):
                session['authenticated'] = True
                return redirect(url_for('index'))
            else:
                error = 'Galat Password! Kripya sahi password enter karein.'
        return render_template('login.html', error=error)

    @app.route('/forgot-password', methods=['GET', 'POST'])
    def forgot_password():
        config = load_config()
        message = None
        success = False
        if request.method == 'POST':
            answer = request.form.get('answer', '').strip().lower()
            new_pass = request.form.get('new_password', '').strip()
            
            if answer == config.get('security_answer', '').lower():
                if new_pass:
                    config['password'] = new_pass
                    save_config(config)
                    success = True
                    message = "Password successfully reset ho gaya hai! Ab aap naye password se login kar sakte hain."
                else:
                    message = "Kripya naya password enter karein."
            else:
                message = "Security Answer galat hai! Dobara koshish karein."

        return render_template('forgot_password.html', message=message, success=success, config_q=config.get('security_question'))


    @app.route('/logout')
    def logout():
        session.clear()
        flash("Login successful. Welcome back!", "success")
        return redirect(url_for('login'))

    # CSRF Token compatibility for templates
    @app.context_processor
    def inject_csrf():
        return dict(csrf_token=lambda: "dummy-csrf-token")

    def register_module(module_name, preferred_bp_name, url_prefix=None):
        try:
            mod = __import__(f'app.routes.{module_name}', fromlist=[preferred_bp_name])
            bp = None
            if hasattr(mod, preferred_bp_name):
                bp = getattr(mod, preferred_bp_name)
            else:
                for attr_name in dir(mod):
                    attr = getattr(mod, attr_name)
                    if isinstance(attr, Blueprint):
                        bp = attr
                        break
            
            if bp:
                if url_prefix:
                    app.register_blueprint(bp, url_prefix=url_prefix)
                else:
                    app.register_blueprint(bp)
        except Exception as e:
            print(f"Error loading module {module_name}: {e}")

    # Register all application blueprints safely
    register_module('issue_challan', 'issue_challan_bp')
    register_module('dashboard', 'dashboard_bp')
    register_module('parties', 'parties_bp')
    register_module('items', 'items_bp')
    register_module('opening_stock', 'opening_stock_bp')
    register_module('stock', 'stock_bp')
    register_module('sites', 'sites_bp')
    register_module('receive_challan', 'receive_challan_bp')
    register_module('stock_ledger', 'stock_ledger_bp')
    register_module('reports', 'reports_bp')
    register_module('settings', 'settings_bp')
    register_module('monthly_billing', 'monthly_billing_bp')
    register_module('audit_logs', 'audit_logs_bp')

    # Root route safe redirect
    @app.route('/')
    def index():
        try:
            return redirect(url_for('dashboard.dashboard'))
        except Exception:
            return redirect(url_for('issue_challan.index'))

    return app