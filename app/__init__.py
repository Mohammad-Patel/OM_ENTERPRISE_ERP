from flask import Flask, redirect, url_for, Blueprint

def create_app():
    app = Flask(__name__)
    app.config['SECRET_KEY'] = 'om_enterprise_secret_key_2026'

    # CSRF Token compatibility for templates
    @app.context_processor
    def inject_csrf():
        return dict(csrf_token=lambda: "dummy-csrf-token")

    # Register all application blueprints safely
    modules = [
        'issue_challan', 'dashboard', 'parties', 'items', 
        'opening_stock', 'stock', 'sites', 'receive_challan', 
        'stock_ledger', 'reports', 'settings', 'monthly_billing'
    ]
    
    for mod_name in modules:
        try:
            mod = __import__(f'app.routes.{mod_name}', fromlist=[f'{mod_name}_bp'])
            app.register_blueprint(getattr(mod, f'{mod_name}_bp'))
        except (ImportError, AttributeError):
            # Fallback placeholder blueprint so url_for never crashes
            bp = Blueprint(mod_name, __name__)
            @bp.route('/')
            @bp.route('/generate')
            def index(name=mod_name):
                return redirect(url_for('issue_challan.index'))
            app.register_blueprint(bp, url_prefix=f'/{mod_name.replace("_", "-")}')

    # Root route redirect to dashboard or issue challan
    @app.route('/')
    def index():
        return redirect('/dashboard')

    return app
