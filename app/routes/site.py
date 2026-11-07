from flask import Blueprint, render_template
site_bp = Blueprint("site", __name__, url_prefix="/site")
@site_bp.route("/")
def index():
    return "Site Module"
