from flask import Blueprint, render_template
party_bp = Blueprint("party", __name__, url_prefix="/party")
@party_bp.route("/")
def index():
    return "Party Module"
