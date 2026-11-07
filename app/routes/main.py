from flask import Blueprint, render_template, redirect, url_for

main_bp = Blueprint("main", __name__)

@main_bp.route("/")
def index():
    return redirect(url_for("issue_challan.index"))
