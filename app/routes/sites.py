import logging

from flask import Blueprint, flash, redirect, request, url_for

from app.services.party_service import get_party
from app.services.site_service import (
    SiteServiceError,
    create_site,
    update_site,
    deactivate_site,
    delete_site,
)

logger = logging.getLogger(__name__)

sites_bp = Blueprint(
    "sites",
    __name__,
    url_prefix="/parties/<party_id>/sites",
)


def _site_form_data():
    """Read and clean site form fields."""
    return {
        "Site_Name": request.form.get("Site_Name", "").strip(),
        "Site_Address": request.form.get("Site_Address", "").strip(),
        "State": request.form.get("State", "").strip(),
        "State_Code": request.form.get("State_Code", "").strip(),
        "Phone": request.form.get("Phone", "").strip(),
    }


def _party_page(party_id):
    """Return to the Party Details page."""
    return redirect(
        url_for("parties.view", party_id=party_id)
    )


def _party_exists(party_id):
    """Check whether the parent party exists."""
    return get_party(party_id) is not None


def _run(party_id, action, success_message, error_message):
    """
    Run a site action, flash the result and go back to the Party Details page.
    Every failure - expected (SiteServiceError) or not - shows a message instead of a 500 page.
    """
    if not _party_exists(party_id):
        flash("Party record nahi mila.", "warning")
        return redirect(url_for("parties.index"))

    try:
        result = action()
        flash(success_message(result) if callable(success_message) else success_message, "success")
    except SiteServiceError as exc:
        flash(str(exc), "danger")
    except Exception:
        logger.exception(error_message)
        flash(error_message, "danger")

    return _party_page(party_id)


@sites_bp.route("/", methods=["GET"])
def index(party_id):
    return _party_page(party_id)


@sites_bp.route("/new", methods=["GET", "POST"])
def create(party_id):
    if request.method == "GET":
        return _party_page(party_id)

    return _run(
        party_id,
        lambda: create_site(party_id, _site_form_data(), username="Administrator"),
        lambda site_id: f"Site successfully added. Site ID: {site_id}",
        "Site add nahi ho saki. Dobara try karein.",
    )


@sites_bp.route("/<site_id>/edit", methods=["GET", "POST"])
def edit(party_id, site_id):
    if request.method == "GET":
        return _party_page(party_id)

    return _run(
        party_id,
        lambda: update_site(
            party_id,
            site_id,
            _site_form_data(),
            reason="Updated from Party Details modal",
            username="Administrator",
        ),
        "Site successfully updated.",
        "Site update nahi ho saki. Dobara try karein.",
    )


@sites_bp.route("/<site_id>/deactivate", methods=["POST"])
def deactivate(party_id, site_id):
    return _run(
        party_id,
        lambda: deactivate_site(
            party_id,
            site_id,
            reason=(
                request.form.get("Reason", "").strip()
                or "Deactivated from Party Details"
            ),
            username="Administrator",
        ),
        "Site deactivated. Historical records preserved.",
        "Site deactivate nahi ho saki. Dobara try karein.",
    )


@sites_bp.route("/<site_id>/delete", methods=["POST"])
def delete(party_id, site_id):
    """Permanently delete an inactive site (refused if challans / bills / stock use it)."""
    return _run(
        party_id,
        lambda: delete_site(
            party_id,
            site_id,
            reason=(
                request.form.get("Reason", "").strip()
                or "Permanently deleted inactive site"
            ),
            username="Administrator",
        ),
        "Inactive site successfully deleted from the system.",
        "Site delete nahi ho saki. Dobara try karein.",
    )
