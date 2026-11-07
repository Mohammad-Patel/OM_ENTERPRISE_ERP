
from flask import Blueprint, flash, redirect, request, url_for

from app.services.party_service import get_party
from app.services.site_service import (
    SiteServiceError,
    create_site,
    update_site,
    deactivate_site,
    delete_site,
)


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


@sites_bp.route("/", methods=["GET"])
def index(party_id):
    return _party_page(party_id)


@sites_bp.route("/new", methods=["GET", "POST"])
def create(party_id):
    if request.method == "GET":
        return _party_page(party_id)

    if not _party_exists(party_id):
        flash("Party record nahi mila.", "warning")
        return redirect(url_for("parties.index"))

    try:
        site_id = create_site(
            party_id,
            _site_form_data(),
            username="Administrator",
        )

        flash(
            f"Site successfully added. Site ID: {site_id}",
            "success",
        )

    except SiteServiceError as exc:
        flash(str(exc), "danger")

    return _party_page(party_id)


@sites_bp.route("/<site_id>/edit", methods=["GET", "POST"])
def edit(party_id, site_id):
    if request.method == "GET":
        return _party_page(party_id)

    if not _party_exists(party_id):
        flash("Party record nahi mila.", "warning")
        return redirect(url_for("parties.index"))

    try:
        update_site(
            party_id,
            site_id,
            _site_form_data(),
            reason="Updated from Party Details modal",
            username="Administrator",
        )

        flash("Site successfully updated.", "success")

    except SiteServiceError as exc:
        flash(str(exc), "danger")

    return _party_page(party_id)


@sites_bp.route("/<site_id>/deactivate", methods=["POST"])
def deactivate(party_id, site_id):
    if not _party_exists(party_id):
        flash("Party record nahi mila.", "warning")
        return redirect(url_for("parties.index"))

    try:
        deactivate_site(
            party_id,
            site_id,
            reason=(
                request.form.get("Reason", "").strip()
                or "Deactivated from Party Details"
            ),
            username="Administrator",
        )

        flash(
            "Site deactivated. Historical records preserved.",
            "success",
        )

    except SiteServiceError as exc:
        flash(str(exc), "danger")

    return _party_page(party_id)



@sites_bp.route("/<site_id>/delete", methods=["POST"])
def delete(party_id, site_id):
    """Legacy Delete route: site ko permanently delete nahi, deactivate kare."""
    if not _party_exists(party_id):
        flash("Party record nahi mila.", "warning")
        return redirect(url_for("parties.index"))

    try:
        delete_site(
            party_id,
            site_id,
            reason=(
                request.form.get("Reason", "").strip()
                or "Site deactivated from Party Details"
            ),
            username="Administrator",
        )
        flash(
            "Site successfully deactivated. Site ID aur historical record preserve hain.",
            "success",
        )
    except SiteServiceError as exc:
        flash(str(exc), "danger")

    return _party_page(party_id)
