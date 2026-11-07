
from flask import (
    Blueprint,
    abort,
    flash,
    redirect,
    render_template,
    request,
    url_for,
)

from app.services.party_service import (
    PartyServiceError,
    create_party,
    delete_party,
    get_party,
    list_parties,
    update_party,
)

parties_bp = Blueprint(
    "parties",
    __name__,
    url_prefix="/parties",
)


def _form_data():
    return {
        "Party_Name": request.form.get("Party_Name", "").strip(),
        "Party_Type": request.form.get("Party_Type", "").strip().upper(),
        "GSTIN": request.form.get("GSTIN", "").strip().upper(),
        "Phone": request.form.get("Phone", "").strip(),
        "Address": request.form.get("Address", "").strip(),
        "State": request.form.get("State", "").strip(),
        "State_Code": request.form.get("State_Code", "").strip(),
    }


@parties_bp.route("/")
def index():
    party_type = request.args.get("type", "").strip().upper()
    search = request.args.get("q", "").strip()

    if party_type not in {"", "GST", "WITHOUT_GST"}:
        party_type = ""

    try:
        parties = list_parties(
            party_type=party_type or None,
            search=search,
            active_only=False,
        )
    except PartyServiceError as exc:
        flash(str(exc), "danger")
        parties = []

    return render_template(
        "parties/index.html",
        parties=parties,
        selected_type=party_type,
        search=search,
    )


@parties_bp.route("/new", methods=["GET", "POST"])
def create():
    if request.method == "POST":
        try:
            party_id = create_party(
                _form_data(),
                username="Administrator",
            )
            flash(
                f"Party created successfully. Party ID: {party_id}",
                "success",
            )
            return redirect(url_for("parties.index"))

        except PartyServiceError as exc:
            flash(str(exc), "danger")
            return render_template(
                "parties/form.html",
                party=request.form,
                page_title="Add Party",
                form_action=url_for("parties.create"),
                is_edit=False,
            )

    return render_template(
        "parties/form.html",
        party={},
        page_title="Add Party",
        form_action=url_for("parties.create"),
        is_edit=False,
    )


@parties_bp.route("/<party_id>/edit", methods=["GET", "POST"])
def edit(party_id):
    party = get_party(party_id)

    if party is None:
        flash("Party record was not found.", "warning")
        return redirect(url_for("parties.index"))

    if request.method == "POST":
        try:
            update_party(
                party_id,
                _form_data(),
                reason=request.form.get("Reason", "").strip(),
                username="Administrator",
            )
            flash("Party updated successfully.", "success")
            return redirect(url_for("parties.index"))

        except PartyServiceError as exc:
            flash(str(exc), "danger")
            form_data = _form_data()
            form_data["Party_ID"] = party_id
            form_data["Created_At"] = party.get("Created_At")

            return render_template(
                "parties/form.html",
                party=form_data,
                page_title="Edit Party",
                form_action=url_for(
                    "parties.edit",
                    party_id=party_id,
                ),
                is_edit=True,
            )

    return render_template(
        "parties/form.html",
        party=party,
        page_title="Edit Party",
        form_action=url_for(
            "parties.edit",
            party_id=party_id,
        ),
        is_edit=True,
    )



@parties_bp.route("/<party_id>/view")
def view(party_id):
    party = get_party(party_id)

    if party is None:
        abort(404)

    from app.services.site_service import (
        SiteServiceError,
        list_sites,
    )

    try:
        sites = list_sites(party_id)
    except SiteServiceError as exc:
        flash(str(exc), "danger")
        sites = []

    return render_template(
        "parties/view.html",
        party=party,
        sites=sites,
    )



@parties_bp.route("/<party_id>/delete", methods=["POST"])
def delete(party_id):
    try:
        delete_party(
            party_id,
            username="Administrator",
        )
        flash(
            f"Party {party_id} permanently deleted.",
            "success",
        )

    except PartyServiceError as exc:
        flash(str(exc), "danger")

    return redirect(url_for("parties.index"))
