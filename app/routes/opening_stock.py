from flask import Blueprint, flash, redirect, render_template, request, url_for

from app.services.opening_stock_service import (
    OpeningStockServiceError,
    create_estimate,
    list_opening_stock,
    quantity_difference,
    verify_opening_stock,
)
from app.services.item_service import list_items
from app.services.party_service import get_party, list_parties
from app.services.site_service import list_sites


opening_stock_bp = Blueprint(
    "opening_stock",
    __name__,
    url_prefix="/opening-stock",
)


def _text(value):
    return str(value or "").strip()


def _load_form_options():
    items = list_items(active_only=True)
    parties = list_parties(active_only=True)

    sites_by_party = {}
    for party in parties:
        party_id = _text(party.get("Party_ID"))
        if party_id:
            sites_by_party[party_id] = list_sites(
                party_id,
                active_only=True,
            )

    return items, parties, sites_by_party


@opening_stock_bp.route("/", methods=["GET"])
def index():
    location_type = _text(request.args.get("location_type")).upper()
    party_id = _text(request.args.get("party_id"))
    site_id = _text(request.args.get("site_id"))

    if location_type not in ("", "GODOWN", "SITE"):
        location_type = ""

    try:
        items, parties, sites_by_party = _load_form_options()
        records = list_opening_stock(
            location_type=location_type,
            party_id=party_id,
            site_id=site_id,
        )

        for record in records:
            record["Quantity_Difference"] = quantity_difference(record)

        return render_template(
            "opening_stock/index.html",
            records=records,
            items=items,
            parties=parties,
            sites_by_party=sites_by_party,
            selected_location_type=location_type,
            selected_party_id=party_id,
            selected_site_id=site_id,
        )
    except Exception as exc:
        flash(f"Opening Stock load nahi ho saka: {exc}", "error")
        return render_template(
            "opening_stock/index.html",
            records=[],
            items=[],
            parties=[],
            sites_by_party={},
            selected_location_type=location_type,
            selected_party_id=party_id,
            selected_site_id=site_id,
        )


@opening_stock_bp.route("/estimate", methods=["POST"])
def estimate():
    data = {
        "Location_Type": _text(request.form.get("Location_Type")),
        "Party_ID": _text(request.form.get("Party_ID")),
        "Site_ID": _text(request.form.get("Site_ID")),
        "Item_ID": _text(request.form.get("Item_ID")),
        "Estimated_Qty": _text(request.form.get("Estimated_Qty")),
        "Estimated_Date": _text(request.form.get("Estimated_Date")),
        "Estimated_Notes": _text(request.form.get("Estimated_Notes")),
    }

    try:
        create_estimate(data)
        flash(
            "Estimated opening stock save ho gaya. "
            "Yeh abhi verified stock nahi hai.",
            "success",
        )
    except OpeningStockServiceError as exc:
        flash(str(exc), "error")
    except Exception as exc:
        flash(f"Opening Stock save nahi hua: {exc}", "error")

    return redirect(url_for("opening_stock.index"))


@opening_stock_bp.route("/<stock_id>/verify", methods=["POST"])
def verify(stock_id):
    data = {
        "Verified_Qty": _text(request.form.get("Verified_Qty")),
        "Verified_Date": _text(request.form.get("Verified_Date")),
        "Verified_Notes": _text(request.form.get("Verified_Notes")),
    }

    try:
        verify_opening_stock(
            stock_id,
            data["Verified_Qty"],
            data["Verified_Date"],
            data["Verified_Notes"],
        )
        flash(
            "Physical verification save ho gayi. "
            "Stock record VERIFIED mark hua.",
            "success",
        )
    except OpeningStockServiceError as exc:
        flash(str(exc), "error")
    except Exception as exc:
        flash(f"Verification save nahi hui: {exc}", "error")

    return redirect(url_for("opening_stock.index"))


@opening_stock_bp.route("/party-sites/<party_id>", methods=["GET"])
def party_sites(party_id):
    """Party selection ke baad active sites ke liye JSON response."""
    from flask import jsonify

    party = get_party(party_id)
    if not party:
        return jsonify({"sites": []}), 404

    sites = list_sites(party_id, active_only=True)
    return jsonify({
        "sites": [
            {
                "Site_ID": site.get("Site_ID"),
                "Site_Name": site.get("Site_Name"),
            }
            for site in sites
        ]
    })
