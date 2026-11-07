from flask import Blueprint, abort, flash, redirect, render_template, request, url_for

from app.services.item_service import (
    ItemServiceError,
    create_item,
    delete_item,
    get_item,
    list_items,
    update_item,
)

items_bp = Blueprint("items", __name__, url_prefix="/items")


def _form_data():
    return {
        "Item_Name": request.form.get("Item_Name", "").strip(),
        "Size": request.form.get("Size", "").strip(),
        "Category": request.form.get("Category", "").strip(),
        "Description": request.form.get("Description", "").strip(),
        "Unit": request.form.get("Unit", "Nos").strip() or "Nos",
        "Unit_Weight_KG": request.form.get("Unit_Weight_KG", "").strip(),
        "Daily_Rental_Rate": request.form.get("Daily_Rental_Rate", "").strip(),
        "Goods_Value_Rate": request.form.get("Goods_Value_Rate", "").strip(),
    }


@items_bp.route("/")
def index():
    search = request.args.get("q", "").strip()
    try:
        items = list_items(search=search)
    except ItemServiceError as exc:
        flash(str(exc), "danger")
        items = []

    return render_template("items/index.html", items=items, search=search)


@items_bp.route("/new", methods=["GET", "POST"])
def create():
    if request.method == "POST":
        try:
            item_id = create_item(_form_data())
            flash(f"Item created successfully. Item ID: {item_id}", "success")
            return redirect(url_for("items.index"))
        except ItemServiceError as exc:
            flash(str(exc), "danger")
            return render_template(
                "items/form.html",
                item=request.form,
                page_title="Add Item",
                form_action=url_for("items.create"),
                is_edit=False,
            )

    return render_template(
        "items/form.html",
        item={},
        page_title="Add Item",
        form_action=url_for("items.create"),
        is_edit=False,
    )


@items_bp.route("/<item_id>/view")
def view(item_id):
    item = get_item(item_id)
    if item is None:
        abort(404)
    return render_template("items/view.html", item=item)


@items_bp.route("/<item_id>/edit", methods=["GET", "POST"])
def edit(item_id):
    item = get_item(item_id)
    if item is None:
        flash("Item record was not found.", "warning")
        return redirect(url_for("items.index"))

    if request.method == "POST":
        try:
            update_item(
                item_id,
                _form_data(),
                reason=request.form.get("Reason", "").strip(),
            )
            flash("Item updated successfully.", "success")
            return redirect(url_for("items.index"))
        except ItemServiceError as exc:
            flash(str(exc), "danger")
            form_item = dict(request.form)
            form_item["Item_ID"] = item_id
            form_item["Item_Code"] = item.get("Item_Code")
            return render_template(
                "items/form.html",
                item=form_item,
                page_title="Edit Item",
                form_action=url_for("items.edit", item_id=item_id),
                is_edit=True,
            )

    return render_template(
        "items/form.html",
        item=item,
        page_title="Edit Item",
        form_action=url_for("items.edit", item_id=item_id),
        is_edit=True,
    )


@items_bp.route("/<item_id>/delete", methods=["POST"])
def delete(item_id):
    try:
        delete_item(
            item_id,
            reason=request.form.get("Reason", "").strip(),
        )
        flash("Item deleted successfully.", "success")
    except ItemServiceError as exc:
        flash(str(exc), "danger")
    return redirect(url_for("items.index"))
