from flask import Blueprint, render_template, request, redirect, url_for, flash
from app.services.billing_service import generate_monthly_bill, BillingServiceError
from app.services.excel_service import get_sheet_rows
from app.utils.number_to_words import amount_in_words  # Utility helper if present

monthly_billing_bp = Blueprint("monthly_billing", __name__, url_prefix="/billing")


@monthly_billing_bp.route("/generate", methods=["GET", "POST"])
def generate():
    """Form to select Month/Dates, Party, and Site to generate Bill."""
    if request.method == "POST":
        party_id = request.form.get("party_id")
        site_id = request.form.get("site_id")
        start_date = request.form.get("start_date")
        end_date = request.form.get("end_date")

        try:
            bill_data = generate_monthly_bill(party_id, site_id, start_date, end_date)
            
            # Fetch Party and Site info for preview
            parties = get_sheet_rows("Parties")
            sites = get_sheet_rows("Sites")
            party = next((p for p in parties if str(p.get("Party_ID")) == str(party_id)), {})
            site = next((s for s in sites if str(s.get("Site_ID")) == str(site_id)), {})

            # Prepare Balance Quantity Summary string for footer box
            balance_strings = []
            for b_item in bill_data["items"]:
                last_seg = b_item["segments"][-1]
                balance_strings.append(f"{b_item['item_name']} = {int(last_seg['qty'])}")
            
            balance_summary_text = ", ".join(balance_strings)

            # Amount in words
            net_total = bill_data["summary"]["net_total"]
            try:
                words_text = amount_in_words(net_total)
            except Exception:
                words_text = f"RUPEES {net_total} ONLY"

            return render_template(
                "print/monthly_invoice.html",
                bill=bill_data,
                party=party,
                site=site,
                balance_summary=balance_summary_text,
                amount_in_words=words_text
            )

        except BillingServiceError as e:
            flash(str(e), "danger")
        except Exception as e:
            flash(f"Error generating bill: {str(e)}", "danger")

    parties = get_sheet_rows("Parties")
    sites = get_sheet_rows("Sites")
    return render_template("billing/form.html", parties=parties, sites=sites)