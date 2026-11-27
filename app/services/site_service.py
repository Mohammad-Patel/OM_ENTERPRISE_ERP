from __future__ import annotations

import json
import re
from datetime import datetime
from typing import Any
from uuid import uuid4

from openpyxl import load_workbook

from app.services.excel_service import (
    ExcelServiceError,
    WORKBOOK_PATH,
    _WORKBOOK_LOCK,
    _save_atomic,
    get_sheet_rows,
    initialize_workbook,
)
from app.services.party_service import get_party


class SiteServiceError(Exception):
    """Raised when a site operation cannot be completed safely."""


SITE_FIELDS = (
    "Site_Name",
    "Site_Address",
    "State",
    "State_Code",
    "Phone",
)

INACTIVE_VALUES = {"NO", "N", "0", "FALSE", "INACTIVE"}

# Sheets whose rows point to a site. A site that is used in any of them is never deleted
# (its history must stay readable) - it can only be made inactive.
SITE_REFERENCE_SHEETS = {
    "Issue_Challan": "Issue Challan",
    "Receive_Challan": "Receive Challan",
    "OpeningStock": "Opening Stock",
    "StockLedger": "Stock Ledger",
    "Monthly_Bills": "Monthly Bills",
    "Damage_Charges": "Damage Charges",
}


def _now() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _normalise(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _same_id(a: Any, b: Any) -> bool:
    """1, '1' and 1.0 are the same ID."""
    def clean(v: Any) -> str:
        return re.sub(r"\.0+$", "", _normalise(v))
    return clean(a) == clean(b)


def _norm_text(value: Any) -> str:
    """Case / spacing / trailing punctuation ignore karke compare ke liye text."""
    text = re.sub(r"\s+", " ", _normalise(value))
    return text.strip(" .,;-").casefold()


def _site_label(site: dict[str, Any]) -> str:
    """Site ka display text: Site Name optional hai, isliye Site Address fallback."""
    return _normalise(site.get("Site_Name")) or _normalise(site.get("Site_Address"))


def _site_address_text(site: dict[str, Any]) -> str:
    """Purani sites mein pata Site Name mein bhara hota tha - use bhi address maanein."""
    return _normalise(site.get("Site_Address")) or _normalise(site.get("Site_Name"))


def _is_active(value: Any) -> bool:
    """Blank legacy status ko active maana jayega."""
    return _normalise(value).upper() not in INACTIVE_VALUES


def _site_id() -> str:
    return f"SITE-{uuid4().hex[:10].upper()}"


def _find_columns(ws) -> dict[str, int]:
    """Excel header names aur unke column numbers return kare."""
    return {
        _normalise(ws.cell(1, col).value): col
        for col in range(1, ws.max_column + 1)
        if _normalise(ws.cell(1, col).value)
    }


def _audit(
    ws,
    username: str,
    action: str,
    site_id: str,
    reason: str = "",
    before: dict | None = None,
    after: dict | None = None,
) -> None:
    headers = [
        _normalise(ws.cell(1, col).value)
        for col in range(1, ws.max_column + 1)
    ]

    record = {
        "Audit_ID": f"AUD-{uuid4().hex[:10].upper()}",
        "Event_Time": _now(),
        "Username": username or "Administrator",
        "Action": action,
        "Entity_Type": "SITE_LOCATION",
        "Entity_ID": site_id,
        "Reason": reason,
        "Before_Data": json.dumps(
            before, ensure_ascii=False, default=str
        ) if before else "",
        "After_Data": json.dumps(
            after, ensure_ascii=False, default=str
        ) if after else "",
    }

    ws.append([
        record.get(header, "") if header else None
        for header in headers
    ])


def list_sites(
    party_id: str,
    active_only: bool = False,
) -> list[dict[str, Any]]:
    """Party ki sites return kare; zaroorat par active sites filter kare."""
    party_id = _normalise(party_id)

    try:
        sites = get_sheet_rows("SiteLocations")
    except ExcelServiceError as exc:
        raise SiteServiceError(str(exc)) from exc

    result = [
        site
        for site in sites
        if _normalise(site.get("Party_ID")) == party_id
        and _normalise(site.get("Site_ID"))
        and (not active_only or _is_active(site.get("Is_Active")))
    ]

    return sorted(
        result,
        key=lambda site: (
            not _is_active(site.get("Is_Active")),
            _site_label(site).casefold(),
        ),
    )


def get_site(
    party_id: str,
    site_id: str,
) -> dict[str, Any] | None:
    for site in list_sites(party_id):
        if _normalise(site.get("Site_ID")) == _normalise(site_id):
            return site
    return None


def _clean_site_data(data: dict[str, Any]) -> dict[str, str]:
    cleaned = {
        field: _normalise(data.get(field))
        for field in SITE_FIELDS
    }

    # Site Name optional hai; Site Address required hai.
    if not cleaned["Site_Address"]:
        raise SiteServiceError("Site Address required hai.")

    return cleaned


def _validate_unique_site(
    party_id: str,
    cleaned: dict[str, str],
    exclude_site_id: str = "",
) -> None:
    """
    Duplicate site roke:
      * same Site Name (jab naam diya ho)
      * same Site Address, jab naam khali ho ya dono sites ka naam same ho
    Alag naam ke saath same address allowed hai (ek hi complex ke do projects).
    """
    new_name = _norm_text(cleaned.get("Site_Name"))
    new_addr = _norm_text(cleaned.get("Site_Address"))

    for site in list_sites(party_id):
        if _normalise(site.get("Site_ID")) == _normalise(exclude_site_id):
            continue

        other_name = _norm_text(site.get("Site_Name"))
        other_addr = _norm_text(_site_address_text(site))

        if new_name and new_name == other_name:
            raise SiteServiceError(
                "Is Party ke liye same naam ka Site pehle se hai."
            )

        if new_addr and new_addr == other_addr and (
            not new_name or new_name == other_name
        ):
            raise SiteServiceError(
                "Is Party ke liye ye Site Address pehle se maujood hai."
            )


def _validate_unique_name(
    party_id: str,
    site_name: str,
    exclude_site_id: str = "",
) -> None:
    """Purana naam - compatibility ke liye."""
    _validate_unique_site(
        party_id, {"Site_Name": site_name, "Site_Address": ""}, exclude_site_id
    )


def _find_site_row(
    ws,
    columns: dict[str, int],
    party_id: str,
    site_id: str,
) -> int | None:
    if "Site_ID" not in columns or "Party_ID" not in columns:
        raise SiteServiceError(
            "SiteLocations mein Site_ID ya Party_ID column missing hai."
        )

    for row_number in range(2, ws.max_row + 1):
        if (
            _normalise(ws.cell(row_number, columns["Site_ID"]).value)
            == _normalise(site_id)
            and _normalise(ws.cell(row_number, columns["Party_ID"]).value)
            == _normalise(party_id)
        ):
            return row_number

    return None


def create_site(
    party_id: str,
    data: dict[str, Any],
    username: str = "Administrator",
) -> str:
    party_id = _normalise(party_id)

    party = get_party(party_id)
    if not party:
        raise SiteServiceError("Selected Party record nahi mila.")

    # Blank legacy Party status ko active treat karein.
    party_active = _is_active(party.get("Is_Active"))

    if not party_active:
        raise SiteServiceError(
            "Inactive Party ke liye naya Site add nahi kar sakte."
        )

    cleaned = _clean_site_data(data)
    _validate_unique_site(party_id, cleaned)

    site_id = _site_id()
    now = _now()

    record = {
        "Site_ID": site_id,
        "Party_ID": party_id,
        **cleaned,
        "Is_Active": "YES",
        "Created_At": now,
        "Updated_At": now,
    }

    path = initialize_workbook(WORKBOOK_PATH)

    with _WORKBOOK_LOCK:
        wb = None
        try:
            wb = load_workbook(path)

            if "SiteLocations" not in wb.sheetnames:
                raise SiteServiceError("SiteLocations sheet nahi mili.")

            if "Audit_Log" not in wb.sheetnames:
                raise SiteServiceError("Audit_Log sheet nahi mili.")

            ws = wb["SiteLocations"]
            audit_ws = wb["Audit_Log"]
            columns = _find_columns(ws)

            missing = set(record) - set(columns)

            if missing:
                raise SiteServiceError(
                    "SiteLocations mein ye columns missing hain: "
                    + ", ".join(sorted(missing))
                    + ". Excel ke header row mein columns add karein."
                )

            ws.append([
                record.get(_normalise(ws.cell(1, col).value), "")
                for col in range(1, ws.max_column + 1)
            ])

            _audit(
                audit_ws,
                username,
                "CREATE",
                site_id,
                after=record,
            )

            _save_atomic(wb, path)
            return site_id

        except (SiteServiceError, ExcelServiceError):
            raise
        except Exception as exc:
            raise SiteServiceError(
                "Site create nahi hua. Workbook Excel mein khula ho "
                "to use band karke dobara try karein."
            ) from exc
        finally:
            if wb is not None:
                wb.close()


def update_site(
    party_id: str,
    site_id: str,
    data: dict[str, Any],
    reason: str = "",
    username: str = "Administrator",
) -> None:
    old_site = get_site(party_id, site_id)

    if not old_site:
        raise SiteServiceError("Site record nahi mila.")

    if not _is_active(old_site.get("Is_Active")):
        raise SiteServiceError(
            "Inactive Site ko update nahi kar sakte."
        )

    cleaned = _clean_site_data(data)
    _validate_unique_site(party_id, cleaned, site_id)

    updated = dict(old_site)
    updated.update(cleaned)
    updated["Updated_At"] = _now()

    path = initialize_workbook(WORKBOOK_PATH)

    with _WORKBOOK_LOCK:
        wb = None
        try:
            wb = load_workbook(path)

            if "SiteLocations" not in wb.sheetnames:
                raise SiteServiceError("SiteLocations sheet nahi mili.")

            if "Audit_Log" not in wb.sheetnames:
                raise SiteServiceError("Audit_Log sheet nahi mili.")

            ws = wb["SiteLocations"]
            audit_ws = wb["Audit_Log"]
            columns = _find_columns(ws)

            found_row = _find_site_row(
                ws, columns, party_id, site_id
            )

            if found_row is None:
                raise SiteServiceError(
                    "Site worksheet mein record nahi mila."
                )

            missing = set(SITE_FIELDS) - set(columns)

            if missing:
                raise SiteServiceError(
                    "SiteLocations mein ye columns missing hain: "
                    + ", ".join(sorted(missing))
                )

            # Workbook ke latest record ka status dobara check karein.
            if (
                "Is_Active" in columns
                and not _is_active(
                    ws.cell(found_row, columns["Is_Active"]).value
                )
            ):
                raise SiteServiceError(
                    "Inactive Site ko update nahi kar sakte."
                )

            for field, value in cleaned.items():
                ws.cell(found_row, columns[field]).value = value

            if "Updated_At" in columns:
                ws.cell(
                    found_row, columns["Updated_At"]
                ).value = updated["Updated_At"]

            _audit(
                audit_ws,
                username,
                "UPDATE",
                site_id,
                reason=reason,
                before=old_site,
                after=updated,
            )

            _save_atomic(wb, path)

        except (SiteServiceError, ExcelServiceError):
            raise
        except Exception as exc:
            raise SiteServiceError(
                "Site update nahi hua. Workbook Excel mein khula ho "
                "to use band karke dobara try karein."
            ) from exc
        finally:
            if wb is not None:
                wb.close()


def deactivate_site(
    party_id: str,
    site_id: str,
    reason: str = "",
    username: str = "Administrator",
) -> None:
    """Site ko inactive karein; Excel row aur Site ID preserve karein."""
    party_id = _normalise(party_id)
    site_id = _normalise(site_id)

    if not party_id or not site_id:
        raise SiteServiceError("Party ID aur Site ID required hain.")

    path = initialize_workbook(WORKBOOK_PATH)

    with _WORKBOOK_LOCK:
        wb = None
        try:
            wb = load_workbook(path)

            if "SiteLocations" not in wb.sheetnames:
                raise SiteServiceError("SiteLocations sheet nahi mili.")

            if "Audit_Log" not in wb.sheetnames:
                raise SiteServiceError("Audit_Log sheet nahi mili.")

            ws = wb["SiteLocations"]
            audit_ws = wb["Audit_Log"]
            columns = _find_columns(ws)

            if "Is_Active" not in columns:
                raise SiteServiceError(
                    "SiteLocations mein Is_Active column missing hai."
                )

            found_row = _find_site_row(
                ws, columns, party_id, site_id
            )

            if found_row is None:
                raise SiteServiceError("Site record nahi mila.")

            before = {
                header: ws.cell(found_row, col).value
                for header, col in columns.items()
            }

            current_status = ws.cell(
                found_row, columns["Is_Active"]
            ).value

            if not _is_active(current_status):
                raise SiteServiceError("Ye Site pehle se inactive hai.")

            updated_at = _now()

            ws.cell(
                found_row, columns["Is_Active"]
            ).value = "NO"

            if "Updated_At" in columns:
                ws.cell(
                    found_row, columns["Updated_At"]
                ).value = updated_at

            after = dict(before)
            after["Is_Active"] = "NO"

            if "Updated_At" in columns:
                after["Updated_At"] = updated_at

            _audit(
                audit_ws,
                username,
                "DEACTIVATE",
                site_id,
                reason=reason or "Site deactivated",
                before=before,
                after=after,
            )

            _save_atomic(wb, path)

        except (SiteServiceError, ExcelServiceError):
            raise
        except Exception as exc:
            raise SiteServiceError(
                "Site deactivate nahi hui. Workbook Excel mein khula ho "
                "to use band karke dobara try karein."
            ) from exc
        finally:
            if wb is not None:
                wb.close()


def _site_references(wb, party_id: str, site_id: str) -> dict[str, int]:
    """{'Issue Challan': 3, 'Monthly Bills': 1} - records that use this site."""
    found: dict[str, int] = {}
    for sheet_name, label in SITE_REFERENCE_SHEETS.items():
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        columns = _find_columns(ws)
        if "Site_ID" not in columns:
            continue
        count = 0
        for row_number in range(2, ws.max_row + 1):
            if not _same_id(ws.cell(row_number, columns["Site_ID"]).value, site_id):
                continue
            if "Party_ID" in columns:
                owner = ws.cell(row_number, columns["Party_ID"]).value
                if _normalise(owner) and not _same_id(owner, party_id):
                    continue
            count += 1
        if count:
            found[label] = count
    return found


def delete_site(
    party_id: str,
    site_id: str,
    reason: str = "",
    username: str = "Administrator",
) -> None:
    """
    Permanently delete an INACTIVE site (only its own row in SiteLocations).

    Safety rules
      * the site must be inactive
      * a site that is used by challans, bills, stock or damage records is NOT deleted
        (history stays readable) - keep it inactive instead
      * only the SiteLocations row is removed; every other sheet is saved untouched
      * the deletion is written to Audit_Log
    """
    party_id = _normalise(party_id)
    site_id = _normalise(site_id)

    if not party_id or not site_id:
        raise SiteServiceError("Party ID aur Site ID required hain.")

    path = initialize_workbook(WORKBOOK_PATH)

    with _WORKBOOK_LOCK:
        wb = None
        try:
            wb = load_workbook(path)

            if "SiteLocations" not in wb.sheetnames:
                raise SiteServiceError("SiteLocations sheet nahi mili.")

            if "Audit_Log" not in wb.sheetnames:
                raise SiteServiceError("Audit_Log sheet nahi mili.")

            ws = wb["SiteLocations"]
            audit_ws = wb["Audit_Log"]
            columns = _find_columns(ws)

            found_row = _find_site_row(ws, columns, party_id, site_id)

            if found_row is None:
                raise SiteServiceError("Site record nahi mila.")

            if (
                "Is_Active" in columns
                and _is_active(ws.cell(found_row, columns["Is_Active"]).value)
            ):
                raise SiteServiceError(
                    "Active Site delete nahi ho sakti. Pehle use deactivate karein."
                )

            used = _site_references(wb, party_id, site_id)
            if used:
                details = ", ".join(f"{label}: {count}" for label, count in used.items())
                raise SiteServiceError(
                    "Is Site ke records maujood hain (" + details + "), isliye ise delete "
                    "nahi kiya ja sakta. History ke liye Site ko inactive hi rehne dein."
                )

            before = {
                header: ws.cell(found_row, col).value
                for header, col in columns.items()
            }

            ws.delete_rows(found_row)

            _audit(
                audit_ws,
                username,
                "DELETE",
                site_id,
                reason=reason or "Permanently deleted inactive site",
                before=before,
            )

            _save_atomic(wb, path)

        except (SiteServiceError, ExcelServiceError):
            raise
        except Exception as exc:
            raise SiteServiceError(
                "Site delete nahi hui. Workbook Excel mein khula ho "
                "to use band karke dobara try karein."
            ) from exc
        finally:
            if wb is not None:
                wb.close()
