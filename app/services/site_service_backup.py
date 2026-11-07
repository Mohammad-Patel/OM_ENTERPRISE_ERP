
from __future__ import annotations

import json
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


def _now() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _normalise(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _active(value: Any) -> bool:
    return _normalise(value).upper() not in {
        "",
        "0",
        "FALSE",
        "NO",
        "N",
        "INACTIVE",
    }


def _site_id() -> str:
    return f"SITE-{uuid4().hex[:10].upper()}"


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
    party_id = _normalise(party_id)

    try:
        sites = get_sheet_rows("SiteLocations")
    except ExcelServiceError as exc:
        raise SiteServiceError(str(exc)) from exc

    result = [
        site for site in sites
        if _normalise(site.get("Party_ID")) == party_id
        and (
            not active_only
            or _active(site.get("Is_Active"))
        )
    ]

    return sorted(
        result,
        key=lambda site: (
            not _active(site.get("Is_Active")),
            _normalise(site.get("Site_Name")).casefold(),
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
    # Site Name hi required hai. Baaki fields optional hain.
    # Phone par jaanbujhkar koi number/length validation nahi hai.
    cleaned = {
        field: _normalise(data.get(field))
        for field in SITE_FIELDS
    }

    if not cleaned["Site_Name"]:
        raise SiteServiceError("Site Name required hai.")

    return cleaned


def _validate_unique_name(
    party_id: str,
    site_name: str,
    exclude_site_id: str = "",
) -> None:
    for site in list_sites(party_id, active_only=True):
        if (
            _normalise(site.get("Site_ID")) != _normalise(exclude_site_id)
            and _normalise(site.get("Site_Name")).casefold()
            == site_name.casefold()
        ):
            raise SiteServiceError(
                "Is Party ke liye same naam ka Active Site pehle se hai."
            )


def _find_columns(ws) -> dict[str, int]:
    return {
        _normalise(ws.cell(1, col).value): col
        for col in range(1, ws.max_column + 1)
        if _normalise(ws.cell(1, col).value)
    }


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

    if not _active(party.get("Is_Active")):
        raise SiteServiceError(
            "Inactive Party ke liye naya Site add nahi kar sakte."
        )

    cleaned = _clean_site_data(data)
    _validate_unique_name(party_id, cleaned["Site_Name"])

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

            required = set(record)
            missing = required - set(columns)
            if missing:
                raise SiteServiceError(
                    "SiteLocations mein ye columns missing hain: "
                    + ", ".join(sorted(missing))
                    + ". Excel ke header row mein columns add karein."
                )

            ws.append([
                record.get(_normalise(ws.cell(1, col).value))
                if _normalise(ws.cell(1, col).value)
                else None
                for col in range(1, ws.max_column + 1)
            ])

            _audit(
                audit_ws, username, "CREATE", site_id, after=record
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

    if not _active(old_site.get("Is_Active")):
        raise SiteServiceError("Inactive Site ko edit nahi kar sakte.")

    cleaned = _clean_site_data(data)
    _validate_unique_name(party_id, cleaned["Site_Name"], site_id)

    updated = dict(old_site)
    updated.update(cleaned)
    updated["Updated_At"] = _now()

    path = initialize_workbook(WORKBOOK_PATH)

    with _WORKBOOK_LOCK:
        wb = None
        try:
            wb = load_workbook(path)
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

            for field, value in updated.items():
                if field in columns:
                    ws.cell(found_row, columns[field]).value = value

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
    old_site = get_site(party_id, site_id)
    if not old_site:
        raise SiteServiceError("Site record nahi mila.")

    if not _active(old_site.get("Is_Active")):
        raise SiteServiceError("Yeh Site pehle se inactive hai.")

    updated = dict(old_site)
    updated["Is_Active"] = "NO"
    updated["Updated_At"] = _now()

    path = initialize_workbook(WORKBOOK_PATH)

    with _WORKBOOK_LOCK:
        wb = None
        try:
            wb = load_workbook(path)
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

            ws.cell(found_row, columns["Is_Active"]).value = "NO"
            ws.cell(found_row, columns["Updated_At"]).value = (
                updated["Updated_At"]
            )

            _audit(
                audit_ws,
                username,
                "DEACTIVATE",
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
                "Site deactivate nahi hua. Workbook ko Excel mein "
                "band karke dobara try karein."
            ) from exc
        finally:
            if wb is not None:
                wb.close()
