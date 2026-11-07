from __future__ import annotations

import json
import os
import re
import uuid
from datetime import datetime
from pathlib import Path
from threading import RLock
from typing import Any

from openpyxl import load_workbook

from app.config import Config
from app.services.excel_service import initialize_workbook


_LOCK = RLock()


class PartyServiceError(Exception):
    """A validation or data-storage error in Party Master."""


def _now() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _text(value: Any) -> str:
    return str(value or "").strip()


def _get_path() -> Path:
    return Path(Config.WORKBOOK_PATH)



def _validate_party(data: dict[str, Any]) -> dict[str, Any]:
    party_name = _text(data.get("Party_Name"))
    party_type = _text(data.get("Party_Type")).upper()
    gstin = _text(data.get("GSTIN")).upper()

    if not party_name:
        raise PartyServiceError("Party name is required.")

    if len(party_name) > 150:
        raise PartyServiceError(
            "Party name must not exceed 150 characters."
        )

    if party_type not in {"GST", "WITHOUT_GST"}:
        raise PartyServiceError(
            "Select GST Party or Without GST Party."
        )

    if party_type == "GST":
        if not gstin:
            raise PartyServiceError(
                "GSTIN is required for a GST Party."
            )

        gstin_pattern = (
            r"[0-9]{2}[A-Z]{5}[0-9]{4}"
            r"[A-Z][1-9A-Z]Z[0-9A-Z]"
        )

        if not re.fullmatch(gstin_pattern, gstin):
            raise PartyServiceError(
                "GSTIN must have a valid 15-character format."
            )
    else:
        # GSTIN is not applicable to a Without GST Party.
        gstin = ""

    phone = _text(data.get("Phone"))
    if phone and len(phone) > 25:
        raise PartyServiceError("Phone number is too long.")

    state = _text(data.get("State"))
    state_code = _text(data.get("State_Code"))

    if state_code and not re.fullmatch(r"[0-9]{2}", state_code):
        raise PartyServiceError(
            "State Code must contain exactly 2 digits."
        )

    address = _text(data.get("Address"))

    return {
        "Party_Name": party_name,
        "Party_Type": party_type,
        "GSTIN": gstin,
        "Phone": phone,
        "Address": address,
        "State": state,
        "State_Code": state_code,
    }



def _read_parties(wb) -> list[dict[str, Any]]:
    ws = wb["Customers"]
    headers = [
        ws.cell(row=1, column=col).value
        for col in range(1, ws.max_column + 1)
    ]

    rows = []
    for row_number in range(2, ws.max_row + 1):
        values = [
            ws.cell(row=row_number, column=col).value
            for col in range(1, len(headers) + 1)
        ]

        if not any(value is not None for value in values):
            continue

        record = dict(zip(headers, values))
        record["_row_number"] = row_number
        rows.append(record)

    return rows


def _append_audit(
    wb,
    action: str,
    party_id: str,
    reason: str,
    before: dict[str, Any] | None,
    after: dict[str, Any] | None,
    username: str,
) -> None:
    ws = wb["Audit_Log"]

    audit_id = "AUD-" + uuid.uuid4().hex[:12].upper()

    def serializable(data):
        if data is None:
            return ""
        return json.dumps(
            {
                key: value
                for key, value in data.items()
                if not key.startswith("_")
            },
            ensure_ascii=False,
            default=str,
        )

    ws.append([
        audit_id,
        _now(),
        username,
        action,
        "PARTY",
        party_id,
        reason,
        serializable(before),
        serializable(after),
    ])


def _save_atomic(wb, path: Path) -> None:
    temporary_path = path.with_name(
        f"{path.stem}.tmp{path.suffix}"
    )

    try:
        wb.save(temporary_path)
        os.replace(temporary_path, path)
    except Exception as exc:
        if temporary_path.exists():
            try:
                temporary_path.unlink()
            except OSError:
                pass
        raise PartyServiceError(
            "Could not save workbook. Close it in Excel if open, "
            "then retry."
        ) from exc


def list_parties(
    party_type: str | None = None,
    search: str = "",
    active_only: bool = True,
) -> list[dict[str, Any]]:
    """Return party records, optionally filtered by type and search."""
    path = initialize_workbook()

    with _LOCK:
        wb = load_workbook(path, read_only=True, data_only=True)
        try:
            parties = _read_parties(wb)
        finally:
            wb.close()

    requested_type = _text(party_type).upper()
    query = _text(search).casefold()
    result = []

    for party in parties:
        is_active = str(party.get("Is_Active", "")).strip().lower() in {
            "true", "1", "yes", "active"
        }

        if active_only and not is_active:
            continue

        if requested_type and party.get("Party_Type") != requested_type:
            continue

        if query:
            searchable = " ".join(
                _text(party.get(field))
                for field in (
                    "Party_ID",
                    "Party_Name",
                    "GSTIN",
                    "Phone",
                    "Contact_Person",
                )
            ).casefold()

            if query not in searchable:
                continue

        party.pop("_row_number", None)
        result.append(party)

    return sorted(
        result,
        key=lambda item: _text(item.get("Party_Name")).casefold(),
    )


def get_party(party_id: str) -> dict[str, Any] | None:
    """Return one party by its ID."""
    target = _text(party_id)

    for party in list_parties(active_only=False):
        if party.get("Party_ID") == target:
            return party

    return None


def create_party(
    data: dict[str, Any],
    username: str = "Administrator",
) -> str:
    """Validate and create a party; return its unique ID."""
    clean = _validate_party(data)
    path = initialize_workbook()

    with _LOCK:
        wb = load_workbook(path)
        try:
            existing = _read_parties(wb)

            for party in existing:
                if party.get("Is_Active") is False:
                    continue

                if str(party.get("Is_Active", "")).lower() in {
                    "false", "0", "no", "inactive"
                }:
                    continue

                if (
                    clean["Party_Type"] == "GST"
                    and _text(party.get("GSTIN")).upper()
                    == clean["GSTIN"]
                ):
                    raise PartyServiceError(
                        "This GSTIN already exists in Party Master."
                    )

            # Generate ID from the highest existing numeric suffix.
            max_number = 0
            for party in existing:
                match = re.fullmatch(
                    r"P(\d+)",
                    _text(party.get("Party_ID")),
                )
                if match:
                    max_number = max(
                        max_number,
                        int(match.group(1)),
                    )

            party_id = f"P{max_number + 1:06d}"
            created_at = _now()

            ws = wb["Customers"]
            headers = [
                ws.cell(row=1, column=col).value
                for col in range(1, ws.max_column + 1)
            ]

            record = {
                **clean,
                "Party_ID": party_id,
                "Is_Active": True,
                "Created_At": created_at,
                "Updated_At": created_at,
            }

            ws.append([record.get(header) for header in headers])

            _append_audit(
                wb,
                "CREATE",
                party_id,
                "Party created",
                None,
                record,
                username,
            )

            _save_atomic(wb, path)
            return party_id
        finally:
            wb.close()


def update_party(
    party_id: str,
    data: dict[str, Any],
    reason: str,
    username: str = "Administrator",
) -> None:
    """Update a party while preserving its ID and creation timestamp."""
    target = _text(party_id)
    reason = _text(reason)

    if not reason:
        raise PartyServiceError(
            "A reason is required when editing a party."
        )

    clean = _validate_party(data)
    path = initialize_workbook()

    with _LOCK:
        wb = load_workbook(path)
        try:
            ws = wb["Customers"]
            headers = [
                ws.cell(row=1, column=col).value
                for col in range(1, ws.max_column + 1)
            ]
            id_column = headers.index("Party_ID") + 1
            row_number = None

            for row in range(2, ws.max_row + 1):
                if ws.cell(row=row, column=id_column).value == target:
                    row_number = row
                    break

            if row_number is None:
                raise PartyServiceError("Party was not found.")

            before = dict(zip(
                headers,
                [
                    ws.cell(row=row_number, column=col).value
                    for col in range(1, len(headers) + 1)
                ],
            ))

            if str(before.get("Is_Active", "")).lower() in {
                "false", "0", "no", "inactive"
            }:
                raise PartyServiceError(
                    "This party is inactive. Reactivation workflow "
                    "must be used before editing it."
                )

            for other in _read_parties(wb):
                if other.get("Party_ID") == target:
                    continue

                if (
                    clean["Party_Type"] == "GST"
                    and _text(other.get("GSTIN")).upper()
                    == clean["GSTIN"]
                    and str(other.get("Is_Active", "")).lower()
                    not in {"false", "0", "no", "inactive"}
                ):
                    raise PartyServiceError(
                        "This GSTIN already belongs to another party."
                    )

            updated = {
                **before,
                **clean,
                "Party_ID": target,
                "Created_At": before.get("Created_At"),
                "Updated_At": _now(),
            }

            for column, header in enumerate(headers, start=1):
                if header in updated:
                    ws.cell(
                        row=row_number,
                        column=column,
                    ).value = updated[header]

            _append_audit(
                wb,
                "UPDATE",
                target,
                reason,
                before,
                updated,
                username,
            )

            _save_atomic(wb, path)
        finally:
            wb.close()


def deactivate_party(
    party_id: str,
    reason: str,
    username: str = "Administrator",
) -> None:
    """Deactivate a party without deleting its historical record."""
    target = _text(party_id)
    reason = _text(reason)

    if not reason:
        raise PartyServiceError(
            "A reason is required to deactivate a party."
        )

    path = initialize_workbook()

    with _LOCK:
        wb = load_workbook(path)
        try:
            ws = wb["Customers"]
            headers = [
                ws.cell(row=1, column=col).value
                for col in range(1, ws.max_column + 1)
            ]

            id_column = headers.index("Party_ID") + 1
            active_column = headers.index("Is_Active") + 1
            updated_column = headers.index("Updated_At") + 1

            row_number = None

            for row in range(2, ws.max_row + 1):
                if ws.cell(row=row, column=id_column).value == target:
                    row_number = row
                    break

            if row_number is None:
                raise PartyServiceError("Party was not found.")

            before = dict(zip(
                headers,
                [
                    ws.cell(row=row_number, column=col).value
                    for col in range(1, len(headers) + 1)
                ],
            ))

            if str(before.get("Is_Active", "")).lower() in {
                "false", "0", "no", "inactive"
            }:
                raise PartyServiceError("Party is already inactive.")

            ws.cell(row=row_number, column=active_column).value = False
            ws.cell(row=row_number, column=updated_column).value = _now()

            after = dict(before)
            after["Is_Active"] = False
            after["Updated_At"] = _now()

            _append_audit(
                wb,
                "DEACTIVATE",
                target,
                reason,
                before,
                after,
                username,
            )

            _save_atomic(wb, path)
        finally:
            wb.close()
