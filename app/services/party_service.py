
from __future__ import annotations

import json
import os
import re
import shutil
import tempfile
import uuid
from datetime import datetime
from pathlib import Path
from threading import RLock
from typing import Any

from openpyxl import load_workbook

from app.config import Config
from app.services.excel_service import initialize_workbook


# Protects concurrent writes within this Python process.
# It does not coordinate separate processes or computers.
_LOCK = RLock()

PARTY_REQUIRED_HEADERS = {
    "Party_ID",
    "Party_Name",
    "Party_Type",
    "GSTIN",
    "Phone",
    "Address",
    "State",
    "State_Code",
    "Is_Active",
    "Created_At",
    "Updated_At",
}

# These names match the existing Audit_Log worksheet.
AUDIT_HEADERS = [
    "Audit_ID",
    "Event_Time",
    "Username",
    "Action",
    "Entity_Type",
    "Entity_ID",
    "Reason",
    "Before_Data",
    "After_Data",
]

MAX_PARTY_NAME_LENGTH = 150
MAX_PHONE_LENGTH = 25
MAX_ADDRESS_LENGTH = 500
MAX_STATE_LENGTH = 100
MAX_REASON_LENGTH = 250

GSTIN_PATTERN = re.compile(
    r"[0-9]{2}[A-Z]{5}[0-9]{4}"
    r"[A-Z][1-9A-Z]Z[0-9A-Z]"
)

VALID_PARTY_TYPES = {"GST", "WITHOUT_GST"}


class PartyServiceError(Exception):
    """Expected validation or workbook-storage error."""


def _now() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _get_path() -> Path:
    return Path(Config.WORKBOOK_PATH)


def _headers(ws) -> list[str]:
    """Return row-one headers and reject duplicate non-empty names."""
    headers = [
        _text(ws.cell(row=1, column=col).value)
        for col in range(1, ws.max_column + 1)
    ]

    while headers and not headers[-1]:
        headers.pop()

    if not headers or not any(headers):
        raise PartyServiceError(
            f"Worksheet '{ws.title}' does not contain a header row."
        )

    nonempty = [header for header in headers if header]
    if len(nonempty) != len(set(nonempty)):
        raise PartyServiceError(
            f"Worksheet '{ws.title}' has duplicate column headers."
        )

    return headers


def _require_headers(ws, required: set[str]) -> list[str]:
    headers = _headers(ws)
    missing = sorted(required - set(headers))

    if missing:
        raise PartyServiceError(
            f"Worksheet '{ws.title}' is missing required columns: "
            + ", ".join(missing)
        )

    return headers


def _json_safe(data: dict[str, Any] | None) -> str:
    """Serialize a record for audit history, excluding internal keys."""
    if data is None:
        return ""

    clean_data = {
        key: value
        for key, value in data.items()
        if not str(key).startswith("_")
    }

    return json.dumps(
        clean_data,
        ensure_ascii=False,
        default=str,
    )


def _validate_party(data: dict[str, Any]) -> dict[str, str]:
    """Validate and normalize the fields accepted by Party Master."""
    party_name = _text(data.get("Party_Name"))
    party_type = _text(data.get("Party_Type")).upper()
    gstin = _text(data.get("GSTIN")).upper()

    if not party_name:
        raise PartyServiceError("Party name is required.")

    if len(party_name) > MAX_PARTY_NAME_LENGTH:
        raise PartyServiceError(
            f"Party name must not exceed {MAX_PARTY_NAME_LENGTH} characters."
        )

    if party_type not in VALID_PARTY_TYPES:
        raise PartyServiceError(
            "Select GST Party or Without GST Party."
        )

    if party_type == "GST":
        if not gstin:
            raise PartyServiceError(
                "GSTIN is required for a GST Party."
            )

        if not GSTIN_PATTERN.fullmatch(gstin):
            raise PartyServiceError(
                "GSTIN must have a valid 15-character format."
            )
    else:
        gstin = ""

    phone = _text(data.get("Phone"))
    if len(phone) > MAX_PHONE_LENGTH:
        raise PartyServiceError(
            f"Phone number must not exceed {MAX_PHONE_LENGTH} characters."
        )

    address = _text(data.get("Address"))
    if len(address) > MAX_ADDRESS_LENGTH:
        raise PartyServiceError(
            f"Address must not exceed {MAX_ADDRESS_LENGTH} characters."
        )

    state = _text(data.get("State"))
    if len(state) > MAX_STATE_LENGTH:
        raise PartyServiceError(
            f"State must not exceed {MAX_STATE_LENGTH} characters."
        )

    state_code = _text(data.get("State_Code"))
    if state_code and not re.fullmatch(r"[0-9]{2}", state_code):
        raise PartyServiceError(
            "State Code must contain exactly 2 digits."
        )

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
    if "Customers" not in wb.sheetnames:
        raise PartyServiceError(
            "Customers worksheet was not found in the workbook."
        )

    ws = wb["Customers"]
    headers = _require_headers(ws, PARTY_REQUIRED_HEADERS)
    records: list[dict[str, Any]] = []

    for row_number in range(2, ws.max_row + 1):
        values = [
            ws.cell(row=row_number, column=col).value
            for col in range(1, len(headers) + 1)
        ]

        if not any(value is not None for value in values):
            continue

        record = dict(zip(headers, values))
        record["_row_number"] = row_number
        records.append(record)

    return records


def _find_party_row(
    ws,
    party_id: str,
    headers: list[str],
) -> int | None:
    id_column = headers.index("Party_ID") + 1

    for row_number in range(2, ws.max_row + 1):
        value = _text(ws.cell(row=row_number, column=id_column).value)
        if value == party_id:
            return row_number

    return None


def _append_audit(
    wb,
    action: str,
    party_id: str,
    reason: str,
    before: dict[str, Any] | None,
    after: dict[str, Any] | None,
    username: str,
) -> None:
    """Append an audit record using the worksheet's actual column positions."""
    if "Audit_Log" not in wb.sheetnames:
        raise PartyServiceError(
            "Audit_Log worksheet is missing. No changes were saved."
        )

    ws = wb["Audit_Log"]
    headers = _headers(ws)

    missing = [name for name in AUDIT_HEADERS if name not in headers]
    if missing:
        raise PartyServiceError(
            "Audit_Log is missing required columns: "
            + ", ".join(missing)
            + ". Existing workbook headers were not changed."
        )

    record = {
        "Audit_ID": "AUD-" + uuid.uuid4().hex[:12].upper(),
        "Event_Time": _now(),
        "Username": _text(username) or "Administrator",
        "Action": action,
        "Entity_Type": "PARTY",
        "Entity_ID": party_id,
        "Reason": _text(reason),
        "Before_Data": _json_safe(before),
        "After_Data": _json_safe(after),
    }

    # Map by header name. Extra blank columns are left untouched.
    next_row = ws.max_row + 1
    for header, value in record.items():
        column_number = headers.index(header) + 1
        ws.cell(
            row=next_row,
            column=column_number,
            value=value,
        )


def _save_atomic(wb, path: Path) -> None:
    """
    Save to a unique temporary file in the same directory, validate it,
    create a timestamped backup, then replace the original workbook.

    This reduces the risk of a partial save. It does not provide full
    multi-process transaction or locking guarantees.
    """
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)

    if not path.exists():
        raise PartyServiceError(
            f"Workbook file was not found: {path}"
        )

    temporary_path: Path | None = None

    try:
        with tempfile.NamedTemporaryFile(
            prefix=f".{path.stem}_",
            suffix=path.suffix,
            dir=path.parent,
            delete=False,
        ) as temp_file:
            temporary_path = Path(temp_file.name)

        wb.save(temporary_path)

        # Verify the generated workbook can be opened before replacing.
        check_wb = load_workbook(temporary_path, read_only=True)
        try:
            if "Customers" not in check_wb.sheetnames:
                raise PartyServiceError(
                    "Saved workbook validation failed: Customers sheet missing."
                )
            if "Audit_Log" not in check_wb.sheetnames:
                raise PartyServiceError(
                    "Saved workbook validation failed: Audit_Log sheet missing."
                )
        finally:
            check_wb.close()

        # Preserve the previous workbook before replacing it.
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        backup_path = path.with_name(
            f"{path.stem}.backup_{timestamp}{path.suffix}"
        )
        shutil.copy2(path, backup_path)

        # Temporary file is in the same directory, allowing atomic replacement
        # on typical local Windows filesystems.
        os.replace(temporary_path, path)
        temporary_path = None

    except PartyServiceError:
        raise

    except PermissionError as exc:
        raise PartyServiceError(
            "Workbook save nahi ho paya. WPS/Excel mein workbook close "
            "karein aur dobara try karein. File permissions bhi check karein."
        ) from exc

    except OSError as exc:
        raise PartyServiceError(
            "Workbook save nahi ho paya. Disk space, file permissions, "
            "backup access aur workbook lock check karein."
        ) from exc

    except Exception as exc:
        raise PartyServiceError(
            "Workbook save failed. Original file ko replace nahi kiya gaya "
            "agar temporary save/validation fail hua."
        ) from exc

    finally:
        if temporary_path is not None and temporary_path.exists():
            try:
                temporary_path.unlink()
            except OSError:
                pass


def _is_active(party: dict[str, Any]) -> bool:
    return _text(party.get("Is_Active")).lower() in {
        "true", "1", "yes", "active"
    }


def _is_duplicate_gstin(
    parties: list[dict[str, Any]],
    gstin: str,
    exclude_party_id: str | None = None,
) -> bool:
    if not gstin:
        return False

    for party in parties:
        if _text(party.get("Party_ID")) == exclude_party_id:
            continue

        if _text(party.get("GSTIN")).upper() == gstin.upper():
            return True

    return False


def list_parties(
    party_type: str | None = None,
    search: str = "",
    active_only: bool = True,
) -> list[dict[str, Any]]:
    """Return parties filtered by optional type, search text, and active status."""
    path = initialize_workbook()

    with _LOCK:
        wb = load_workbook(path, read_only=True, data_only=True)
        try:
            parties = _read_parties(wb)
        finally:
            wb.close()

    requested_type = _text(party_type).upper()
    if requested_type and requested_type not in VALID_PARTY_TYPES:
        raise PartyServiceError("Invalid party type filter.")

    query = _text(search).casefold()
    results: list[dict[str, Any]] = []

    searchable_fields = (
        "Party_ID",
        "Party_Name",
        "GSTIN",
        "Phone",
        "Contact_Person",
        "Address",
        "State",
    )

    for original in parties:
        party = dict(original)

        if active_only and not _is_active(party):
            continue

        if (
            requested_type
            and _text(party.get("Party_Type")).upper() != requested_type
        ):
            continue

        if query:
            searchable = " ".join(
                _text(party.get(field))
                for field in searchable_fields
            ).casefold()

            if query not in searchable:
                continue

        party.pop("_row_number", None)
        results.append(party)

    return sorted(
        results,
        key=lambda item: _text(item.get("Party_Name")).casefold(),
    )


def get_party(party_id: str) -> dict[str, Any] | None:
    """Return one party by ID, including inactive parties."""
    target = _text(party_id)
    if not target:
        return None

    for party in list_parties(active_only=False):
        if _text(party.get("Party_ID")) == target:
            return party

    return None


def create_party(
    data: dict[str, Any],
    username: str = "Administrator",
) -> str:
    """Validate, create, and audit a party."""
    clean = _validate_party(data)
    path = initialize_workbook()

    with _LOCK:
        wb = load_workbook(path)
        try:
            parties = _read_parties(wb)

            if (
                clean["Party_Type"] == "GST"
                and _is_duplicate_gstin(parties, clean["GSTIN"])
            ):
                raise PartyServiceError(
                    "This GSTIN already exists in Party Master."
                )

            max_number = 0
            existing_ids: set[str] = set()

            for party in parties:
                existing_id = _text(party.get("Party_ID"))
                if existing_id:
                    existing_ids.add(existing_id)

                match = re.fullmatch(r"P(\d+)", existing_id)
                if match:
                    max_number = max(max_number, int(match.group(1)))

            next_number = max_number + 1
            party_id = f"P{next_number:06d}"

            while party_id in existing_ids:
                next_number += 1
                party_id = f"P{next_number:06d}"

            created_at = _now()
            record = {
                **clean,
                "Party_ID": party_id,
                "Is_Active": True,
                "Created_At": created_at,
                "Updated_At": created_at,
            }

            ws = wb["Customers"]
            headers = _require_headers(ws, PARTY_REQUIRED_HEADERS)
            next_row = ws.max_row + 1

            for column_number, header in enumerate(headers, start=1):
                if header and header in record:
                    ws.cell(
                        row=next_row,
                        column=column_number,
                        value=record[header],
                    )

            _append_audit(
                wb,
                "CREATE",
                party_id,
                "Party created",
                None,
                record,
                username,
            )

            _save_atomic(wb, Path(path))
            return party_id
        finally:
            wb.close()


def update_party(
    party_id: str,
    data: dict[str, Any],
    reason: str,
    username: str = "Administrator",
) -> None:
    """Update party fields, preserve its ID/creation time, and record an audit."""
    target = _text(party_id)
    reason = _text(reason)

    if not target:
        raise PartyServiceError("Party ID is required.")

    if not reason:
        raise PartyServiceError(
            "A reason is required when editing a party."
        )

    if len(reason) > MAX_REASON_LENGTH:
        raise PartyServiceError(
            f"Reason for change must not exceed {MAX_REASON_LENGTH} characters."
        )

    clean = _validate_party(data)
    path = initialize_workbook()

    with _LOCK:
        wb = load_workbook(path)
        try:
            if "Customers" not in wb.sheetnames:
                raise PartyServiceError(
                    "Customers worksheet was not found in the workbook."
                )

            ws = wb["Customers"]
            headers = _require_headers(ws, PARTY_REQUIRED_HEADERS)
            row_number = _find_party_row(ws, target, headers)

            if row_number is None:
                raise PartyServiceError("Party was not found.")

            before = {
                header: ws.cell(
                    row=row_number,
                    column=column_number,
                ).value
                for column_number, header in enumerate(headers, start=1)
                if header
            }

            parties = _read_parties(wb)
            if (
                clean["Party_Type"] == "GST"
                and _is_duplicate_gstin(
                    parties,
                    clean["GSTIN"],
                    exclude_party_id=target,
                )
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

            for column_number, header in enumerate(headers, start=1):
                if header and header in updated:
                    ws.cell(
                        row=row_number,
                        column=column_number,
                        value=updated[header],
                    )

            _append_audit(
                wb,
                "UPDATE",
                target,
                reason,
                before,
                updated,
                username,
            )

            _save_atomic(wb, Path(path))
        finally:
            wb.close()


def _find_party_references(wb, party_id: str) -> list[str]:
    """
    Conservatively block deletion if the exact party ID occurs in another
    worksheet. Review any reported reference before deleting linked records.
    """
    references: list[str] = []

    for ws in wb.worksheets:
        if ws.title in {"Customers", "Audit_Log"}:
            continue

        found_in_sheet = False

        for row in ws.iter_rows(min_row=1):
            for cell in row:
                if _text(cell.value) == party_id:
                    references.append(f"{ws.title}!{cell.coordinate}")
                    found_in_sheet = True
                    break

            if found_in_sheet:
                break

    return references


def delete_party(
    party_id: str,
    username: str = "Administrator",
) -> None:
    """
    Permanently delete an unreferenced party and preserve a deletion audit.
    This operation is irreversible unless the workbook is restored from backup.
    """
    target = _text(party_id)
    if not target:
        raise PartyServiceError("Party ID is required.")

    path = initialize_workbook()

    with _LOCK:
        wb = load_workbook(path)
        try:
            if "Customers" not in wb.sheetnames:
                raise PartyServiceError(
                    "Customers worksheet was not found in the workbook."
                )

            ws = wb["Customers"]
            headers = _require_headers(ws, PARTY_REQUIRED_HEADERS)
            row_number = _find_party_row(ws, target, headers)

            if row_number is None:
                raise PartyServiceError("Party was not found.")

            before = {
                header: ws.cell(
                    row=row_number,
                    column=column_number,
                ).value
                for column_number, header in enumerate(headers, start=1)
                if header
            }

            references = _find_party_references(wb, target)
            if references:
                preview = ", ".join(references[:5])
                remaining = len(references) - 5
                extra = (
                    f" and {remaining} more worksheet reference(s)"
                    if remaining > 0
                    else ""
                )
                raise PartyServiceError(
                    "Party cannot be deleted because its ID is referenced "
                    f"in other workbook sheets: {preview}{extra}. "
                    "Review linked records first."
                )

            # Write the audit entry before removing the customer row.
            _append_audit(
                wb,
                "DELETE",
                target,
                "Permanent deletion confirmed by user",
                before,
                None,
                username,
            )

            ws.delete_rows(row_number, 1)
            _save_atomic(wb, Path(path))
        finally:
            wb.close()
