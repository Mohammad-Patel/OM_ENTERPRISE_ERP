
from __future__ import annotations

import os
import shutil
from datetime import datetime
from pathlib import Path

from app.config import Config


BACKUP_PREFIX = "om_enterprise_backup_"
DEFAULT_KEEP_COUNT = 30


class BackupServiceError(Exception):
    """Raised when a workbook backup cannot be completed."""


def create_backup(
    source_path: str | Path | None = None,
    backup_dir: str | Path | None = None,
) -> Path:
    """
    Create a timestamped copy of the Excel workbook.

    The original workbook is never modified by this function.
    Returns the path of the new backup.
    """
    source = Path(source_path or Config.WORKBOOK_PATH)
    destination_dir = Path(backup_dir or Config.BACKUP_DIR)

    if not source.exists():
        raise BackupServiceError(
            f"Workbook not found; backup not created: {source}"
        )

    if not source.is_file() or source.stat().st_size == 0:
        raise BackupServiceError(
            f"Workbook is empty or invalid: {source}"
        )

    try:
        destination_dir.mkdir(parents=True, exist_ok=True)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        destination = destination_dir / (
            f"{BACKUP_PREFIX}{timestamp}.xlsx"
        )
        temporary = destination.with_name(
            destination.stem + ".tmp.xlsx"
        )

        try:
            shutil.copy2(source, temporary)
            os.replace(temporary, destination)
        finally:
            if temporary.exists():
                try:
                    temporary.unlink()
                except OSError:
                    pass

        if not destination.exists() or destination.stat().st_size == 0:
            raise BackupServiceError(
                "Backup file was not created correctly."
            )

        return destination

    except BackupServiceError:
        raise
    except Exception as exc:
        raise BackupServiceError(
            f"Could not create workbook backup: {exc}"
        ) from exc


def list_backups(
    backup_dir: str | Path | None = None,
) -> list[Path]:
    """List generated backups, newest first."""
    directory = Path(backup_dir or Config.BACKUP_DIR)

    if not directory.exists():
        return []

    return sorted(
        (
            path
            for path in directory.glob(f"{BACKUP_PREFIX}*.xlsx")
            if path.is_file()
        ),
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )


def prune_old_backups(
    keep_count: int = DEFAULT_KEEP_COUNT,
    backup_dir: str | Path | None = None,
) -> int:
    """
    Delete older backups created by this service only.

    Returns the number of deleted backup files.
    """
    if keep_count < 1:
        raise ValueError("keep_count must be at least 1.")

    backups = list_backups(backup_dir)
    deleted_count = 0

    for old_backup in backups[keep_count:]:
        try:
            old_backup.unlink()
            deleted_count += 1
        except OSError as exc:
            raise BackupServiceError(
                f"Could not remove old backup: {old_backup}"
            ) from exc

    return deleted_count
