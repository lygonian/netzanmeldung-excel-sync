"""Excel sync helpers for public netzanmeldung demo data."""

from .sync import SyncResult, prepare_new_rows, sync_netzanmeldungen, update_workbook

__all__ = [
    "SyncResult",
    "prepare_new_rows",
    "sync_netzanmeldungen",
    "update_workbook",
]

__version__ = "0.1.0"
