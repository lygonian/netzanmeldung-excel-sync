from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path
from typing import Sequence

from .config import (
    DEFAULT_MIN_START_DATE,
    DEFAULT_TABLE_NAME,
    ENV_KONTAKTE_CSV,
    ENV_MIN_START_DATE,
    ENV_OUTPUT_PATH,
    ENV_PROJEKTE_CSV,
    ENV_TABLE_NAME,
    ENV_TERMINE_CSV,
    ENV_WORKBOOK_PATH,
)
from .demo import run_demo
from .sync import SyncResult, sync_netzanmeldungen


def main(argv: Sequence[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        if args.demo:
            result = run_demo(args.demo_output_dir)
        else:
            result = sync_netzanmeldungen(
                termine_csv=_required_path(args.termine_csv, ENV_TERMINE_CSV, parser),
                projekte_csv=_required_path(args.projekte_csv, ENV_PROJEKTE_CSV, parser),
                kontakte_csv=_required_path(args.kontakte_csv, ENV_KONTAKTE_CSV, parser),
                workbook_path=_required_path(args.workbook, ENV_WORKBOOK_PATH, parser),
                output_path=args.output or _optional_path(ENV_OUTPUT_PATH),
                table_name=args.table_name,
                min_start_date=args.min_start_date,
            )
    except Exception as exc:
        print(f"Fehler: {exc}", file=sys.stderr)
        return 1

    _print_result(result)
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Netzanmeldungs-Excel aus CSV-Exports aktualisieren.")
    parser.add_argument("--termine-csv", type=Path, help="CSV mit Terminen.")
    parser.add_argument("--projekte-csv", type=Path, help="CSV mit Projekten.")
    parser.add_argument("--kontakte-csv", type=Path, help="CSV mit Kontakten.")
    parser.add_argument("--workbook", type=Path, help="Bestehende .xlsx-Datei mit Tabelle.")
    parser.add_argument("--output", type=Path, help="Zielpfad fuer aktualisierte .xlsx-Datei.")
    parser.add_argument("--table-name", default=os.getenv(ENV_TABLE_NAME, DEFAULT_TABLE_NAME))
    parser.add_argument("--min-start-date", default=os.getenv(ENV_MIN_START_DATE, DEFAULT_MIN_START_DATE))
    parser.add_argument("--demo", action="store_true", help="Synthetische Demo-Daten verarbeiten.")
    parser.add_argument("--demo-output-dir", type=Path, default=Path("demo_output"))
    return parser


def _required_path(value: Path | None, env_name: str, parser: argparse.ArgumentParser) -> Path:
    if value is not None:
        return value
    env_value = os.getenv(env_name)
    if env_value:
        return Path(env_value)
    parser.error(f"Fehlender Pfad: Argument setzen oder {env_name} definieren.")
    raise AssertionError("argparse exits before this line")


def _optional_path(env_name: str) -> Path | None:
    env_value = os.getenv(env_name)
    return Path(env_value) if env_value else None


def _print_result(result: SyncResult) -> None:
    print("Netzanmeldung Excel Sync abgeschlossen")
    print(f"eligible_projects: {result.eligible_projects}")
    print(f"added_rows: {result.added_rows}")
    print(f"deleted_rows: {result.deleted_rows}")
    print(f"table_ref: {result.table_ref}")
    print(f"output_path: {result.output_path}")
