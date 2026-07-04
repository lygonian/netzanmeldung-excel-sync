# Netzanmeldung Excel Sync

Bereinigte Python-Version eines Excel-Syncs fuer Netzanmeldungslisten. Das Tool
uebernimmt geeignete DC-Montage-Termine aus CSV-Exports, verknuepft sie mit
Projekt- und Kontaktdaten und erweitert eine bestehende Excel-Tabelle, ohne
private API-Zugaenge oder echte Unternehmensdaten zu benoetigen.

## Funktionen

- filtert DC-Montage-Termine ab einem Stichtag und ignoriert geloeschte Termine,
- verknuepft Termine mit Projekt- und Kontaktdaten ueber IDs,
- fuegt noch fehlende Projekte in eine Excel-Tabelle ein,
- entfernt leere Dubletten, wenn dieselbe Projektnummer bereits eine gefuellte
  Einspeiseranfrage hat,
- aktualisiert den Excel-Tabellenbereich und erhaelt vorhandene Zeilenformate,
- enthaelt einen reproduzierbaren Demo-Modus mit synthetischen Daten.

## Projektstruktur

- `src/netzanmeldung_excel_sync/`: Paket mit Sync-Logik, CLI und Demo-Erzeugung
- `src/netzanmeldung_excel_sync/sample_data/`: synthetische Demo-CSV-Dateien
- `tests/`: fokussierte Tests fuer Filterung, Merge, Excel-Update und Cleanup
- `docs/datenschema.md`: erwartete Eingabespalten und Excel-Struktur
- `.env.example`: Vorlage fuer optionale Umgebungsvariablen

## Voraussetzungen

- Python 3.11 oder neuer
- Windows, macOS oder Linux
- Excel-Datei im `.xlsx`-Format mit einer Tabelle, standardmaessig `Tabelle1`

## Installation

```powershell
python -m venv .venv
.\.venv\Scripts\python -m pip install --upgrade pip
.\.venv\Scripts\python -m pip install -e .[dev]
```

Unter macOS/Linux entsprechend:

```bash
python3 -m venv .venv
. .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -e ".[dev]"
```

## Konfiguration

Die CLI kann Pfade direkt per Argument oder ueber Umgebungsvariablen lesen. Eine
`.env`-Datei wird nicht automatisch geladen; `.env.example` dient als Vorlage
fuer Shell-Variablen oder eigene lokale Skripte.

Relevante Variablen:

- `NETZANMELDUNG_TERMINE_CSV`
- `NETZANMELDUNG_PROJEKTE_CSV`
- `NETZANMELDUNG_KONTAKTE_CSV`
- `NETZANMELDUNG_WORKBOOK_PATH`
- `NETZANMELDUNG_OUTPUT_PATH`
- `NETZANMELDUNG_TABLE_NAME`
- `NETZANMELDUNG_MIN_START_DATE`

## Ausfuehrung

```powershell
netzanmeldung-excel-sync `
  --termine-csv inputs\termine.csv `
  --projekte-csv inputs\projekte.csv `
  --kontakte-csv inputs\kontakte.csv `
  --workbook inputs\netzanmeldungen.xlsx `
  --output output\netzanmeldungen_aktualisiert.xlsx
```

Alternativ als Modul:

```powershell
python -m netzanmeldung_excel_sync --demo
```

## Beispiel Oder Demo

Der Demo-Modus erzeugt lokal eine synthetische Excel-Datei und verarbeitet die
mitgelieferten Demo-CSV-Dateien:

```powershell
netzanmeldung-excel-sync --demo
```

Die erzeugten Dateien liegen in `demo_output/` und sind absichtlich per
`.gitignore` ausgeschlossen.

## Tests

```powershell
python -m pytest
```

## Einschraenkungen

Die oeffentliche Version enthaelt keinen internen Hero-Sync, keine produktiven
API-Zugaenge, keine privaten Pfade und keine echten Netzanmeldungsdaten. Sie
erwartet bereits exportierte CSV-Dateien mit dem in `docs/datenschema.md`
beschriebenen Schema.

## Datenschutz

Alle mitgelieferten Beispieldaten sind synthetisch. Sie enthalten keine echten
Personen-, Kunden-, Projekt-, Vertrags- oder Unternehmensdaten.
