DEFAULT_TABLE_NAME = "Tabelle1"
DEFAULT_MIN_START_DATE = "2025-04-01"
DEFAULT_MOUNT_TYPES = (
    "DC-Montage ohne Ger\u00fcst",
    "DC-Montage mit Ger\u00fcst",
    "DC-Montage ohne Geruest",
    "DC-Montage mit Geruest",
)

TERMINE_REQUIRED_COLUMNS = ("project_match_id", "start", "localized_type", "deleted")
PROJEKTE_REQUIRED_COLUMNS = ("id", "contact_id", "project_nr")
KONTAKTE_REQUIRED_COLUMNS = ("id", "first_name", "last_name")
OUTPUT_COLUMNS = ("project_nr", "first_name", "last_name", "start_datum")

WORKBOOK_HEADERS = (
    "project_nr",
    "first_name",
    "last_name",
    "start_datum",
    "Netzbetreiber",
    " Einspeiseranfrage",
    "installateur mail",
    "AC-fertig",
    "Anfrage Inbetriebsetzung",
    "Mastr",
    "Zaehlertausch",
    "Abschluss",
    "Kommentar",
    "Garantieverlaengerung",
)

ENV_TERMINE_CSV = "NETZANMELDUNG_TERMINE_CSV"
ENV_PROJEKTE_CSV = "NETZANMELDUNG_PROJEKTE_CSV"
ENV_KONTAKTE_CSV = "NETZANMELDUNG_KONTAKTE_CSV"
ENV_WORKBOOK_PATH = "NETZANMELDUNG_WORKBOOK_PATH"
ENV_OUTPUT_PATH = "NETZANMELDUNG_OUTPUT_PATH"
ENV_TABLE_NAME = "NETZANMELDUNG_TABLE_NAME"
ENV_MIN_START_DATE = "NETZANMELDUNG_MIN_START_DATE"
