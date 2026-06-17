from __future__ import annotations
from pathlib import Path

STORE_ID_COLUMN = "Code_de_magasin"

FILE_COLUMNS = [
    "Code_de_magasin",
    "Tag Pays",
    "status",
    "Nom_etablissement",
    "Ligne_Adresse_1",
    "Localite",
    "pos_zipcode",
    "Latitude",
    "Longitude",
    "Siteweb",
    "pos_schedule_sunday",
    "pos_schedule_monday",
    "pos_schedule_tuesday",
    "pos_schedule_wednesday",
    "pos_schedule_thursday",
    "pos_schedule_friday",
    "pos_schedule_saturday",
    "Description_etablissement",
]

SYNC_FIELDS = [
    "Nom_etablissement",
    "Siteweb",
    "Ligne_Adresse_1",
    "Localite",
    "pos_zipcode",
    "Tag Pays",
    "Latitude",
    "Longitude",
    "Description_etablissement",
    "pos_schedule_sunday",
    "pos_schedule_monday",
    "pos_schedule_tuesday",
    "pos_schedule_wednesday",
    "pos_schedule_thursday",
    "pos_schedule_friday",
    "pos_schedule_saturday",
]

SCHEDULE_COLUMNS = [
    "pos_schedule_sunday",
    "pos_schedule_monday",
    "pos_schedule_tuesday",
    "pos_schedule_wednesday",
    "pos_schedule_thursday",
    "pos_schedule_friday",
    "pos_schedule_saturday",
]

OPTIONAL_COLUMNS = [
    "Numero_telephone",
    "Horaires_ouverture_exceptionnels",
]

DAY_SCHEDULE_CODES = [
    ("pos_schedule_monday", "MO"),
    ("pos_schedule_tuesday", "TU"),
    ("pos_schedule_wednesday", "WE"),
    ("pos_schedule_thursday", "TH"),
    ("pos_schedule_friday", "FR"),
    ("pos_schedule_saturday", "SA"),
    ("pos_schedule_sunday", "SU"),
]

RESULT_COLUMNS = [
    "SHOP_CODE",
    "ACTION",
    "STATUS",
    "LOCATION_NAME",
    "LOCATION_LANGUAGE",
    "PHONE_NUMBER",
    "EMAIL",
    "WEBSITE_URL",
    "ADR_STREET_NAME",
    "ADR_STREET_NUMBER",
    "ADR_ZIPCODE",
    "ADR_CITY",
    "ADR_COUNTRY",
    "ADR_LATITUDE",
    "ADR_LONGITUDE",
    "REGULAR_HOURS",
    "SPECIAL_HOURS (OPTIONAL)",
]

CLIENT_COLUMN_RENAME = {
    "Pays": "Tag Pays",
    "Ligne_adresse1": "Ligne_Adresse_1",
}

OUTPUT_DIR = Path("output")
OUTPUT_RESULTAT = OUTPUT_DIR / "resultat"
OUTPUT_LOG = OUTPUT_DIR / "log"
DEFAULT_COMPARE_OUTPUT = OUTPUT_RESULTAT / "fichier_synchronisation.xlsx"
DEFAULT_VALIDATE_OUTPUT = OUTPUT_LOG / "problemes_qualite.xlsx"
STATUS_OPEN = "OPEN"
STATUS_CLOSED = "CLOSED"

# Actions du flow IT (comparaison mois N vs mois N-1)
ACTION_NO_CHANGE = "no_change"
ACTION_UPDATE = "update"
ACTION_CREATE = "create"
ACTION_CLOSE = "close_permanently"
ACTION_COLUMN = "action"

EXPORT_ACTION_MAP = {
    ACTION_UPDATE: "UPDATE",
    ACTION_CREATE: "CREATE",
    ACTION_CLOSE: "UPDATE",
}

