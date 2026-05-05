#!/usr/bin/env python3
"""
csv_importer.py
Lead Reactivation — CSV → Google Sheet Importer (Admin Tool)

Run:
    pip install -r requirements.txt
    python csv_importer.py
"""

import csv
import sys
from pathlib import Path

import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


# ─── FIXED CONFIG ─────────────────────────────────────────────────────────────

CREDENTIALS_FILE = "credentials.json"
DRIVE_FOLDER_ID  = "1B0htenH8mGMhjzq1C-c0fFwls_F5OsvE"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

STANDARD_HEADERS = [
    "first_name",
    "last_name",
    "phone_number",
    "lead_source",
    "original_interest",
    "date_added",
    "agent_name",
    "transfer_number",
    "call_status",
    "interest_level",
    "timeline",
    "pre_approved",
    "working_with_agent",
    "transfer_attempted",
    "notes",
    "next_action",
    "last_called",
    "recording",
    "attempt_count",
]

# Columns filled by the calling system — leave blank on import except attempt_count
SYSTEM_COLUMNS = {
    "call_status", "interest_level", "timeline", "pre_approved",
    "working_with_agent", "transfer_attempted", "notes", "next_action",
    "last_called", "recording",
}

# Flexible aliases for each CSV-sourced field (all lowercase for matching)
FIELD_ALIASES: dict[str, list[str]] = {
    "first_name": [
        "first_name", "first name", "firstname",
        "name", "full_name", "full name",
    ],
    "last_name": [
        "last_name", "last name", "lastname", "surname",
    ],
    "phone_number": [
        "phone", "phone_number", "phone number",
        "mobile", "contact", "number",
    ],
    "lead_source": [
        "lead_source", "lead source", "source", "platform",
    ],
    "original_interest": [
        "original_interest", "original interest", "interest",
        "requirement", "property", "enquiry", "inquiry",
    ],
    "date_added": [
        "date_added", "date added", "date", "created_at", "created",
    ],
}

# The 6 fields we import from the CSV (everything else comes from user input
# or is left blank for the calling system)
CSV_IMPORT_FIELDS = list(FIELD_ALIASES.keys())


# ─── STEP 1 — COLLECT INPUTS ──────────────────────────────────────────────────

def get_inputs() -> dict:
    """Prompt the user for the 4 required inputs."""
    print()
    csv_path        = input("  CSV file path     : ").strip()
    client_name     = input("  Client name       : ").strip()
    agent_name      = input("  Agent name        : ").strip()
    transfer_number = input("  Transfer number   : ").strip()
    print()

    if not csv_path:
        _abort("CSV file path is required.")
    if not client_name:
        _abort("Client name is required.")

    return {
        "csv_path":        csv_path,
        "client_name":     client_name,
        "agent_name":      agent_name,
        "transfer_number": transfer_number,
    }


# ─── STEP 2 — READ CSV ────────────────────────────────────────────────────────

def read_csv(csv_path: str) -> tuple[list[str], list[dict]]:
    """
    Read a local CSV file.
    Returns (headers, rows_as_dicts).
    Handles UTF-8 BOM and skips fully-blank rows.
    """
    path = Path(csv_path)
    if not path.exists():
        _abort(f"File not found: {csv_path}")

    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        headers = list(reader.fieldnames or [])
        rows = [
            dict(row)
            for row in reader
            if any(str(v).strip() for v in row.values())
        ]

    if not headers:
        _abort("CSV has no header row.")
    if not rows:
        _abort("CSV has no data rows.")

    return headers, rows


# ─── STEP 3 — AUTO MAP FIELDS ─────────────────────────────────────────────────

def map_fields(csv_headers: list[str]) -> dict[str, str | None]:
    """
    Build { standard_field: matching_csv_column } using case-insensitive
    alias matching. Returns None for fields not found in the CSV.
    """
    # Normalise every CSV header to lowercase for comparison
    lookup: dict[str, str] = {h.strip().lower(): h for h in csv_headers}

    mapping: dict[str, str | None] = {}
    for std_field, aliases in FIELD_ALIASES.items():
        matched = next(
            (lookup[alias] for alias in aliases if alias in lookup),
            None
        )
        mapping[std_field] = matched

    return mapping


def _needs_name_split(mapping: dict[str, str | None]) -> bool:
    """
    True when no explicit last_name column was found — which means the
    first_name column likely holds a full name like "John Smith".
    """
    return mapping.get("last_name") is None


def _split_name(full_name: str) -> tuple[str, str]:
    """'John Smith Jr'  →  ('John', 'Smith Jr')"""
    parts = full_name.strip().split(None, 1)
    return (parts[0] if parts else ""), (parts[1] if len(parts) > 1 else "")


# ─── STEP 4 — BUILD OUTPUT ROW ────────────────────────────────────────────────

def transform_row(
    csv_row: dict,
    mapping: dict[str, str | None],
    agent_name: str,
    transfer_number: str,
    split_names: bool,
) -> list[str]:
    """
    Convert one CSV dict into a list that matches STANDARD_HEADERS order.

    Logic:
      - CSV fields  : pulled via mapping, empty string if unmapped
      - Name split  : if no last_name column, split first_name on first space
      - User fields : agent_name, transfer_number filled from prompts
      - System cols : left blank (filled later by the calling system)
      - attempt_count: always "0"
    """
    values: dict[str, str] = {}

    # Pull the 6 CSV-sourced fields
    for std_field in CSV_IMPORT_FIELDS:
        src_col = mapping.get(std_field)
        values[std_field] = str(csv_row.get(src_col, "")).strip() if src_col else ""

    # Split full name into first + last when no explicit last_name column
    if split_names and " " in values.get("first_name", ""):
        values["first_name"], values["last_name"] = _split_name(values["first_name"])

    # User-supplied fields
    values["agent_name"]      = agent_name
    values["transfer_number"] = transfer_number

    # System columns → blank
    for col in SYSTEM_COLUMNS:
        values[col] = ""

    # attempt_count → always 0 on first import
    values["attempt_count"] = "0"

    # Return in exact STANDARD_HEADERS order
    return [values.get(col, "") for col in STANDARD_HEADERS]


# ─── STEP 5 — CREATE GOOGLE SHEET ─────────────────────────────────────────────

def authenticate(creds_file: str) -> tuple[gspread.Client, object]:
    """
    Load service account credentials and return (gspread_client, drive_service).
    Uses the credentials file for BOTH gspread and the Drive API.
    """
    if not Path(creds_file).exists():
        _abort(
            f"{creds_file} not found.\n"
            "  1. Go to Google Cloud Console > IAM > Service Accounts\n"
            "  2. Create / select a service account\n"
            "  3. Keys > Add Key > JSON > download\n"
            "  4. Rename the file to credentials.json and place it here."
        )

    # gspread.service_account() is the most reliable method for service accounts
    gc    = gspread.service_account(filename=creds_file)
    creds = Credentials.from_service_account_file(creds_file, scopes=SCOPES)
    drive = build("drive", "v3", credentials=creds)

    return gc, drive


def create_google_sheet(
    client_name: str,
    gc: gspread.Client,
    drive,
) -> tuple[gspread.Spreadsheet, str]:
    """
    Create 'Lead Reactivation - {client_name}' inside DRIVE_FOLDER_ID.
    Returns (spreadsheet, web_view_url).
    """
    sheet_name = f"Lead Reactivation - {client_name}"

    metadata = {
        "name":     sheet_name,
        "mimeType": "application/vnd.google-apps.spreadsheet",
        "parents":  [DRIVE_FOLDER_ID],
    }

    try:
        created = drive.files().create(
            body=metadata,
            fields="id,webViewLink",
        ).execute()
    except HttpError as exc:
        _abort(f"Google Drive API error:\n  {exc}")

    spreadsheet = gc.open_by_key(created["id"])
    return spreadsheet, created["webViewLink"]


# ─── STEP 6 — WRITE HEADERS + DATA ────────────────────────────────────────────

def write_to_sheet(
    spreadsheet: gspread.Spreadsheet,
    data_rows: list[list[str]],
) -> None:
    """
    Write standardised headers (row 1, bold, frozen) then all data rows.
    Appends in chunks of 500 to stay inside API payload limits.
    """
    ws = spreadsheet.sheet1
    ws.clear()

    # Row 1: headers
    ws.update("A1", [STANDARD_HEADERS])
    ws.freeze(rows=1)

    # Format header row: bold + light grey background
    last_col = _col_letter(len(STANDARD_HEADERS))
    ws.format(
        f"A1:{last_col}1",
        {
            "textFormat": {"bold": True},
            "backgroundColor": {"red": 0.91, "green": 0.91, "blue": 0.91},
        },
    )

    # Rows 2+: data (chunked)
    for i in range(0, len(data_rows), 500):
        ws.append_rows(data_rows[i : i + 500], value_input_option="RAW")


# ─── MAIN ─────────────────────────────────────────────────────────────────────

def main() -> None:
    _banner()

    # ── Inputs ───────────────────────────────────────────────────────────────
    inputs = get_inputs()
    csv_path        = inputs["csv_path"]
    client_name     = inputs["client_name"]
    agent_name      = inputs["agent_name"]
    transfer_number = inputs["transfer_number"]

    # ── Read CSV ─────────────────────────────────────────────────────────────
    print("[1] Reading CSV...")
    headers, rows = read_csv(csv_path)
    print(f"     {len(rows)} rows  |  {len(headers)} columns")
    print(f"     Columns: {', '.join(headers)}")

    # ── Map fields ───────────────────────────────────────────────────────────
    print("\n[2] Mapping fields to standard schema...")
    mapping     = map_fields(headers)
    split_names = _needs_name_split(mapping)

    for std_field in CSV_IMPORT_FIELDS:
        src = mapping[std_field]
        tag = f'<-- "{src}"' if src else "<-- (not found, will be empty)"
        print(f"     {std_field:<22}  {tag}")

    if split_names and mapping.get("first_name"):
        print(f'\n     NOTE: No last_name column found.')
        print(f'           Values in "{mapping["first_name"]}" will be split into first + last.')

    # ── Authenticate ─────────────────────────────────────────────────────────
    print("\n[3] Authenticating with Google...")
    gc, drive = authenticate(CREDENTIALS_FILE)
    print("     OK - Authenticated")

    # ── Create sheet ─────────────────────────────────────────────────────────
    print(f"\n[4] Creating Google Sheet...")
    spreadsheet, sheet_url = create_google_sheet(client_name, gc, drive)
    print(f"     OK - Created: Lead Reactivation - {client_name}")
    print(f"     ID: {spreadsheet.id}")

    # ── Build data rows ───────────────────────────────────────────────────────
    print("\n[5] Processing rows...")
    data_rows = [
        transform_row(row, mapping, agent_name, transfer_number, split_names)
        for row in rows
    ]
    print(f"     OK - {len(data_rows)} rows ready")

    # ── Write to sheet ────────────────────────────────────────────────────────
    print(f"\n[6] Writing to Google Sheet...")
    write_to_sheet(spreadsheet, data_rows)
    print(f"     OK - Done")

    # ── Summary ──────────────────────────────────────────────────────────────
    _divider()
    print("  >> IMPORT COMPLETE")
    _divider()
    print(f"  Sheet name    : Lead Reactivation - {client_name}")
    print(f"  Rows imported : {len(data_rows)}")
    print()
    print("  Sheet URL:")
    print(f"  {sheet_url}")
    print()
    print("  Use this URL in your dashboard for this client.")
    _divider()


# ─── UTILITIES ────────────────────────────────────────────────────────────────

def _col_letter(n: int) -> str:
    """1-indexed column number → spreadsheet column letter (19 → S)."""
    result = ""
    while n:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


def _banner() -> None:
    _divider()
    print("  Lead Reactivation - CSV Importer  (Admin Tool)")
    _divider()


def _divider() -> None:
    print("\n" + "-" * 52)


def _abort(msg: str) -> None:
    print(f"\n  ERROR: {msg}\n")
    sys.exit(1)


# ─── ENTRY POINT ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    main()
