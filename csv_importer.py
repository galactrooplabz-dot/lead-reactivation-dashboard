#!/usr/bin/env python3
"""
csv_importer.py
Lead Reactivation — CSV → Google Sheet Importer (Admin Tool)

Reads any local CSV, auto-maps columns to the standardized schema,
creates a new Google Sheet inside your Drive folder, and writes the data.

Usage:
    python csv_importer.py --csv leads.csv --client "Phoenix Real Estate Group"

Optional flags:
    --folder  DRIVE_FOLDER_ID   (overrides the default in this file)
    --creds   credentials.json  (path to service account key file)
"""

import csv
import sys
import argparse
from pathlib import Path

import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


# ─── CONFIGURATION ────────────────────────────────────────────────────────────

CREDENTIALS_FILE    = "credentials.json"
GOOGLE_DRIVE_FOLDER = "1B0htenH8mGMhjzq1C-c0fFwls_F5OsvE"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# Every sheet we create will have exactly these 19 columns in this order.
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

# These columns are filled by the calling system, not by the importer.
SYSTEM_COLUMNS = {
    "agent_name", "transfer_number", "call_status", "interest_level",
    "timeline", "pre_approved", "working_with_agent", "transfer_attempted",
    "notes", "next_action", "last_called", "recording", "attempt_count",
}

# Flexible aliases for each required import field.
# Keys must match the non-system entries in STANDARD_HEADERS.
FIELD_ALIASES: dict[str, list[str]] = {
    "first_name":        ["first_name", "firstname", "name", "full_name"],
    "last_name":         ["last_name", "lastname", "surname"],
    "phone_number":      ["phone", "phone_number", "mobile", "contact"],
    "original_interest": ["interest", "original_interest", "property", "requirement"],
    "lead_source":       ["source", "lead_source"],
    "date_added":        ["date", "date_added", "created_at"],
}


# ─── STEP 1 — AUTHENTICATION ──────────────────────────────────────────────────

def build_clients(creds_file: str) -> tuple[gspread.Client, object]:
    """
    Authenticate with Google using a service account key file.
    Returns (gspread_client, drive_service).
    """
    path = Path(creds_file)
    if not path.exists():
        _abort(
            f"credentials.json not found at: {creds_file}\n"
            "  Download your service account key from Google Cloud Console\n"
            "  and place it in the same folder as this script."
        )

    creds = Credentials.from_service_account_file(str(path), scopes=SCOPES)
    gc    = gspread.Client(auth=creds)
    drive = build("drive", "v3", credentials=creds)
    return gc, drive


# ─── STEP 2 — READ CSV ────────────────────────────────────────────────────────

def read_csv(csv_file_path: str) -> tuple[list[str], list[dict]]:
    """
    Read a local CSV file.
    Returns (header_row, list_of_row_dicts).
    Skips fully-empty rows. Handles UTF-8 BOM automatically.
    """
    path = Path(csv_file_path)
    if not path.exists():
        _abort(f"CSV file not found: {csv_file_path}")

    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        headers = list(reader.fieldnames or [])
        rows = [
            dict(row)
            for row in reader
            if any(str(v).strip() for v in row.values())
        ]

    if not headers:
        _abort("CSV file has no header row.")
    if not rows:
        _abort("CSV file has no data rows.")

    return headers, rows


# ─── STEP 3 — MAP FIELDS ──────────────────────────────────────────────────────

def map_fields(csv_headers: list[str]) -> dict[str, str | None]:
    """
    Auto-detect which CSV column corresponds to each standard field.
    Matching is case-insensitive and ignores leading/trailing whitespace.

    Returns { standard_field: csv_column_name } or { standard_field: None }
    when no match is found (will be left blank in the output sheet).
    """
    # Build a lowercase lookup: normalised_name → original_header_name
    normalised = {h.strip().lower(): h for h in csv_headers}
    mapping: dict[str, str | None] = {}

    for std_field, aliases in FIELD_ALIASES.items():
        match = next(
            (normalised[alias] for alias in aliases if alias in normalised),
            None
        )
        mapping[std_field] = match

    return mapping


# ─── STEP 4 — CREATE GOOGLE SHEET ─────────────────────────────────────────────

def create_google_sheet(
    client_name: str,
    folder_id: str,
    gc: gspread.Client,
    drive,
) -> tuple[gspread.Spreadsheet, str]:
    """
    Create a new Google Sheet named 'Lead Reactivation - {client_name}'
    directly inside the specified Drive folder.
    Returns (spreadsheet_object, web_view_url).
    """
    sheet_name = f"Lead Reactivation - {client_name}"

    metadata = {
        "name":     sheet_name,
        "mimeType": "application/vnd.google-apps.spreadsheet",
        "parents":  [folder_id],
    }

    try:
        created = drive.files().create(
            body=metadata,
            fields="id,webViewLink",
        ).execute()
    except HttpError as exc:
        _abort(f"Google Drive API error while creating sheet:\n  {exc}")

    spreadsheet = gc.open_by_key(created["id"])
    return spreadsheet, created["webViewLink"]


# ─── STEP 5 + 6 — WRITE HEADERS AND INSERT DATA ───────────────────────────────

def insert_rows(
    spreadsheet: gspread.Spreadsheet,
    rows: list[dict],
    mapping: dict[str, str | None],
) -> int:
    """
    1. Write STANDARD_HEADERS to row 1, freeze and bold them.
    2. Map each CSV row to the standard column order.
    3. Append all data rows in chunks (safe for large files).
    Returns total rows written.
    """
    ws = spreadsheet.sheet1
    ws.clear()

    # ── Headers ──────────────────────────────────────────────────────────────
    ws.update("A1", [STANDARD_HEADERS])
    ws.freeze(rows=1)
    last_col = _col_letter(len(STANDARD_HEADERS))
    ws.format(
        f"A1:{last_col}1",
        {"textFormat": {"bold": True}, "backgroundColor": {"red": 0.95, "green": 0.95, "blue": 0.95}},
    )

    # ── Data ─────────────────────────────────────────────────────────────────
    data: list[list[str]] = []
    for csv_row in rows:
        out_row: list[str] = []
        for col in STANDARD_HEADERS:
            if col in SYSTEM_COLUMNS:
                out_row.append("")          # filled later by the calling system
            else:
                src_col = mapping.get(col)
                val = csv_row.get(src_col, "") if src_col else ""
                out_row.append(str(val).strip())
        data.append(out_row)

    # Append in chunks to stay inside the 2 MB API payload limit
    _batch_append(ws, data)

    return len(data)


def _batch_append(ws: gspread.Worksheet, data: list[list], chunk_size: int = 500) -> None:
    """Append rows in chunks to avoid hitting API size limits."""
    for i in range(0, len(data), chunk_size):
        ws.append_rows(data[i : i + chunk_size], value_input_option="RAW")


# ─── MAIN ─────────────────────────────────────────────────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Import a CSV file into a standardized Lead Reactivation Google Sheet.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            '  python csv_importer.py --csv leads.csv --client "Phoenix Real Estate Group"\n'
            '  python csv_importer.py --csv leads.csv --client "Boise Med Spa" --creds /path/to/key.json\n'
        ),
    )
    parser.add_argument("--csv",    required=True,                   help="Path to the local CSV file")
    parser.add_argument("--client", required=True,                   help='Client name, e.g. "Phoenix Real Estate Group"')
    parser.add_argument("--folder", default=GOOGLE_DRIVE_FOLDER,     help="Google Drive folder ID (optional override)")
    parser.add_argument("--creds",  default=CREDENTIALS_FILE,        help="Path to service account key JSON (default: credentials.json)")
    args = parser.parse_args()

    _banner()

    # ── Step 1: Auth ─────────────────────────────────────────────────────────
    _step(1, f"Authenticating  [{args.creds}]")
    gc, drive = build_clients(args.creds)
    _ok()

    # ── Step 2: Read CSV ──────────────────────────────────────────────────────
    _step(2, f"Reading CSV     [{args.csv}]")
    headers, rows = read_csv(args.csv)
    _ok(f"{len(rows):,} data rows  •  {len(headers)} columns")
    print(f"         Columns detected: {', '.join(headers)}")

    # ── Step 3: Map fields ────────────────────────────────────────────────────
    _step(3, "Mapping fields")
    mapping = map_fields(headers)
    print()
    for std_field, src_col in mapping.items():
        arrow = f"← \"{src_col}\"" if src_col else "← (not found — will be empty)"
        print(f"         {std_field:<22}  {arrow}")

    # ── Step 4: Create sheet ──────────────────────────────────────────────────
    _step(4, f"Creating sheet  [Lead Reactivation - {args.client}]")
    spreadsheet, sheet_url = create_google_sheet(args.client, args.folder, gc, drive)
    _ok(f"id={spreadsheet.id}")

    # ── Step 5+6: Write + insert ──────────────────────────────────────────────
    _step(5, "Writing headers and inserting data")
    total = insert_rows(spreadsheet, rows, mapping)
    _ok(f"{total:,} rows written")

    # ── Final summary ─────────────────────────────────────────────────────────
    _divider()
    print("  ✅  IMPORT COMPLETE")
    _divider()
    print(f"  Sheet name  :  Lead Reactivation - {args.client}")
    print(f"  Rows imported: {total:,}")
    print()
    print("  Sheet URL:")
    print(f"  {sheet_url}")
    print()
    print("  Use this URL in your dashboard for this client.")
    _divider()


# ─── HELPERS ──────────────────────────────────────────────────────────────────

def _col_letter(n: int) -> str:
    """Convert 1-indexed column number to spreadsheet letter (1→A, 19→S, 27→AA)."""
    result = ""
    while n:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


def _banner() -> None:
    _divider()
    print("  Lead Reactivation — CSV Importer")
    _divider()


def _divider() -> None:
    print("\n" + "=" * 56)


def _step(n: int, label: str) -> None:
    print(f"\n[{n}] {label}  ", end="", flush=True)


def _ok(detail: str = "") -> None:
    suffix = f"  ({detail})" if detail else ""
    print(f"✓{suffix}")


def _abort(msg: str) -> None:
    print(f"\n  ERROR: {msg}\n")
    sys.exit(1)


# ─── ENTRY POINT ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    main()
