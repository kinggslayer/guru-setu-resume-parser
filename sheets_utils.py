import os
import gspread
import streamlit as st

from google.oauth2 import service_account

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

SHEET_NAME = "Resume_Master_DB"

# The exact header row append_candidate() writes against, in order. Paste
# this as row 1 of Resume_Master_DB. gspread keys every row off this row,
# so a missing header means the value is silently unreadable on the way
# back out (and blank headers make get_all_records() raise).
SHEET_HEADERS = [
    "file_id", "content_hash", "resume_file_name",
    "full_name", "email", "phone", "gender", "age", "city",
    "subjects", "grade_levels", "languages",
    "qualification", "extra_qualifications",
    "college_type", "college", "education_history",
    "experience_years", "current_institution", "current_designation",
    "previous_institutions", "skills", "resume_link",
    # appended in the portal-format change
    "state", "preferred_job_type", "availability",
    "current_salary", "current_salary_unit",
    "expected_salary", "expected_salary_unit", "tags",
]


def ensure_headers():
    """
    Make row 1 match SHEET_HEADERS. Only ever ADDS the missing trailing
    headers — existing header cells are left untouched, so a sheet that has
    been renamed or reordered by hand is never clobbered. Returns the list
    of headers that were added.
    """

    sheet = get_sheet()

    current = sheet.row_values(1)

    if current == SHEET_HEADERS:
        return []

    # Only extend when the existing headers are a prefix of the expected
    # ones; anything else means the sheet was customised and needs a human.
    if current and current != SHEET_HEADERS[:len(current)]:
        raise ValueError(
            "Row 1 of Resume_Master_DB does not match the expected headers. "
            "Fix it by hand rather than letting this overwrite your columns. "
            f"Expected to start with: {SHEET_HEADERS[:len(current)]}"
        )

    missing = SHEET_HEADERS[len(current):]

    sheet.update(
        range_name="A1",
        values=[SHEET_HEADERS]
    )

    return missing


def _get_credentials():
    """
    Works both locally (a JSON key file on disk) and on Streamlit
    Community Cloud (the JSON content pasted into Secrets — no file
    needed). Streamlit secrets take priority when present.
    """

    try:
        if "gcp_service_account" in st.secrets:
            return service_account.Credentials.from_service_account_info(
                dict(st.secrets["gcp_service_account"]),
                scopes=SCOPES
            )
    except Exception:
        pass

    service_account_file = os.getenv(
        "GOOGLE_SERVICE_ACCOUNT_FILE",
        "symbolic-axe-502107-r1-2ff6db019a1f.json"
    )

    return service_account.Credentials.from_service_account_file(
        service_account_file,
        scopes=SCOPES
    )

_sheet_cache = {"sheet": None}


def get_sheet(force_refresh=False):

    if _sheet_cache["sheet"] is not None and not force_refresh:
        return _sheet_cache["sheet"]

    creds = _get_credentials()

    client = gspread.authorize(creds)

    sheet = client.open(SHEET_NAME).sheet1

    _sheet_cache["sheet"] = sheet

    return sheet


def append_candidate(candidate):
    """
    College, education history and previous institutions are joined with
    "; " rather than ", " because their individual entries contain commas
    ("B.Sc Physics, Fergusson College, 2012" is ONE entry). A comma join
    makes the cell impossible to split back apart when exporting to the
    portal; a semicolon join round-trips cleanly.
    """

    sheet = get_sheet()

    row = [
        candidate.get("file_id"),
        candidate.get("content_hash"),
        candidate.get("resume_file_name"),
        candidate.get("full_name"),
        candidate.get("email"),
        candidate.get("phone"),
        candidate.get("gender"),
        candidate.get("age"),
        candidate.get("city"),
        ", ".join(candidate.get("subjects") or []),
        ", ".join(candidate.get("grade_levels") or []),
        ", ".join(candidate.get("languages") or []),
        candidate.get("qualification"),
        ", ".join(candidate.get("extra_qualifications") or []),
        candidate.get("college_type"),
        "; ".join(candidate.get("college") or []),
        "; ".join(candidate.get("education_history") or []),
        candidate.get("experience_years"),
        candidate.get("current_institution"),
        candidate.get("current_designation"),
        "; ".join(candidate.get("previous_institutions") or []),
        ", ".join(candidate.get("skills") or []),
        candidate.get("resume_link"),

        # --- appended columns -----------------------------------------
        # New fields go on the RIGHT-HAND END so every existing column
        # keeps its position and rows written before this change still
        # line up. Add the matching headers to Resume_Master_DB (see
        # SHEET_HEADERS below) or get_all_records() will fail on the
        # blank header cells.
        candidate.get("state"),
        candidate.get("preferred_job_type"),
        candidate.get("availability"),
        candidate.get("current_salary"),
        candidate.get("current_salary_unit"),
        candidate.get("expected_salary"),
        candidate.get("expected_salary_unit"),
        "; ".join(candidate.get("tags") or [])
    ]

    sheet.append_row(row)


def get_existing_hashes():

    sheet = get_sheet()

    records = sheet.get_all_records()

    hashes = set()

    for row in records:

        h = row.get("content_hash")

        if h:
            hashes.add(str(h))

    return hashes


def get_existing_contacts():

    sheet = get_sheet()

    records = sheet.get_all_records()

    emails = set()
    phones = set()

    for row in records:

        if row.get("email"):
            emails.add(str(row["email"]).strip().lower())

        if row.get("phone"):
            phones.add(str(row["phone"]).strip())

    return emails, phones