"""
The master database: every candidate ever extracted, in one workbook.

WHERE IT LIVES
--------------
In the same Google Drive folder as the resumes, as `_resume_master_database.xlsx`.

It has to live outside the app because Streamlit Community Cloud gives the
container no persistent disk — the filesystem is wiped on every reboot,
redeploy and wake-from-sleep. Drive is already authenticated for reading the
resumes, so it needs no new credential; it does need the folder shared with
the service account as **Editor** rather than Viewer.

WHAT'S IN IT
------------
The 29 portal import columns, plus four tracking columns:

    Source File ID    - the Drive file id, so the same file is never re-parsed
    Content Hash      - hash of the extracted text, so the same CV uploaded
                        under a different name is also caught
    Source File Name  - which file a row came from, for tracing a bad row back
    Extracted At      - when it was added

The tracking columns are stripped from the portal download, so the file you
upload to the portal stays exactly the 29-column template.
"""

import io
import re
from datetime import datetime, timezone

import pandas as pd

from portal_export import PORTAL_COLUMNS, to_portal_dataframe
from drive_utils import (
    download_bytes_from_drive,
    upload_bytes_to_drive,
    download_bytes_by_id,
    upload_bytes_by_id,
    get_file_metadata,
)

GOOGLE_SHEET_MIME = "application/vnd.google-apps.spreadsheet"


def extract_file_id(value):
    """
    Accept either a bare Drive file id or a full spreadsheet/file URL.

    Pasting the URL straight from the browser is the obvious thing to do,
    so it shouldn't be an error.
    """

    text = str(value or "").strip()

    if not text:
        return None

    match = re.search(r"/d/([a-zA-Z0-9_-]+)", text)

    if match:
        return match.group(1)

    # Already an id.
    if re.fullmatch(r"[a-zA-Z0-9_-]{20,}", text):
        return text

    return None


# ---------------------------------------------------------------------------
#  Legacy column names
#
#  An existing master written by the earlier Sheets-based pipeline uses
#  snake_case headers. Mapping them across means that history is carried
#  forward and, more importantly, that its content hashes still count for
#  deduplication instead of every old resume being parsed again.
# ---------------------------------------------------------------------------

LEGACY_COLUMN_MAP = {
    "full_name": "Full Name",
    "email": "Email",
    "phone": "Phone",
    "gender": "Gender",
    "age": "Age",
    "city": "City",
    "state": "State",
    "subjects": "Subjects",
    "grade_levels": "Grade Levels",
    "languages": "Languages",
    "skills": "Skills",
    "qualification": "Qualification",
    "extra_qualifications": "Extra Qualifications",
    "college_type": "College Type",
    "college": "College",
    "experience_years": "Experience (years)",
    "current_institution": "Current Institution",
    "preferred_job_type": "Preferred Job Type",
    "availability": "Availability",
    "current_salary": "Current Salary",
    "current_salary_unit": "Current Salary Unit",
    "expected_salary": "Expected Salary",
    "expected_salary_unit": "Expected Salary Unit",
    "tags": "Tags",
    "resume_link": "CV URL",
    "file_id": "Source File ID",
    "content_hash": "Content Hash",
    "resume_file_name": "Source File Name",
}


def apply_legacy_names(df):
    """Rename any legacy snake_case headers onto the current ones."""

    renames = {
        column: LEGACY_COLUMN_MAP[column]
        for column in df.columns
        if column in LEGACY_COLUMN_MAP
        and LEGACY_COLUMN_MAP[column] not in df.columns
    }

    return df.rename(columns=renames) if renames else df


MASTER_FILE_NAME = "_resume_master_database.xlsx"

EXCEL_MIME = (
    "application/vnd.openxmlformats-officedocument"
    ".spreadsheetml.sheet"
)

SOURCE_FILE_ID = "Source File ID"
CONTENT_HASH = "Content Hash"
SOURCE_FILE_NAME = "Source File Name"
EXTRACTED_AT = "Extracted At"

TRACKING_COLUMNS = [
    SOURCE_FILE_ID,
    CONTENT_HASH,
    SOURCE_FILE_NAME,
    EXTRACTED_AT,
]

MASTER_COLUMNS = PORTAL_COLUMNS + TRACKING_COLUMNS


def empty_master():
    """A master with the right shape but no rows."""

    return pd.DataFrame(columns=MASTER_COLUMNS)


def check_master_file(file_id):
    """
    Confirm a pinned master file can actually be used, with a specific
    message when it can't. A native Google Sheet cannot be read or written
    as a workbook, and that failure is otherwise cryptic.
    """

    meta = get_file_metadata(file_id)

    if meta.get("mimeType") == GOOGLE_SHEET_MIME:
        raise ValueError(
            f"'{meta.get('name')}' is a native Google Sheet, which can't be "
            "used as the master workbook. In Google Sheets open it and "
            "choose File -> Download -> Microsoft Excel (.xlsx), then upload "
            "that .xlsx to Drive and point the app at the uploaded file."
        )

    return meta


def load_master(folder_id, file_id=None):
    """
    Read the master workbook from Drive.

    file_id pins a specific existing workbook; without it the workbook is
    found by name inside the resumes folder (and created there on first run).

    Returns (dataframe, existed). `existed` is False on the very first run,
    which lets the caller say "starting a new database" rather than
    reporting an empty one as if something went wrong.

    A master written by an older version may use legacy headers or be
    missing newer columns; both are repaired here rather than left to blow
    up downstream.
    """

    if file_id:
        check_master_file(file_id)
        data = download_bytes_by_id(file_id)

    else:
        data = download_bytes_from_drive(folder_id, MASTER_FILE_NAME)

    if data is None:
        return empty_master(), False

    df = pd.read_excel(io.BytesIO(data), dtype=str)

    df = apply_legacy_names(df)

    if df.empty and not len(df.columns):
        return empty_master(), True

    for column in MASTER_COLUMNS:
        if column not in df.columns:
            df[column] = ""

    df = df.fillna("")

    # Keep any extra columns the user added by hand in Excel; they're
    # harmless and losing someone's manual notes would be worse.
    extra = [c for c in df.columns if c not in MASTER_COLUMNS]

    return df[MASTER_COLUMNS + extra], True


def save_master(folder_id, df, file_id=None):
    """Write the master back to Drive, replacing the previous version."""

    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)

    if file_id:
        upload_bytes_by_id(file_id, buffer.getvalue(), EXCEL_MIME)
        return file_id

    return upload_bytes_to_drive(
        folder_id,
        MASTER_FILE_NAME,
        buffer.getvalue(),
        EXCEL_MIME
    )


def already_seen(df):
    """
    The Drive file ids and content hashes already in the master.

    Two separate checks on purpose: the file id catches re-running the same
    folder, and the content hash catches the same CV re-uploaded under a new
    name (which is common — "resume (1).pdf", "resume final.pdf").
    """

    if df.empty:
        return set(), set()

    def values(column):
        if column not in df.columns:
            return set()
        return {
            str(v).strip()
            for v in df[column]
            if str(v).strip() and str(v).strip().lower() != "nan"
        }

    return values(SOURCE_FILE_ID), values(CONTENT_HASH)


def to_master_rows(records):
    """Parsed records -> master rows (portal columns plus tracking)."""

    portal_df = to_portal_dataframe(records)

    if portal_df.empty:
        return empty_master()

    # to_portal_dataframe drops unusable rows, so line the tracking data up
    # by position on the records that actually survived rather than
    # assuming the two lists still match index for index.
    usable = [
        r for r in records
        if any([
            str(r.get("full_name") or "").strip(),
            str(r.get("email") or "").strip(),
            str(r.get("phone") or "").strip(),
        ])
    ]

    stamp = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M")

    portal_df[SOURCE_FILE_ID] = [
        str(r.get("file_id") or "") for r in usable
    ]
    portal_df[CONTENT_HASH] = [
        str(r.get("content_hash") or "") for r in usable
    ]
    portal_df[SOURCE_FILE_NAME] = [
        str(r.get("resume_file_name") or "") for r in usable
    ]
    portal_df[EXTRACTED_AT] = stamp

    return portal_df


def merge_into_master(master_df, new_df):
    """
    Append new rows, then drop any row whose file id or content hash already
    appears earlier. Existing rows always win, so a manual correction made
    in the master workbook is never overwritten by a re-parse.
    """

    if new_df.empty:
        return master_df, 0

    combined = pd.concat([master_df, new_df], ignore_index=True)

    before = len(combined)

    for column in (SOURCE_FILE_ID, CONTENT_HASH):

        # Blank values must not collapse together — pandas treats them as
        # equal, which would silently delete distinct candidates whose
        # tracking data is missing.
        has_value = combined[column].astype(str).str.strip() != ""

        keyed = combined[has_value]
        unkeyed = combined[~has_value]

        keyed = keyed.drop_duplicates(subset=[column], keep="first")

        combined = pd.concat([keyed, unkeyed], ignore_index=True)

    added = len(combined) - len(master_df)

    return combined, max(added, 0)


def strip_tracking(df):
    """Master -> exactly the portal's 29 columns, for the upload file."""

    return df[[c for c in PORTAL_COLUMNS if c in df.columns]]
