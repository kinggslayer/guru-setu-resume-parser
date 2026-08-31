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

from portal_export import (
    PORTAL_COLUMNS,
    to_portal_dataframe,
    drop_same_teacher,
)
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


def cell(row, column):
    """
    A cell's text, with missing values normalised to "".

    Values read back from a workbook can be the literal strings "None" or
    "nan" if an earlier version wrote them that way. Those must count as
    empty: any non-empty stand-in makes two rows that are both missing the
    field look identical, which silently merges unrelated candidates.
    """

    value = row.get(column, "")

    if value is None:
        return ""

    text = str(value).strip()

    if text.lower() in ("none", "nan", "null"):
        return ""

    return text


def drop_blank_rows(df):
    """
    Remove rows where every cell is empty. A row counts as real if it has
    any identifying value at all — name, email or phone — so a partially
    parsed candidate is still kept for review.
    """

    if df.empty:
        return df

    # Name, email or phone ONLY — deliberately NOT Source File ID.
    #
    # A row whose extraction failed still carries a file id, so counting
    # that as content keeps rows with no usable data at all. The portal
    # rejects those anyway (isImportRowUsable), so matching its rule here
    # keeps the master and the portal in agreement.
    identifying = [
        column
        for column in ("Full Name", "Email", "Phone")
        if column in df.columns
    ]

    if not identifying:
        return df

    def has_content(row):
        return any(cell(row, column) for column in identifying)

    keep = [index for index, row in df.iterrows() if has_content(row)]

    return df.loc[keep]


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

    # Drop rows that are entirely blank.
    #
    # Clearing cell contents in Excel or Sheets (select + Delete) leaves the
    # rows in place, so the workbook still reports its old length. Without
    # this, emptying the master by hand appears to do nothing: the app keeps
    # reporting the original count and keeps skipping files.
    df = drop_blank_rows(df)

    # Keep any extra columns the user added by hand in Excel; they're
    # harmless and losing someone's manual notes would be worse.
    extra = [c for c in df.columns if c not in MASTER_COLUMNS]

    return df[MASTER_COLUMNS + extra].reset_index(drop=True), True


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


def _mobile(value):
    """Last 10 digits, matching the portal's localMobile()."""

    digits = "".join(c for c in str(value or "") if c.isdigit())

    return digits[-10:] if len(digits) >= 10 else ""


def merge_into_master(master_df, new_df):
    """
    Append new rows, dropping duplicates on three levels. Existing rows
    always win, so a manual correction in the master workbook is never
    overwritten by a re-parse.

    Each level catches something the others miss:

      Source File ID  - the same Drive file parsed again
      Content Hash    - the same CV re-uploaded under a different name
      name + phone    - the SAME PERSON from two files that are neither
                        byte-identical nor textually identical (a CV
                        re-exported from Word, or a scan whose OCR came out
                        differently). Uses the portal's isSameTeacher rule
                        so the parser and portal agree.

    Returns (combined, added, report). The report explains every dropped
    row — without it a shrinking candidate count looks like data loss.
    """

    if new_df.empty:
        return master_df, 0, []

    columns = list(master_df.columns) or list(new_df.columns)

    seen_ids = {}
    seen_hashes = {}
    seen_people = {}

    kept = []
    report = []

    def label(row):
        name = cell(row, "Full Name")
        phone = cell(row, "Phone")
        source = cell(row, SOURCE_FILE_NAME)
        return f"{name or '(no name)'} {phone}".strip() + (
            f" [{source}]" if source else ""
        )

    for origin, frame in (("master", master_df), ("new", new_df)):

        for _, row in frame.iterrows():

            file_id = cell(row, SOURCE_FILE_ID)
            content = cell(row, CONTENT_HASH)

            name = cell(row, "Full Name").lower()
            mobile = _mobile(cell(row, "Phone"))

            person = (name, mobile) if name and mobile else None

            # Blank keys must never match each other, or unrelated rows
            # with missing data would collapse into one.
            if file_id and file_id in seen_ids:
                if origin == "new":
                    report.append({
                        "dropped": label(row),
                        "reason": "same Drive file already recorded",
                        "matched": seen_ids[file_id],
                    })
                continue

            if content and content in seen_hashes:
                if origin == "new":
                    report.append({
                        "dropped": label(row),
                        "reason": "identical text to another resume",
                        "matched": seen_hashes[content],
                    })
                continue

            if person and person in seen_people:
                if origin == "new":
                    report.append({
                        "dropped": label(row),
                        "reason": "same name and phone",
                        "matched": seen_people[person],
                    })
                continue

            if file_id:
                seen_ids[file_id] = label(row)
            if content:
                seen_hashes[content] = label(row)
            if person:
                seen_people[person] = label(row)

            kept.append(row)

    combined = pd.DataFrame(kept, columns=columns) if kept else empty_master()

    combined = drop_blank_rows(combined).reset_index(drop=True)

    added = max(len(combined) - len(master_df), 0)

    return combined, added, report


def strip_tracking(df):
    """Master -> exactly the portal's 29 columns, for the upload file."""

    return df[[c for c in PORTAL_COLUMNS if c in df.columns]]
