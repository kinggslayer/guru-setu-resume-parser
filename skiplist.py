"""
Remember which files were tried and produced nothing usable.

Without this, every run re-parses the same broken PDFs and unreadable
scans — and each attempt is a paid OpenAI call, OCR included. On a folder
of a few thousand CVs that is a recurring charge for zero result.

The list lives beside the resumes in Drive as a small JSON file, for the
same reason the master does: Streamlit Community Cloud gives the container
no persistent disk.

A failure is not necessarily permanent — a rate limit or a dropped
connection could cause one — so the list is advisory. The "Retry files that
failed before" checkbox ignores it for one run.
"""

import json

from drive_utils import download_cache_from_drive, upload_cache_to_drive


FAILED_FILE_NAME = "_parser_skipped_files.json"


def load_failed(folder_id):
    """
    {file_id: reason} for everything previously attempted without success.

    Returns an empty dict on any problem: not being able to read this list
    should cost a little money, never stop a run.
    """

    if not folder_id:
        return {}

    try:
        raw = download_cache_from_drive(folder_id, FAILED_FILE_NAME)

    except Exception:
        return {}

    if not raw:
        return {}

    try:
        data = json.loads(raw)

    except (ValueError, TypeError):
        return {}

    if not isinstance(data, dict):
        return {}

    return {str(k): str(v) for k, v in data.items()}


def save_failed(folder_id, failed):
    """Write the list back. Silent on failure, for the same reason."""

    if not folder_id or not failed:
        return False

    try:
        upload_cache_to_drive(
            folder_id,
            FAILED_FILE_NAME,
            json.dumps(failed, indent=1, sort_keys=True)
        )

        return True

    except Exception:
        return False


def reason_for(result):
    """
    Why a parsed file produced no candidate, or None if it did.

    Only outcomes that will repeat are recorded. A duplicate is not a
    failure, and a file skipped as the same person is genuinely handled —
    recording either would wrongly suppress it forever.
    """

    if not result:
        return None

    if result.get("keep"):
        return None

    # These two DO get recorded, unlike a "keep".
    #
    # A file whose person is already in the master, or whose text matches
    # another resume, can never produce a new candidate — so downloading
    # and reading it again on every future run is pure waste. Recorded
    # separately from real failures so "Retry files that failed before"
    # brings them back if a master row is ever deleted.
    if result.get("already_in_master"):
        return "already in the master (same person)"

    if result.get("duplicate_of_content"):
        return "identical text to another resume"

    # A repeat WITHIN one run is not recorded: which copy won is an
    # accident of ordering, and suppressing the loser would make the
    # result depend on the order files happened to be processed in.
    if result.get("skipped_existing_contact"):
        return None

    if result.get("needs_ocr"):
        return "no readable text (scan or image-only)"

    if result.get("extraction_failed"):
        return "extraction failed"

    if result.get("error"):
        return str(result["error"])[:200]

    return "produced no usable candidate"
