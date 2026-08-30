from sheets_utils import (
    append_candidate,
    get_existing_hashes,
    get_existing_contacts,
    ensure_headers,
    get_sheet
)
import streamlit as st
import os
import io
import shutil
import time
import hashlib

import pandas as pd

from concurrent.futures import ThreadPoolExecutor, as_completed

from drive_utils import (
    get_files_from_folder,
    download_file,
)

from extractor import (
    read_pdf,
    read_docx,
    extract_resume_data
)

from portal_export import to_portal_dataframe, drop_same_teacher

from review import (
    add_review_column,
    count_flagged,
    strip_review_column,
    REVIEW_COLUMN
)

from portal_vocab import (
    TEACHER_STATUSES,
    GENDERS,
    JOB_TYPES,
    AVAILABILITY,
    COLLEGE_TYPES
)

EXCEL_MIME = (
    "application/vnd.openxmlformats-officedocument"
    ".spreadsheetml.sheet"
)

st.set_page_config(
    page_title="Teacher Resume Parser",
    layout="wide"
)

st.title("Teacher Resume Parser")


def get_secret(name, default=None):
    """
    Works both locally (.env via os.environ, loaded by extractor.py's
    load_dotenv()) and on Streamlit Community Cloud (st.secrets, set in
    the app's Settings -> Secrets). Streamlit secrets take priority so a
    deployed app doesn't need a .env file at all.
    """
    try:
        if name in st.secrets:
            return st.secrets[name]
    except Exception:
        pass
    return os.getenv(name, default)


OPENAI_API_KEY = get_secret("OPENAI_API_KEY")


if not OPENAI_API_KEY:
    st.error(
        "OPENAI_API_KEY is not set. Add it to your .env file locally, or "
        "to this app's Secrets if it's deployed on Streamlit Cloud."
    )
    st.stop()
os.environ.setdefault("OPENAI_API_KEY", OPENAI_API_KEY)

drive_link = st.text_input(
    "Paste Google Drive Folder Link"
)

OPENAI_WORKERS = int(get_secret("OPENAI_WORKERS", 2))
DOWNLOAD_WORKERS = 5

def to_excel_bytes(df):
    """Build the .xlsx in memory — no temp file to go stale between reruns."""

    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)

    return buffer.getvalue()


def render_review_and_download(state_key, file_name, heading):
    """
    Show the editable review grid for whatever is in st.session_state[state_key]
    and offer the corrected file for download.

    The edited table is written straight back into session state, so a fix
    survives the next rerun (clicking download, expanding a section, etc.)
    instead of being thrown away.
    """

    df = st.session_state.get(state_key)

    if df is None or df.empty:
        return

    st.subheader(heading)

    flagged_before = count_flagged(df)

    if flagged_before:
        st.warning(
            f"{flagged_before} of {len(df)} rows are missing a name, phone or "
            "subject and have been sorted to the top. Fix them below — the "
            "download always uses what's currently in the grid."
        )
    else:
        st.success(f"All {len(df)} rows look complete.")

    edited = st.data_editor(
        df,
        width="stretch",
        num_rows="dynamic",
        hide_index=True,
        key=f"editor_{state_key}",
        column_config={
            REVIEW_COLUMN: st.column_config.TextColumn(
                REVIEW_COLUMN,
                help=(
                    "Filled in when the row was parsed. It does not refresh "
                    "as you type — the live count below does."
                ),
                disabled=True,
                width="medium"
            ),
            "Status": st.column_config.SelectboxColumn(
                "Status", options=TEACHER_STATUSES
            ),
            "Gender": st.column_config.SelectboxColumn(
                "Gender", options=[""] + GENDERS
            ),
            "Preferred Job Type": st.column_config.SelectboxColumn(
                "Preferred Job Type", options=[""] + JOB_TYPES
            ),
            "Availability": st.column_config.SelectboxColumn(
                "Availability", options=[""] + AVAILABILITY
            ),
            "Current Salary Unit": st.column_config.SelectboxColumn(
                "Current Salary Unit", options=["", "pm", "lpa"]
            ),
            "Expected Salary Unit": st.column_config.SelectboxColumn(
                "Expected Salary Unit", options=["", "pm", "lpa"]
            ),
            "College Type": st.column_config.SelectboxColumn(
                "College Type", options=[""] + COLLEGE_TYPES
            ),
            "CV URL": st.column_config.LinkColumn("CV URL"),
        }
    )

    st.session_state[state_key] = edited

    flagged_after = count_flagged(edited)

    if flagged_after:
        st.caption(
            f"{flagged_after} row(s) still incomplete. You can download "
            "anyway — the portal skips rows with no name, email and phone."
        )
    else:
        st.caption("Nothing outstanding.")

    # The review column is dropped so the file matches the import template
    # exactly, with no column the portal doesn't recognise.
    download_df = strip_review_column(edited)

    st.download_button(
        label=f"Download portal import file ({len(download_df)} teachers)",
        data=to_excel_bytes(download_df),
        file_name=file_name,
        mime=EXCEL_MIME,
        type="primary",
        key=f"download_{state_key}"
    )

    st.caption(
        "Upload it in the portal under Import / Export -> Import teachers. "
        "Note: edits made here are NOT written back to the master sheet, so "
        "the next run still de-duplicates against the original values."
    )


def hash_text(text):
    normalized = " ".join(text.split()).lower()
    return hashlib.sha256(normalized.encode("utf-8")).hexdigest()

def process_resume(file_info, output_dir, existing_hashes):
    file_id = file_info["id"]
    file_name = file_info["name"]

    try:

        file_path = os.path.join(output_dir, file_name)

        if file_name.lower().endswith(".pdf"):
            text = read_pdf(file_path)
        elif file_name.lower().endswith(".docx"):
            text = read_docx(file_path)
        else:
            return None

        content_hash = hash_text(text)

        if content_hash in existing_hashes:
            return {
                "duplicate_of_content": True,
                "resume_file_name": file_name,
                "content_hash": content_hash
            }

        data = extract_resume_data(text)
        data["content_hash"] = content_hash
        data["file_id"] = file_id
        data["resume_link"] = f"https://drive.google.com/file/d/{file_id}/view"
        data["resume_file_name"] = file_name
        data["duplicate_of_content"] = False
        return data

    except Exception as e:

        return {"error": f"{file_name}: {str(e)}"}


if st.button("Process Resumes"):

    if not drive_link:
        st.error("Please enter a Google Drive folder link.")
        st.stop()

    output_dir = "temp_resumes"

    if os.path.exists(output_dir):
        shutil.rmtree(output_dir)

    os.makedirs(output_dir)

    st.write("Fetching files from Google Drive...")

    files, skipped_files = get_files_from_folder(drive_link)

    if skipped_files:
        skipped_names = ", ".join(f["name"] for f in skipped_files)
        st.warning(
            f"Skipped {len(skipped_files)} native Google Docs/Sheets/Slides "
            f"(not downloadable as-is): {skipped_names}"
        )

    if not files:
        st.error("No downloadable files found.")
        st.stop()

    st.write(f"Found {len(files)} files")

    try:
        added_headers = ensure_headers()

        if added_headers:
            st.success(
                "Added new columns to the master sheet: "
                + ", ".join(added_headers)
            )

    except Exception as e:
        st.error(
            f"Master sheet headers need attention: {e} "
            "Processing stopped — fix row 1 first, or new fields will be "
            "written under blank headers and lost."
        )
        st.stop()

    existing_hashes = get_existing_hashes()
    existing_emails, existing_phones = get_existing_contacts()

    st.info(
        f"{len(existing_hashes)} resumes already in the master sheet by content — "
        f"{len(files)} files in this folder will be checked against them now "
        f"(downloads happen regardless; duplicates are skipped before any OpenAI call)."
    )

    download_progress = st.progress(0)
    process_progress = st.progress(0)

    st.write("Downloading + processing resumes...")

    results = []
    download_failures = []
    total = len(files)
    downloaded_count = 0
    processed_count = 0

    download_executor = ThreadPoolExecutor(max_workers=DOWNLOAD_WORKERS)
    groq_executor = ThreadPoolExecutor(max_workers=OPENAI_WORKERS)

    download_futures = {
        download_executor.submit(
            download_file, f["id"], f["name"], output_dir
        ): f
        for f in files
    } if files else {}

    process_futures = {}

    for future in as_completed(download_futures):

        file_info = download_futures[future]
        downloaded_count += 1
        if total:
            download_progress.progress(downloaded_count / total)

        try:
            future.result()
            process_future = groq_executor.submit(
                process_resume,
                file_info,
                output_dir,
                existing_hashes
            )
            process_futures[process_future] = file_info

        except Exception as e:
            download_failures.append({"name": file_info["name"], "error": str(e)})
            st.error(f"Download failed: {file_info['name']} ({str(e)})")

    for future in as_completed(process_futures):

        file_info = process_futures[future]
        processed_count += 1
        if process_futures:
            process_progress.progress(processed_count / len(process_futures))

        result = future.result()

        if result:
            if "error" in result:
                st.error(result["error"])
            else:

                if (
                    not result.get("extraction_failed")
                    and not result.get("duplicate_of_content")
                ):

                    email = str(result.get("email", "")).strip().lower()
                    phone = str(result.get("phone", "")).strip()

                    if (email and email in existing_emails) or (phone and phone in existing_phones):

                        st.warning(
                            f"Skipping existing candidate: {result.get('full_name')}"
                        )

                        result["skipped_existing_contact"] = True
                        results.append(result)
                        continue

                    time.sleep(1)
                    append_candidate(result)
                    result["written_to_sheet"] = True

                    if result.get("content_hash"):
                        existing_hashes.add(str(result["content_hash"]))

                    if email:
                        existing_emails.add(email)
                    if phone:
                        existing_phones.add(phone)

                results.append(result)

    download_executor.shutdown(wait=True)
    groq_executor.shutdown(wait=True)

    duplicate_content_count = sum(1 for r in results if r.get("duplicate_of_content"))
    skipped_existing_count = sum(1 for r in results if r.get("skipped_existing_contact"))

    failed_extraction = [r for r in results if r.get("extraction_failed")]

    unique_results = [
        r for r in results
        if r.get("written_to_sheet")
    ]

    st.write(
        f"Done. {len(results)} files processed "
        f"({duplicate_content_count} duplicate content, "
        f"{skipped_existing_count} matched an existing candidate, "
        f"{len(failed_extraction)} failed extraction and were NOT written to the sheet). "
        f"{len(unique_results)} new candidates actually added to the sheet this run."
    )

    needs_ocr = [r for r in results if r.get("needs_ocr")]

    if needs_ocr:

        ocr_names = ", ".join(
            r.get("resume_file_name", "unknown") for r in needs_ocr
        )

        st.warning(
            f"{len(needs_ocr)} file(s) had almost no readable text — they are "
            f"scans or images, so no OpenAI call was made for them. Convert "
            f"them to text (Google Docs OCR works) and re-run: {ocr_names}"
        )

    failed_extraction = [r for r in failed_extraction if not r.get("needs_ocr")]

    if failed_extraction:

        failed_names = ", ".join(
            r.get("resume_file_name", "unknown") for r in failed_extraction
        )

        st.error(
            f"⚠️ {len(failed_extraction)} resumes could not be extracted "
            f"(errored after all retries — only email/phone were regex-extracted "
            f"for these, everything else is blank). "
            f"Re-run later to retry them: {failed_names}"
        )

    if unique_results:

        # -----------------------------------------------------------------
        #  Hand the result to session state rather than rendering it here.
        #  Streamlit re-runs the whole script on every interaction, and this
        #  block only executes on the run where the button was clicked — so
        #  anything drawn inside it would vanish the moment the user typed
        #  in the review grid. The grid is rendered outside, from state.
        # -----------------------------------------------------------------
        portal_df = to_portal_dataframe(unique_results)
        portal_df = drop_same_teacher(portal_df)

        st.session_state["run_df"] = add_review_column(portal_df)

    else:
        st.session_state.pop("run_df", None)

        if duplicate_content_count > 0:
            st.success(
                f"All {duplicate_content_count} resumes were already processed before."
            )
        else:
            st.warning("No candidate data extracted.")


# Rendered on EVERY run, not just the one where Process was clicked, so the
# grid and any edits made in it survive typing, sorting and the download
# click — Streamlit re-executes this script from the top each time.
render_review_and_download(
    state_key="run_df",
    file_name="guru_setu_teacher_import.xlsx",
    heading="Review and edit before uploading"
)


# =============================================================================
#  Export the whole master sheet in portal format
#  Useful for a first-time bulk load, or to re-import everything after the
#  mapping changes. Independent of the run above — reads the Google Sheet.
# =============================================================================

st.divider()

with st.expander("Export the entire master sheet for the portal"):

    st.caption(
        "Converts every candidate already in Resume_Master_DB into the "
        "portal's teacher import format. The portal skips duplicates on "
        "import, so re-uploading an overlapping file is safe."
    )

    if st.button("Build portal file from master sheet"):

        try:
            sheet = get_sheet()
            records = sheet.get_all_records()

        except Exception as e:
            st.error(f"Could not read the master sheet: {e}")
            records = None

        if records is not None:

            if not records:
                st.warning("The master sheet is empty.")

            else:

                full_df = to_portal_dataframe(records)
                full_df = drop_same_teacher(full_df)

                st.session_state["master_df"] = add_review_column(full_df)

                st.success(
                    f"{len(full_df)} teachers ready "
                    f"(from {len(records)} sheet rows)."
                )

    # Inside the expander but OUTSIDE the button block, for the same reason
    # as the run grid above: the button's body only runs once.
    render_review_and_download(
        state_key="master_df",
        file_name="guru_setu_teacher_import_all.xlsx",
        heading="Review and edit the full master export"
    )
