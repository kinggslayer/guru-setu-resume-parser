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
    extract_folder_id,
    trash_file,
)

from duplicates import find_duplicate_groups, summarise

from extractor import (
    read_pdf,
    read_docx,
    extract_resume_data,
    ocr_pdf,
    MIN_USABLE_CHARS
)

from portal_export import to_portal_dataframe, drop_same_teacher

from master_store import (
    extract_file_id,
    load_master,
    save_master,
    already_seen,
    to_master_rows,
    merge_into_master,
    strip_tracking,
    MASTER_FILE_NAME
)

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


# ---------------------------------------------------------------------------
#  Settings
#
#  The master link is remembered in session state, so it survives every
#  rerun while the app stays open. It is NOT permanent: Streamlit Community
#  Cloud gives the app no disk, so a reboot or a wake-from-sleep starts a
#  fresh session. Putting MASTER_FILE_LINK in Settings -> Secrets makes it
#  the permanent default, which this box pre-fills from.
# ---------------------------------------------------------------------------

if "master_link" not in st.session_state:
    st.session_state["master_link"] = get_secret("MASTER_FILE_LINK", "") or ""

with st.expander("Settings", expanded=not st.session_state["master_link"]):

    master_link = st.text_input(
        "Master database file",
        key="master_link",
        help=(
            "Link to the master .xlsx that stores every candidate ever "
            f"extracted. Leave blank and the app creates {MASTER_FILE_NAME} "
            "in the resumes folder on the first run."
        )
    )

    if master_link:
        resolved = extract_file_id(master_link)

        if resolved:
            st.caption(f"Using master file `{resolved}`")
        else:
            st.warning(
                "That doesn't look like a Drive link — paste the full URL "
                "from the browser address bar."
            )
    else:
        st.caption(
            f"No master file set. One will be created as {MASTER_FILE_NAME} "
            "in the resumes folder."
        )

    st.caption(
        "To make this permanent, add it to Settings -> Secrets as "
        'MASTER_FILE_LINK = "https://docs.google.com/..." — otherwise it '
        "resets when the app restarts."
    )

master_link = st.session_state["master_link"]

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

    # Hide the tracking columns: they're for dedup, not for editing, and
    # showing them makes the grid far harder to scan.
    df = strip_tracking(df)

    if REVIEW_COLUMN not in df.columns:
        df = add_review_column(df)

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

    # Edits go to their own key. Writing them back over state_key would
    # strip the master's tracking columns, and the next merge would lose
    # the dedup data.
    edit_key = f"{state_key}__edited"

    edited = st.data_editor(
        st.session_state.get(edit_key, df),
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

    st.session_state[edit_key] = edited

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
        "Edits here change the downloaded file only — they are not written "
        "back to the master database in Drive."
    )


def render_duplicate_cleanup():
    """
    Show byte-identical duplicate files found in the folder and offer to
    move the extra copies to Drive's trash.

    Rendered outside the Process button's body so it survives reruns, and
    gated behind an explicit confirmation because it changes the user's
    Drive rather than just this app's data.
    """

    groups = st.session_state.get("dup_groups")

    if not groups:
        return

    duplicate_count, group_count = summarise(groups)

    st.subheader("Duplicate files in this folder")

    st.write(
        f"{duplicate_count} duplicate file(s) across {group_count} set(s). "
        "These are byte-identical copies, matched on Drive's own checksum — "
        "nothing was downloaded to find them."
    )

    rows = []

    for group in groups:
        for duplicate in group["duplicates"]:
            rows.append({
                "Duplicate (will be trashed)": duplicate["name"],
                "Keeping": group["keep"]["name"],
            })

    st.dataframe(pd.DataFrame(rows), width="stretch", hide_index=True)

    st.caption(
        "The copy already recorded in the master database is always the one "
        "kept — trashing it instead would make the next run re-parse that "
        "resume and pay for it again."
    )

    confirm = st.checkbox(
        f"Yes, move these {duplicate_count} file(s) to my Drive trash",
        key="dup_confirm"
    )

    if st.button("Move duplicates to trash", disabled=not confirm):

        trashed = 0
        failures = []

        for group in groups:
            for duplicate in group["duplicates"]:

                try:
                    trash_file(duplicate["id"])
                    trashed += 1

                except Exception as e:
                    failures.append(f"{duplicate['name']}: {e}")

        if trashed:
            st.success(
                f"Moved {trashed} file(s) to the Drive trash. They stay "
                "recoverable there for 30 days if this was a mistake."
            )

        if failures:
            st.error(
                "Could not trash some files (the folder may be shared as "
                "Viewer rather than Editor):\n\n"
                + "\n\n".join(f"- {f}" for f in failures)
            )

        st.session_state.pop("dup_groups", None)


def render_bottom_panels():
    """
    Everything shown below the run output. Kept in one function because
    the app has an early st.stop() when a folder holds nothing new, and
    these panels must still appear on that path — trashing duplicates is
    precisely what a user wants after re-running an already-parsed folder.
    """

    render_duplicate_cleanup()

    render_review_and_download(
        state_key="master_df",
        file_name="guru_setu_teacher_import.xlsx",
        heading="Master database — review and download"
    )


def hash_text(text):
    normalized = " ".join(text.split()).lower()
    return hashlib.sha256(normalized.encode("utf-8")).hexdigest()

def process_resume(file_info, output_dir, existing_hashes):
    file_id = file_info["id"]
    file_name = file_info["name"]

    try:

        file_path = os.path.join(output_dir, file_name)
        ocr_used = False

        if file_name.lower().endswith(".pdf"):

            text = read_pdf(file_path)

            # Scanned or photographed CV: no text layer to read, so fall
            # back to transcribing the rendered pages. Everything after
            # this point treats the result like any other resume text.
            if len(text.strip()) < MIN_USABLE_CHARS:
                text = ocr_pdf(file_path)
                ocr_used = bool(text.strip())

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
        data["ocr_used"] = ocr_used
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

    try:
        folder_id = extract_folder_id(drive_link)

    except ValueError as e:
        st.error(f"{e}. Paste the folder URL from your browser's address bar.")
        st.stop()

    files, skipped_files = get_files_from_folder(drive_link)

    if skipped_files:
        skipped_names = ", ".join(f["name"] for f in skipped_files)
        st.warning(
            f"Skipped {len(skipped_files)} native Google Docs/Sheets/Slides "
            f"(not downloadable as-is): {skipped_names}"
        )

    # The master workbook lives in this folder; it is not a resume.
    files = [f for f in files if f["name"] != MASTER_FILE_NAME]

    if not files:
        st.error("No downloadable files found.")
        st.stop()

    st.write(f"Found {len(files)} files")

    # ------------------------------------------------------------------
    #  Load the master database so files parsed in ANY previous run are
    #  skipped before they cost an OpenAI call.
    # ------------------------------------------------------------------
    master_file_id = extract_file_id(master_link)

    if master_link and not master_file_id:
        st.error(
            "That master database link doesn't contain a file id. Paste the "
            "full URL from the browser address bar."
        )
        st.stop()

    try:
        master_df, master_existed = load_master(folder_id, master_file_id)

    except Exception as e:
        st.error(
            f"Could not read the master database: {e} "
            "Processing stopped rather than risk starting a second, "
            "competing master file."
        )
        st.stop()

    seen_file_ids, seen_hashes = already_seen(master_df)

    if master_existed:
        st.info(
            f"Master database has {len(master_df)} candidates from previous "
            f"runs. Files already in it will be skipped."
        )
    else:
        st.info(
            f"No master database yet — one will be created in this folder "
            f"as {MASTER_FILE_NAME}."
        )

    # Byte-identical copies, found from Drive's own checksums — free, and
    # done before parsing so the duplicate never costs a download or a call.
    dup_groups = find_duplicate_groups(files, seen_file_ids)

    st.session_state["dup_groups"] = dup_groups

    if dup_groups:
        duplicate_count, group_count = summarise(dup_groups)

        st.warning(
            f"{duplicate_count} duplicate file(s) in this folder across "
            f"{group_count} set(s) — identical copies under different names. "
            "They are skipped below, and you can trash them at the bottom "
            "of the page."
        )

        # Never parse a copy we're already parsing.
        duplicate_ids = {
            d["id"] for g in dup_groups for d in g["duplicates"]
        }

        files = [f for f in files if f["id"] not in duplicate_ids]

    # Skip already-known files up front: no download, no OpenAI call.
    known_files = [f for f in files if f["id"] in seen_file_ids]
    files = [f for f in files if f["id"] not in seen_file_ids]

    if known_files:
        st.write(
            f"Skipping {len(known_files)} file(s) already in the master "
            f"database. {len(files)} left to parse."
        )

    if not files:
        st.success("Nothing new in this folder — the master is up to date.")
        st.session_state["master_df"] = master_df
        st.session_state.pop("master_df__edited", None)
        render_bottom_panels()
        st.stop()

    existing_hashes = set(seen_hashes)
    existing_emails = set()
    existing_phones = set()

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

                    result["keep"] = True

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
        if r.get("keep")
    ]

    st.write(
        f"Done. {len(results)} files processed "
        f"({duplicate_content_count} duplicate content, "
        f"{skipped_existing_count} same person as another file, "
        f"{len(failed_extraction)} failed extraction). "
        f"{len(unique_results)} candidates in the sheet below."
    )

    ocr_used_count = sum(1 for r in results if r.get("ocr_used"))

    if ocr_used_count:
        st.info(
            f"{ocr_used_count} scanned resume(s) had no text layer and were "
            "read by transcribing the page images. Check those rows in the "
            "grid — transcription is less reliable than real text."
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

        # The reason is nearly always identical for every file (bad key, no
        # credit, model not enabled), so show the distinct reasons once
        # rather than repeating the same line per resume.
        reasons = sorted({
            r["extraction_error"]
            for r in failed_extraction
            if r.get("extraction_error")
        })

        if reasons:
            st.error(
                "OpenAI rejected the request. Reason"
                + ("s" if len(reasons) > 1 else "")
                + ":\n\n"
                + "\n\n".join(f"- {reason}" for reason in reasons)
            )

        st.error(
            f"⚠️ {len(failed_extraction)} resumes could not be extracted "
            f"(errored after all retries — only email/phone were regex-extracted "
            f"for these, everything else is blank). "
            f"Re-run later to retry them: {failed_names}"
        )

    if unique_results:

        # -----------------------------------------------------------------
        #  Merge this run into the master and write it back to Drive, then
        #  hand the whole master to session state.
        #
        #  Rendering happens outside this block: Streamlit re-runs the
        #  script on every interaction and this body only executes on the
        #  click, so a grid drawn here would vanish the moment the user
        #  typed in it.
        # -----------------------------------------------------------------
        new_rows = to_master_rows(unique_results)

        master_df, added = merge_into_master(master_df, new_rows)

        try:
            save_master(folder_id, master_df, master_file_id)

            where = (
                "your master database file"
                if master_file_id
                else f"{MASTER_FILE_NAME} in your Drive folder"
            )

            st.success(
                f"{added} new candidate(s) added. The master database now "
                f"holds {len(master_df)} candidates, saved to {where}."
            )

        except Exception as e:
            st.error(
                f"Could not save the master database: {e}\n\n"
                "This run's results are still shown below — download them "
                "now, because they are NOT saved. The usual cause is the "
                "folder being shared with the service account as Viewer "
                "rather than Editor."
            )

        st.session_state["master_df"] = master_df
        st.session_state.pop("master_df__edited", None)

    else:
        st.session_state.pop("master_df", None)
        st.session_state.pop("master_df__edited", None)

        if duplicate_content_count > 0:
            st.success(
                f"All {duplicate_content_count} resumes were already processed before."
            )
        else:
            st.warning("No candidate data extracted.")


# Rendered on EVERY run, not just the one where Process was clicked, so the
# grid and any edits made in it survive typing, sorting and the download
# click — Streamlit re-executes this script from the top each time.
render_bottom_panels()
