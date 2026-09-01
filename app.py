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
    has_user_credentials,
)

from duplicates import find_duplicate_groups, summarise

from usage import tracker

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
    apply_edits,
    SOURCE_FILE_ID,
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
    review_issues,
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

drive_link = st.text_area(
    "Google Drive folder link(s)",
    height=80,
    help="One per line to process several folders in one run."
)

include_subfolders = st.checkbox(
    "Include sub-folders",
    value=True,
    help="Walks folders inside the one you paste."
)

uploaded_files = st.file_uploader(
    "...or upload CVs directly",
    type=["pdf", "docx"],
    accept_multiple_files=True,
    help=(
        "For resumes that arrived by WhatsApp or email. These are parsed "
        "alongside any Drive folders above."
    )
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


def _whoami(owner_email=None):
    """
    Which Google account the app would act as, for error messages.

    Imported lazily rather than at module load: it is a diagnostic nicety,
    and a stale drive_utils on the server should degrade this one message
    rather than stop the whole app from starting.
    """

    try:
        from drive_utils import whoami

        return whoami(owner_email)

    except Exception:
        return None


def folder_id_for_save():
    """
    The folder the master lives in, for saving.

    Recomputed from the inputs rather than passed down: the review panels
    render on every rerun, including ones where the Process button never
    ran, so the button block's local variables aren't available here.
    """

    # drive_link may now hold several lines; the master goes in the first.
    lines = [line.strip() for line in str(drive_link).splitlines() if line.strip()]

    if not lines:
        return None

    try:
        return extract_folder_id(lines[0])

    except Exception:
        return None


def master_file_id_for_save():
    """The pinned master file id, if one is configured."""

    return extract_file_id(master_link)


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
    # showing them makes the grid far harder to scan. The file id is kept
    # (hidden in the editor) so edits can be matched back to master rows.
    df = strip_tracking(df, keep_id=True)

    if REVIEW_COLUMN not in df.columns:
        df = add_review_column(df)

    st.subheader(heading)

    # ------------------------------------------------------------------
    #  Filter and search.
    #
    #  A master of a few thousand rows is unusable without these: finding
    #  one teacher, or working through only the incomplete rows, otherwise
    #  means scrolling.
    # ------------------------------------------------------------------
    total_rows = len(df)

    controls = st.columns([2, 1])

    with controls[0]:
        query = st.text_input(
            "Search",
            key=f"search_{state_key}",
            placeholder="Name, phone, email, city, subject..."
        )

    with controls[1]:
        only_flagged = st.checkbox(
            "Only rows needing review",
            key=f"flagged_{state_key}"
        )

    if only_flagged:
        df = df[[bool(review_issues(row)) for _, row in df.iterrows()]]

    if query:
        needle = query.strip().lower()

        # Search across every column: the useful match might be a subject
        # or a city, not just a name.
        matches = [
            any(needle in str(value).lower() for value in row.values)
            for _, row in df.iterrows()
        ]

        df = df[matches]

    if len(df) != total_rows:
        st.caption(f"Showing {len(df)} of {total_rows} rows.")

    if df.empty:
        st.info("Nothing matches. Clear the search or filter to see the rest.")
        return

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
            SOURCE_FILE_ID: None,
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

    # ------------------------------------------------------------------
    #  Save corrections back to Drive.
    #
    #  Without this, a phone fixed here is lost the moment the file is
    #  downloaded: the master keeps the wrong value and the next run
    #  reuses it.
    # ------------------------------------------------------------------
    if st.button("Save these corrections to the master database",
                 key=f"save_{state_key}"):

        target_folder = folder_id_for_save()
        target_file = master_file_id_for_save()

        if not target_folder and not target_file:
            st.error(
                "Nowhere to save to. Set the master database file in "
                "Settings, or paste the Drive folder link above."
            )
            return

        try:
            updated, changed, appended = apply_edits(
                st.session_state.get("master_df"),
                strip_review_column(edited)
            )

            if not changed and not appended:
                st.info("No changes to save.")

            else:
                save_master(target_folder, updated, target_file)

                st.session_state["master_df"] = updated
                st.session_state.pop(f"{state_key}__edited", None)

                parts = []
                if changed:
                    parts.append(f"{changed} row(s) updated")
                if appended:
                    parts.append(f"{appended} row(s) added")

                st.success(
                    ", ".join(parts).capitalize()
                    + " in the master database."
                )

        except Exception as e:
            st.error(
                f"Could not save: {e}\n\n"
                "Your edits are still in the grid — download them now so "
                "they aren't lost."
            )

    # The review column is dropped so the file matches the import template
    # exactly, with no column the portal doesn't recognise.
    download_df = strip_review_column(edited)

    # And the hidden key, so the file matches the import template exactly.
    if SOURCE_FILE_ID in download_df.columns:
        download_df = download_df.drop(columns=[SOURCE_FILE_ID])

    st.download_button(
        label=f"Download portal import file ({len(download_df)} teachers)",
        data=to_excel_bytes(download_df),
        file_name=file_name,
        mime=EXCEL_MIME,
        type="primary",
        key=f"download_{state_key}"
    )

    scope = (
        "just the candidates added in this run"
        if state_key == "run_df"
        else "every candidate ever extracted"
    )

    st.caption(
        f"This file holds {scope}. Upload it in the portal under "
        "Import / Export -> Import teachers. Edits here change the "
        "downloaded file only — they are not written back to the master "
        "database in Drive."
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

    all_duplicates = [
        duplicate
        for group in groups
        for duplicate in group["duplicates"]
    ]

    # canTrash was reported for the SERVICE ACCOUNT. When credentials for
    # the file's OWNER are configured, the deletion runs as them instead,
    # so that verdict no longer applies for those files.
    def owner_of(duplicate):
        owners = duplicate.get("owners") or []
        return owners[0].get("emailAddress") if owners else None

    as_user = any(
        has_user_credentials(owner_of(d)) for d in all_duplicates
    )

    if as_user:
        trashable = [
            d for d in all_duplicates
            if has_user_credentials(owner_of(d))
            or d.get("capabilities", {}).get("canTrash") is not False
        ]

        missing_accounts = sorted({
            owner_of(d) for d in all_duplicates
            if d not in trashable and owner_of(d)
        })

        if missing_accounts:
            st.info(
                "No credentials for "
                + ", ".join(missing_accounts)
                + ". Add that account with `python cleanup_duplicates.py "
                "--setup` to delete its files from here too."
            )

    else:
        # Drive already said which files it will refuse. Offering a button
        # that is certain to fail is worse than not showing it.
        trashable = [
            duplicate
            for duplicate in all_duplicates
            if duplicate.get("capabilities", {}).get("canTrash") is not False
        ]

    rows = []
    blocked = []

    for group in groups:
        for duplicate in group["duplicates"]:

            capabilities = duplicate.get("capabilities", {})

            can_trash = capabilities.get("canTrash")
            can_edit = capabilities.get("canEdit")

            owner = "unknown"
            owners = duplicate.get("owners") or []

            if owners:
                owner = owners[0].get("emailAddress", "unknown")

            if can_trash is False:
                blocked.append((duplicate["name"], owner, can_edit, duplicate["id"]))

            rows.append({
                "Duplicate (will be trashed)": duplicate["name"],
                "Keeping": group["keep"]["name"],
                "Owner": owner,
                "Can trash": "yes" if can_trash else "no",
            })

    st.dataframe(pd.DataFrame(rows), width="stretch", hide_index=True)

    if blocked and not as_user:

        # canEdit separates the two very different causes. Without it the
        # advice is a guess, and the wrong guess wastes the user's time.
        read_only = [b for b in blocked if b[2] is False]

        if read_only:
            st.error(
                f"{len(read_only)} file(s) are shared with the app as "
                "**Viewer**. Right-click the folder in Drive -> Share -> "
                "change the service account to **Editor**, then re-run."
            )

        else:
            owners = sorted({owner for _, owner, _, _ in blocked})

            st.warning(
                f"These files are owned by **{', '.join(owners)}** — your own "
                "account — and Drive only lets the *owner* move a file to "
                "trash. The app has edit access but cannot delete on your "
                "behalf, so you'll need to remove these yourself.\n\n"
                "It's one click each:"
            )

            for name, _, _, file_id in blocked:
                st.markdown(
                    f"- [{name}](https://drive.google.com/file/d/{file_id}/view) "
                    "— open, then use the bin icon"
                )

            st.caption(
                "Nothing breaks if you leave them. Duplicates are already "
                "excluded from parsing, so they cost you nothing — deleting "
                "them just tidies the folder."
            )

    st.caption(
        "The copy already recorded in the master database is always the one "
        "kept — trashing it instead would make the next run re-parse that "
        "resume and pay for it again."
    )

    if not trashable:
        st.info(
            "The app can't trash these itself — it signs in as the service "
            "account, and Drive only lets a file's owner delete it.\n\n"
            "Use the links above to remove them, or enable in-app deletion "
            "by running `python cleanup_duplicates.py --setup` once and "
            "pasting the result into Settings -> Secrets."
        )
        return

    if as_user:
        st.caption(
            "Deleting as your own Google account, so these can be removed "
            "from here."
        )

    confirm = st.checkbox(
        f"Yes, move these {len(trashable)} file(s) to my Drive trash",
        key="dup_confirm"
    )

    if st.button("Move duplicates to trash", disabled=not confirm):

        trashed = 0
        failures = []

        for duplicate in trashable:

            owners = duplicate.get("owners") or []
            owner_email = owners[0].get("emailAddress") if owners else None

            try:
                trash_file(duplicate["id"], owner_email=owner_email)
                trashed += 1

            except Exception as e:
                failures.append((duplicate, str(e)))

        if trashed:
            st.success(
                f"Moved {trashed} file(s) to the Drive trash. They stay "
                "recoverable there for 30 days if this was a mistake."
            )

        if failures:

            denied = [
                (duplicate, message)
                for duplicate, message in failures
                if "insufficientFilePermissions" in message or "403" in message
            ]

            if denied:
                # canEdit separates the two causes; guessing wastes time.
                read_only = [
                    d for d, _ in denied
                    if d.get("capabilities", {}).get("canEdit") is False
                ]

                if read_only:
                    st.error(
                        f"{len(read_only)} file(s) are shared with the app as "
                        "**Viewer**. In Drive: right-click the folder -> "
                        "Share -> change the service account to **Editor**, "
                        "then re-run."
                    )

                else:

                    # Name BOTH accounts. A 403 here almost always means the
                    # saved token belongs to a different Gmail than the file
                    # owner, which is invisible without saying so.
                    acting_as = _whoami(
                        (denied[0][0].get("owners") or [{}])[0]
                        .get("emailAddress")
                    )

                    file_owner = (
                        (denied[0][0].get("owners") or [{}])[0]
                        .get("emailAddress", "unknown")
                    )

                    if acting_as and acting_as.lower() != file_owner.lower():
                        st.error(
                            f"Signed in as **{acting_as}**, but these files "
                            f"are owned by **{file_owner}**. Only the owner "
                            "can delete a file.\n\n"
                            "Add the owner's account:\n\n"
                            "`python cleanup_duplicates.py --setup --account "
                            f"{file_owner}`\n\n"
                            "then paste the new [[google_accounts]] block "
                            "into Settings -> Secrets."
                        )

                    else:
                        st.warning(
                            "Drive refused the deletion even though the app "
                            f"is signed in as {acting_as or 'your account'}. "
                            "If these files were shared into your Drive by "
                            "someone else, only they can delete them. "
                            "Otherwise remove them yourself:"
                        )

                    for duplicate, _ in denied:
                        st.markdown(
                            f"- [{duplicate['name']}]"
                            f"(https://drive.google.com/file/d/"
                            f"{duplicate['id']}/view) — open, then the bin icon"
                        )

                    st.caption(
                        "Or leave them. Duplicates are excluded before any "
                        "download or API call, so they cost you nothing."
                    )

            other = [
                f"{d['name']}: {m}" for d, m in failures
                if (d, m) not in denied
            ]

            if other:
                st.error(
                    "Could not trash some files:\n\n"
                    + "\n\n".join(f"- {f}" for f in other)
                )

        if trashed:
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
        state_key="run_df",
        file_name="guru_setu_teacher_import_this_run.xlsx",
        heading="This run — review and download"
    )

    render_review_and_download(
        state_key="master_df",
        file_name="guru_setu_teacher_import_all.xlsx",
        heading="Master database — review and download"
    )


def blank_safe(value):
    """
    Text of a field that may be None, "None" or NaN, as a clean string.

    Missing values must come back as "" so they are falsy. Anything that
    turns a missing value into a non-empty string will make two records
    that are both missing that field look identical to each other.
    """

    if value is None:
        return ""

    text = str(value).strip()

    if text.lower() in ("none", "nan", "null"):
        return ""

    return text


def hash_text(text):
    normalized = " ".join(text.split()).lower()
    return hashlib.sha256(normalized.encode("utf-8")).hexdigest()

def process_resume(file_info, output_dir, existing_hashes):
    file_id = file_info["id"]
    file_name = file_info["name"]

    file_path = os.path.join(output_dir, file_name)

    try:

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
        # An uploaded file has no Drive page to link to.
        data["resume_link"] = (
            "" if str(file_id).startswith("upload-")
            else f"https://drive.google.com/file/d/{file_id}/view"
        )
        data["resume_file_name"] = file_name
        data["duplicate_of_content"] = False
        data["ocr_used"] = ocr_used
        return data

    except Exception as e:

        return {"error": f"{file_name}: {str(e)}"}

    finally:

        # Delete the CV as soon as it has been parsed. Keeping ~1000 PDFs
        # on disk for the whole run is what exhausts the container: the
        # free tier has about 1 GB, and a large folder blows past it long
        # before the run finishes, killing it with no output at all.
        try:
            if os.path.exists(file_path):
                os.remove(file_path)

        except OSError:
            pass


if st.button("Process Resumes"):

    if not drive_link and not uploaded_files:
        st.error(
            "Paste a Google Drive folder link, or upload some CVs."
        )
        st.stop()

    output_dir = "temp_resumes"

    if os.path.exists(output_dir):
        shutil.rmtree(output_dir)

    os.makedirs(output_dir)

    tracker.reset()

    st.write("Fetching files from Google Drive...")

    links = [line.strip() for line in drive_link.splitlines() if line.strip()]

    folder_ids = []

    for link in links:

        try:
            folder_ids.append(extract_folder_id(link))

        except ValueError:
            st.error(
                f"Not a Drive folder link: {link}\n\n"
                "Paste the URL from your browser's address bar."
            )
            st.stop()

    # The master is written to the first folder when no file is pinned.
    folder_id = folder_ids[0] if folder_ids else None

    if folder_id is None and not extract_file_id(master_link):
        st.error(
            "Uploads alone have nowhere to store the master database. "
            "Either set the master file in Settings, or paste a Drive "
            "folder link so one can be created there."
        )
        st.stop()

    files = []
    skipped_files = []

    for link in links:

        found, skipped = get_files_from_folder(
            link, recursive=include_subfolders
        )

        files.extend(found)
        skipped_files.extend(skipped)

    # The same file can sit in two of the folders given; parse it once.
    by_id = {}

    for item in files:
        by_id[item["id"]] = item

    files = list(by_id.values())

    if len(links) > 1:
        st.write(f"Read {len(links)} folders.")

    # Uploaded CVs are written straight to the work directory and given a
    # synthetic id derived from their bytes — so re-uploading the same file
    # is recognised as already parsed, exactly like a Drive file id.
    for upload in uploaded_files or []:

        data = upload.getvalue()

        upload_id = "upload-" + hashlib.sha256(data).hexdigest()[:24]

        with open(os.path.join(output_dir, upload.name), "wb") as handle:
            handle.write(data)

        files.append({
            "id": upload_id,
            "name": upload.name,
            "mimeType": upload.type or "application/pdf",
            "uploaded": True,
        })

    if uploaded_files:
        st.write(f"{len(uploaded_files)} uploaded file(s) added.")

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
        source = (
            f"file `{master_file_id}`"
            if master_file_id
            else f"`{MASTER_FILE_NAME}` in the resumes folder"
        )

        st.info(
            f"Master database has {len(master_df)} candidates from previous "
            f"runs, read from {source}. Files already in it will be skipped."
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

    st.write("Downloading + processing resumes...")

    results = []
    download_failures = []

    already_local = [f for f in files if f.get("uploaded")]
    files = [f for f in files if not f.get("uploaded")]

    queue = already_local + files
    total = len(queue)

    # ------------------------------------------------------------------
    #  Batched processing.
    #
    #  Submitting a thousand downloads at once fills the container's disk
    #  long before parsing drains it — downloads run 5-wide, parsing 2-wide
    #  — and the free tier is killed at about 1 GB with no output at all.
    #
    #  Batching also means the master is saved as the run progresses, so a
    #  crash at file 900 keeps the first 880 rather than losing everything.
    # ------------------------------------------------------------------
    batch_size = max(int(get_secret("BATCH_SIZE", 40) or 40), 1)

    overall_progress = st.progress(0.0)
    status = st.empty()

    processed_count = 0
    saved_total = 0

    for batch_start in range(0, total, batch_size):

        batch = queue[batch_start:batch_start + batch_size]

        status.write(
            f"Batch {batch_start // batch_size + 1} of "
            f"{(total + batch_size - 1) // batch_size} — "
            f"{processed_count} of {total} files done"
        )

        download_executor = ThreadPoolExecutor(max_workers=DOWNLOAD_WORKERS)
        groq_executor = ThreadPoolExecutor(max_workers=OPENAI_WORKERS)

        process_futures = {}

        to_download = [f for f in batch if not f.get("uploaded")]

        download_futures = {
            download_executor.submit(
                download_file, f["id"], f["name"], output_dir
            ): f
            for f in to_download
        }

        # Uploads are already on disk, so they skip straight to parsing.
        for file_info in [f for f in batch if f.get("uploaded")]:
            process_futures[
                groq_executor.submit(
                    process_resume, file_info, output_dir, existing_hashes
                )
            ] = file_info

        for future in as_completed(download_futures):

            file_info = download_futures[future]

            try:
                future.result()

                process_futures[
                    groq_executor.submit(
                        process_resume, file_info, output_dir, existing_hashes
                    )
                ] = file_info

            except Exception as e:
                download_failures.append(
                    {"name": file_info["name"], "error": str(e)}
                )
                st.error(f"Download failed: {file_info['name']} ({str(e)})")

        for future in as_completed(process_futures):

            processed_count += 1

            if total:
                overall_progress.progress(min(processed_count / total, 1.0))

            result = future.result()

            if not result:
                continue

            if "error" in result:
                st.error(result["error"])
                continue

            if (
                not result.get("extraction_failed")
                and not result.get("duplicate_of_content")
            ):

                # blank_safe, NOT str(...): result["email"] is None when no
                # address was found, and str(None) is the truthy string
                # "none" — which then matches every other record that also
                # had no email, silently dropping them.
                email = blank_safe(result.get("email")).lower()
                phone = blank_safe(result.get("phone"))

                if (
                    (email and email in existing_emails)
                    or (phone and phone in existing_phones)
                ):
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

        # Save what this batch produced before starting the next one.
        batch_keepers = [
            r for r in results[saved_total:] if r.get("keep")
        ]

        saved_total = len(results)

        if batch_keepers:

            try:
                master_df, batch_added, _ = merge_into_master(
                    master_df, to_master_rows(batch_keepers)
                )

                save_master(folder_id, master_df, master_file_id)

            except Exception as e:
                st.error(
                    f"Could not save after this batch: {e} "
                    "Processing continues, but stop and fix this — later "
                    "batches may be lost too."
                )

    status.write(f"Processed {processed_count} of {total} files.")

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

    cost_line = tracker.summary()

    if cost_line:
        st.caption(cost_line)

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

        # Rows were merged and saved batch by batch as the run progressed.
        # Merging again is a no-op for the data — every row is already
        # present — but it recomputes the report so the duplicate summary
        # below covers the whole run rather than only the last batch.
        master_df, _, merge_report = merge_into_master(master_df, new_rows)

        added = len([r for r in unique_results if r.get("keep")])

        from_master = [
            item for item in merge_report
            if item.get("origin") == "master"
        ]

        if from_master:
            st.warning(
                f"{len(from_master)} duplicate row(s) that were already in "
                "the master database have been cleaned up. This is why the "
                "candidate count can go DOWN — those rows were duplicates of "
                "each other, not new data being lost. It happens once; "
                "later runs will show none."
            )

        if merge_report:
            st.warning(
                f"{len(merge_report)} row(s) were merged as duplicates. "
                "If any of these are actually different people, tell me — "
                "the rule is name + phone, matching the portal's own."
            )

            st.dataframe(
                pd.DataFrame([
                    {
                        "Dropped": item["dropped"],
                        "Because": item["reason"],
                        "Same as": item["matched"],
                        "From": (
                            "already in master"
                            if item.get("origin") == "master"
                            else "this run"
                        ),
                    }
                    for item in merge_report
                ]),
                width="stretch",
                hide_index=True
            )

        try:
            # Safety net: batches already saved, this catches anything the
            # final merge changed.
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

        # This run's rows on their own, for reviewing just what was added
        # without scrolling a master that may hold thousands of rows.
        #
        # Taken from the MERGED frame rather than from new_rows, so a row
        # dropped as a duplicate doesn't appear here as if it had been saved.
        run_ids = {
            str(value).strip()
            for value in new_rows.get(SOURCE_FILE_ID, [])
            if str(value).strip()
        }

        if run_ids and SOURCE_FILE_ID in master_df.columns:
            st.session_state["run_df"] = master_df[
                master_df[SOURCE_FILE_ID].astype(str).str.strip().isin(run_ids)
            ].reset_index(drop=True)
        else:
            st.session_state.pop("run_df", None)

        st.session_state["master_df"] = master_df
        st.session_state.pop("master_df__edited", None)
        st.session_state.pop("run_df__edited", None)

    else:
        st.session_state.pop("master_df", None)
        st.session_state.pop("master_df__edited", None)
        st.session_state.pop("run_df", None)
        st.session_state.pop("run_df__edited", None)

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
