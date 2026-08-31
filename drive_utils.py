from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

import io
import os
import threading
import streamlit as st

from concurrent.futures import ThreadPoolExecutor, as_completed

SCOPES = ["https://www.googleapis.com/auth/drive"]

GOOGLE_NATIVE_MIME_PREFIX = "application/vnd.google-apps"

_thread_local = threading.local()


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


def get_drive_service():

    if not hasattr(_thread_local, "service"):

        creds = _get_credentials()

        _thread_local.service = build(
            "drive",
            "v3",
            credentials=creds
        )

    return _thread_local.service


def extract_folder_id(folder_url):

    if "/folders/" in folder_url:
        return folder_url.split("/folders/")[1].split("?")[0]

    raise ValueError("Invalid Google Drive folder link")


def get_files_from_folder(folder_url, service=None):

    if service is None:
        service = get_drive_service()

    folder_id = extract_folder_id(folder_url)

    all_files = []
    page_token = None

    while True:

        results = service.files().list(
            q=f"'{folder_id}' in parents and trashed=false",
            # md5Checksum is computed by Drive itself, so byte-identical
            # copies can be spotted without downloading anything.
            # md5Checksum is computed by Drive itself, so byte-identical
            # copies can be spotted without downloading anything.
            #
            # capabilities/owners answer "why can't this be trashed?"
            # definitively, instead of leaving the user to guess between
            # wrong sharing level and wrong file owner.
            fields=(
                "nextPageToken, "
                "files(id,name,mimeType,md5Checksum,size,createdTime,"
                "ownedByMe,capabilities(canTrash,canEdit),"
                "owners(emailAddress))"
            ),
            pageSize=1000,
            pageToken=page_token
        ).execute()

        all_files.extend(results.get("files", []))

        page_token = results.get("nextPageToken")

        if not page_token:
            break
    downloadable = [
        f for f in all_files
        if not f.get("mimeType", "").startswith(GOOGLE_NATIVE_MIME_PREFIX)
    ]

    skipped = [
        f for f in all_files
        if f.get("mimeType", "").startswith(GOOGLE_NATIVE_MIME_PREFIX)
    ]

    return downloadable, skipped


def download_file(file_id, file_name, output_dir):

    service = get_drive_service()

    request = service.files().get_media(fileId=file_id)

    file_path = os.path.join(output_dir, file_name)

    fh = io.FileIO(file_path, "wb")

    downloader = MediaIoBaseDownload(fh, request)

    done = False

    while not done:
        status, done = downloader.next_chunk()

    fh.close()

    return file_path


def download_all_files(files, output_dir, max_workers=5, progress_callback=None):

    successful = []
    failed = []
    completed = 0
    total = len(files)

    with ThreadPoolExecutor(max_workers=max_workers) as executor:

        future_to_file = {
            executor.submit(
                download_file, f["id"], f["name"], output_dir
            ): f
            for f in files
        }

        for future in as_completed(future_to_file):

            file_info = future_to_file[future]
            completed += 1

            if progress_callback:
                progress_callback(completed, total)

            try:
                path = future.result()
                successful.append(path)

            except Exception as e:
                failed.append({"name": file_info["name"], "error": str(e)})

    return successful, failed

def _find_file_in_folder(service, folder_id, file_name):

    results = service.files().list(
        q=f"'{folder_id}' in parents and name='{file_name}' and trashed=false",
        fields="files(id,name)"
    ).execute()

    files = results.get("files", [])

    return files[0]["id"] if files else None


def download_cache_from_drive(folder_id, file_name, service=None):
    """
    Returns the raw JSON text of file_name from the given Drive folder,
    or None if it doesn't exist yet (first-ever run).
    """

    if service is None:
        service = get_drive_service()

    file_id = _find_file_in_folder(service, folder_id, file_name)

    if file_id is None:
        return None

    request = service.files().get_media(fileId=file_id)

    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)

    done = False
    while not done:
        status, done = downloader.next_chunk()

    return buf.getvalue().decode("utf-8")


def upload_cache_to_drive(folder_id, file_name, json_text, service=None):
    """
    Creates or updates file_name inside folder_id with json_text content.
    Update (not create-a-new-copy) if a file with that name already exists,
    so we don't accumulate duplicate cache files on every run.
    """

    if service is None:
        service = get_drive_service()

    media = MediaIoBaseUpload(
        io.BytesIO(json_text.encode("utf-8")),
        mimetype="application/json",
        resumable=False
    )

    existing_id = _find_file_in_folder(service, folder_id, file_name)

    if existing_id:
        service.files().update(
            fileId=existing_id,
            media_body=media
        ).execute()
        return existing_id

    else:
        created = service.files().create(
            body={"name": file_name, "parents": [folder_id]},
            media_body=media,
            fields="id"
        ).execute()
        return created["id"]

def download_bytes_from_drive(folder_id, file_name, service=None):
    """
    Raw bytes of file_name from a Drive folder, or None if it isn't there
    yet. Used for the master workbook, which is binary rather than JSON.
    """

    if service is None:
        service = get_drive_service()

    file_id = _find_file_in_folder(service, folder_id, file_name)

    if file_id is None:
        return None

    request = service.files().get_media(fileId=file_id)

    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)

    done = False
    while not done:
        status, done = downloader.next_chunk()

    return buf.getvalue()


def upload_bytes_to_drive(folder_id, file_name, data, mimetype, service=None):
    """
    Create or UPDATE file_name in folder_id. Updating in place matters for
    the master database: creating each time would leave a trail of files
    all called the same thing, and the next run would pick one at random.
    """

    if service is None:
        service = get_drive_service()

    media = MediaIoBaseUpload(
        io.BytesIO(data),
        mimetype=mimetype,
        resumable=False
    )

    existing_id = _find_file_in_folder(service, folder_id, file_name)

    if existing_id:
        service.files().update(
            fileId=existing_id,
            media_body=media
        ).execute()
        return existing_id

    created = service.files().create(
        body={"name": file_name, "parents": [folder_id]},
        media_body=media,
        fields="id"
    ).execute()

    return created["id"]


def get_file_metadata(file_id, service=None):
    """Name and mimeType for a file id, used to sanity-check the master."""

    if service is None:
        service = get_drive_service()

    return service.files().get(
        fileId=file_id,
        fields="id,name,mimeType"
    ).execute()


def download_bytes_by_id(file_id, service=None):
    """Raw bytes of a Drive file by id."""

    if service is None:
        service = get_drive_service()

    request = service.files().get_media(fileId=file_id)

    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)

    done = False
    while not done:
        status, done = downloader.next_chunk()

    return buf.getvalue()


def upload_bytes_by_id(file_id, data, mimetype, service=None):
    """Replace the contents of an existing Drive file, in place."""

    if service is None:
        service = get_drive_service()

    media = MediaIoBaseUpload(
        io.BytesIO(data),
        mimetype=mimetype,
        resumable=False
    )

    return service.files().update(
        fileId=file_id,
        media_body=media
    ).execute()


def trash_file(file_id, service=None):
    """
    Move a file to the Drive trash.

    Prefers the user's own credentials when configured, because the service
    account cannot trash files it does not own — which is every file the
    user uploaded themselves.

    Deliberately NOT files().delete(), which is permanent and unrecoverable.
    Trashed files sit in Drive's bin for 30 days, so a wrong call here can
    be undone by the user.
    """

    if service is None:
        service = get_user_drive_service() or get_drive_service()

    return service.files().update(
        fileId=file_id,
        body={"trashed": True}
    ).execute()


# ---------------------------------------------------------------------------
#  Acting as the user, rather than as the service account
#
#  Drive only lets a file's OWNER trash it, and the service account is a
#  separate identity from the user's Gmail — so it can read and edit the
#  resumes but can never delete them.
#
#  If the user supplies OAuth credentials of their own, Drive calls that need
#  ownership are made as them instead. Minting the refresh token is a one-off,
#  done by cleanup_duplicates.py, which prints the block to paste into
#  Settings -> Secrets.
# ---------------------------------------------------------------------------

USER_OAUTH_KEYS = (
    "GOOGLE_OAUTH_CLIENT_ID",
    "GOOGLE_OAUTH_CLIENT_SECRET",
    "GOOGLE_OAUTH_REFRESH_TOKEN",
)


def _oauth_setting(name):

    try:
        if name in st.secrets:
            return st.secrets[name]

    except Exception:
        pass

    return os.getenv(name)


def has_user_credentials():
    """True when all three OAuth values are configured."""

    return all(_oauth_setting(key) for key in USER_OAUTH_KEYS)


def get_user_drive_service():
    """
    A Drive client authenticated as the user.

    Returns None when OAuth isn't configured, so callers can fall back to
    the service account rather than crash.
    """

    if not has_user_credentials():
        return None

    if getattr(_thread_local, "user_service", None) is not None:
        return _thread_local.user_service

    from google.oauth2.credentials import Credentials

    creds = Credentials(
        token=None,
        refresh_token=_oauth_setting("GOOGLE_OAUTH_REFRESH_TOKEN"),
        client_id=_oauth_setting("GOOGLE_OAUTH_CLIENT_ID"),
        client_secret=_oauth_setting("GOOGLE_OAUTH_CLIENT_SECRET"),
        token_uri="https://oauth2.googleapis.com/token",
        scopes=SCOPES,
    )

    service = build("drive", "v3", credentials=creds)

    _thread_local.user_service = service

    return service
