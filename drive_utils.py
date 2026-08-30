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
            fields="nextPageToken, files(id,name,mimeType)",
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