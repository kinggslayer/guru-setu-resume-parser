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
    """
    The folder id from a Drive URL.

    Trailing slashes, query strings and any path after the id are stripped.
    A URL copied with a trailing slash previously yielded "1ABC/", which
    Drive rejects — and the run reported an empty folder rather than a bad
    link, which is a confusing way to fail.
    """

    text = str(folder_url or "").strip()

    if "/folders/" in text:

        tail = text.split("/folders/")[1]

        # Cut at the first query string, fragment or further path segment.
        for separator in ("?", "#", "/"):
            tail = tail.split(separator)[0]

        folder_id = tail.strip()

        if folder_id:
            return folder_id

    raise ValueError("Invalid Google Drive folder link")


FOLDER_MIME = "application/vnd.google-apps.folder"
SHORTCUT_MIME = "application/vnd.google-apps.shortcut"


def get_files_from_folder(folder_url, service=None, recursive=True):
    """
    Files in a Drive folder.

    recursive walks sub-folders too. Organising CVs by batch or month is
    the obvious thing to do, and without this those files are invisible —
    the app reports "no files" on a folder that plainly has some.
    """

    if service is None:
        service = get_drive_service()

    root_id = extract_folder_id(folder_url)

    all_files = []
    seen_folders = set()
    pending = [root_id]

    while pending:

        folder_id = pending.pop()

        # Drive allows a folder to appear under several parents; without
        # this guard that becomes an infinite loop.
        if folder_id in seen_folders:
            continue

        seen_folders.add(folder_id)

        for item in _list_folder(service, folder_id):

            if item.get("mimeType") == FOLDER_MIME:

                if recursive:
                    pending.append(item["id"])

                continue

            all_files.append(item)

    # A shortcut is a pointer, not a file. Its own mimeType is a Google
    # native type, so without this it would be silently skipped — and a
    # folder full of shortcuts to CVs would look empty.
    all_files = _resolve_shortcuts(service, all_files)

    downloadable = [
        f for f in all_files
        if not f.get("mimeType", "").startswith(GOOGLE_NATIVE_MIME_PREFIX)
    ]

    skipped = [
        f for f in all_files
        if f.get("mimeType", "").startswith(GOOGLE_NATIVE_MIME_PREFIX)
    ]

    return downloadable, skipped


def _resolve_shortcuts(service, files):
    """
    Replace each shortcut with the file it points at.

    The target's own metadata is fetched, because the shortcut carries no
    checksum or owner of its own — and both are needed for duplicate
    detection and deletion.
    """

    resolved = []

    for item in files:

        if item.get("mimeType") != SHORTCUT_MIME:
            resolved.append(item)
            continue

        target_id = (item.get("shortcutDetails") or {}).get("targetId")

        if not target_id:
            continue

        try:
            target = service.files().get(
                fileId=target_id,
                fields=(
                    "id,name,mimeType,md5Checksum,size,createdTime,"
                    "ownedByMe,capabilities(canTrash,canEdit),"
                    "owners(emailAddress)"
                )
            ).execute()

            resolved.append(target)

        except Exception:
            # Target deleted or not shared — nothing to parse.
            continue

    return resolved


def _list_folder(service, folder_id):
    """Every direct child of one folder, following pagination."""

    items = []
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
                "owners(emailAddress),"
                "shortcutDetails(targetId,targetMimeType))"
            ),
            pageSize=1000,
            pageToken=page_token
        ).execute()

        items.extend(results.get("files", []))

        page_token = results.get("nextPageToken")

        if not page_token:
            break

    return items


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


def trash_file(file_id, service=None, owner_email=None):
    """
    Move a file to the Drive trash.

    owner_email selects which account to act as. Only a file's owner may
    trash it, so with several Gmails configured the right one has to be
    chosen per file — the service account can never do it.

    Deliberately NOT files().delete(), which is permanent and unrecoverable.
    Trashed files sit in Drive's bin for 30 days, so a wrong call here can
    be undone by the user.
    """

    if service is None:

        service = get_user_drive_service(owner_email)

        if service is None:
            # Deliberately NOT falling back to the service account: it can
            # never trash a file it doesn't own, so the fallback would only
            # turn a clear "no credentials" into a confusing 403.
            raise PermissionError(
                f"No Google account credentials configured for {owner_email}. "
                "Run: python cleanup_duplicates.py --setup --account "
                f"{owner_email}"
            )

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


def _accounts_from_secrets():
    """
    Every Google account the app can act as, as {email: refresh_token}.

    Two shapes are accepted. A [[google_accounts]] array of tables covers
    several Gmails; the single GOOGLE_OAUTH_REFRESH_TOKEN is kept working
    so an existing setup doesn't break when a second account is added.
    """

    accounts = {}

    try:
        listed = st.secrets.get("google_accounts", [])

    except Exception:
        listed = []

    for entry in listed or []:

        email = str(entry.get("email", "")).strip().lower()
        token = str(entry.get("refresh_token", "")).strip()

        if email and token:
            accounts[email] = token

    single = _oauth_setting("GOOGLE_OAUTH_REFRESH_TOKEN")

    if single:
        email = str(
            _oauth_setting("GOOGLE_OAUTH_EMAIL") or "default"
        ).strip().lower()

        accounts.setdefault(email, single)

    return accounts


def has_user_credentials(owner_email=None):
    """
    Whether the app can act as a user.

    owner_email asks specifically about that account — the answer differs
    per file, because a token for one Gmail cannot delete another Gmail's
    files, and claiming otherwise would produce a button that always fails.
    """

    if not (
        _oauth_setting("GOOGLE_OAUTH_CLIENT_ID")
        and _oauth_setting("GOOGLE_OAUTH_CLIENT_SECRET")
    ):
        return False

    accounts = _accounts_from_secrets()

    if not accounts:
        return False

    if owner_email is None:
        return True

    email = str(owner_email).strip().lower()

    # "default" is the single-account setup, which predates per-account
    # tokens and so can't be matched against an owner — assume it fits.
    return email in accounts or "default" in accounts


def get_user_drive_service(owner_email=None):
    """
    A Drive client authenticated as the user who owns the file.

    Returns None when no suitable account is configured, so callers can
    fall back rather than crash.
    """

    accounts = _accounts_from_secrets()

    if not accounts:
        return None

    email = str(owner_email or "").strip().lower()

    if email and email in accounts:
        token = accounts[email]
        cache_key = email

    elif "default" in accounts:
        token = accounts["default"]
        cache_key = "default"

    elif not email:
        cache_key = next(iter(accounts))
        token = accounts[cache_key]

    else:
        # A token exists, but not for this owner — using it would fail.
        return None

    cache = getattr(_thread_local, "user_services", None)

    if cache is None:
        cache = {}
        _thread_local.user_services = cache

    if cache_key in cache:
        return cache[cache_key]

    from google.oauth2.credentials import Credentials

    creds = Credentials(
        token=None,
        refresh_token=token,
        client_id=_oauth_setting("GOOGLE_OAUTH_CLIENT_ID"),
        client_secret=_oauth_setting("GOOGLE_OAUTH_CLIENT_SECRET"),
        token_uri="https://oauth2.googleapis.com/token",
        scopes=SCOPES,
    )

    service = build("drive", "v3", credentials=creds)

    cache[cache_key] = service

    return service




def whoami(owner_email=None):
    """
    The email of the account the app would act as for this owner.

    Used in error messages: a 403 while "acting as the user" almost always
    means the token belongs to a different Gmail than the file's owner, and
    naming both makes that obvious instead of mysterious.
    """

    service = get_user_drive_service(owner_email)

    if service is None:
        return None

    try:
        about = service.about().get(fields="user(emailAddress)").execute()

        return about.get("user", {}).get("emailAddress")

    except Exception:
        return None
