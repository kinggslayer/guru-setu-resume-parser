"""
Delete duplicate resume files from a Drive folder.

WHY THIS IS SEPARATE FROM THE APP
---------------------------------
The Streamlit app authenticates as a *service account*, which is its own
Google identity. Drive only lets a file's OWNER move it to trash, and you
own these files — so the app can find duplicates but can never remove them.

This script signs in as YOU instead, so the deletions are permitted.

Run it on your own computer; it opens a browser for the Google sign-in.

    python cleanup_duplicates.py <folder-link>              # preview only
    python cleanup_duplicates.py <folder-link> --delete     # actually trash

It defaults to a preview. Nothing is touched until you pass --delete, and
even then files go to the Drive bin (recoverable for 30 days), never to a
permanent delete.

ONE-TIME SETUP
--------------
1. https://console.cloud.google.com -> your project (guru-setu-parser)
2. APIs & Services -> OAuth consent screen
     User type: External -> Create
     App name: anything; your email in both email boxes -> Save
     Audience -> Test users -> Add your own Gmail address
3. APIs & Services -> Credentials -> Create credentials
     -> OAuth client ID -> Application type: Desktop app -> Create
4. Download the JSON, save it next to this script as client_secret.json

The first run opens a browser and asks you to sign in. It saves token.json
so later runs don't ask again. Both files are already in .gitignore.
"""

import os
import sys

from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

from duplicates import find_duplicate_groups, summarise


# Full drive scope: trashing a file is a write, so readonly won't do.
SCOPES = ["https://www.googleapis.com/auth/drive"]

CLIENT_SECRET_FILE = "client_secret.json"
TOKEN_FILE = "token.json"


def get_service():
    """Sign in as the user, reusing the saved token when it's still good."""

    creds = None

    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    if not creds or not creds.valid:

        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())

        else:

            if not os.path.exists(CLIENT_SECRET_FILE):
                sys.exit(
                    f"Missing {CLIENT_SECRET_FILE}. See the setup steps at "
                    "the top of this file."
                )

            flow = InstalledAppFlow.from_client_secrets_file(
                CLIENT_SECRET_FILE, SCOPES
            )

            creds = flow.run_local_server(port=0)

        with open(TOKEN_FILE, "w", encoding="utf-8") as handle:
            handle.write(creds.to_json())

    return build("drive", "v3", credentials=creds)


def folder_id_from(link):

    if "/folders/" in link:
        return link.split("/folders/")[1].split("?")[0].split("/")[0]

    # Assume a bare id was passed.
    return link.strip()


def list_files(service, folder_id):
    """Every non-trashed file in the folder, with checksums."""

    files = []
    page_token = None

    while True:

        response = service.files().list(
            q=f"'{folder_id}' in parents and trashed=false",
            fields=(
                "nextPageToken, "
                "files(id,name,md5Checksum,size,createdTime,"
                "capabilities(canTrash))"
            ),
            pageSize=1000,
            pageToken=page_token
        ).execute()

        files.extend(response.get("files", []))

        page_token = response.get("nextPageToken")

        if not page_token:
            break

    return files


def main():

    if len(sys.argv) < 2:
        sys.exit(
            "Usage: python cleanup_duplicates.py <folder-link> [--delete]"
        )

    link = sys.argv[1]
    really_delete = "--delete" in sys.argv

    service = get_service()
    folder_id = folder_id_from(link)

    print("Reading folder...")
    files = list_files(service, folder_id)
    print(f"{len(files)} file(s) in the folder.\n")

    groups = find_duplicate_groups(files)

    if not groups:
        print("No duplicates. Nothing to do.")
        return

    duplicate_count, group_count = summarise(groups)

    print(f"{duplicate_count} duplicate file(s) in {group_count} set(s):\n")

    for group in groups:

        print(f"  KEEP   {group['keep']['name']}")

        for duplicate in group["duplicates"]:
            print(f"  DELETE {duplicate['name']}")

        print()

    if not really_delete:
        print(
            "Preview only — nothing was changed.\n"
            "Re-run with --delete to move these to the Drive bin."
        )
        return

    # Last chance: the list above is the plan, this is the commitment.
    answer = input(f"Move {duplicate_count} file(s) to the bin? [y/N] ")

    if answer.strip().lower() not in ("y", "yes"):
        print("Cancelled. Nothing was changed.")
        return

    trashed = 0

    for group in groups:

        for duplicate in group["duplicates"]:

            try:
                service.files().update(
                    fileId=duplicate["id"],
                    body={"trashed": True}
                ).execute()

                print(f"  trashed {duplicate['name']}")
                trashed += 1

            except Exception as e:
                print(f"  FAILED  {duplicate['name']}: {e}")

    print(
        f"\nDone. {trashed} file(s) moved to the Drive bin, where they stay "
        "recoverable for 30 days."
    )


if __name__ == "__main__":
    main()
