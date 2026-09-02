"""
Parse a Drive folder from the command line.

WHY THIS EXISTS
---------------
The Streamlit app's script is tied to the browser session. When the laptop
sleeps the websocket drops and the server cancels the run mid-batch — which
is why long folders never finish there.

This does the same work with no browser involved:

  * closing the terminal window is the only thing that stops it
  * if the laptop sleeps, the process is suspended and CONTINUES on wake,
    rather than being cancelled
  * run it on a server (or with nohup) and it is genuinely unattended

It writes to the same master file as the app, using the same dedup rules,
so the two are interchangeable — parse here, review and download there.

USAGE
-----
    python parse_folder.py <folder-link> [options]
    python parse_folder.py                 # uses DRIVE_FOLDER_LINK from .env

    --master <link>      the master workbook (or set MASTER_FILE_LINK)
    --batch <n>          files per save, default 25
    --limit <n>          stop after this many files, for a trial run
    --no-subfolders      only the folder itself
    --retry-failed       re-try files that failed on an earlier run

    python parse_folder.py <link> --limit 10     # cheap first test

Credentials come from .env or the environment, exactly as the app does:
OPENAI_API_KEY, and GOOGLE_SERVICE_ACCOUNT_FILE pointing at the service
account JSON.

RUNNING IT UNATTENDED
---------------------
    Windows :  start /b python parse_folder.py "<link>" > run.log 2>&1
    Mac/Linux: nohup python parse_folder.py "<link>" > run.log 2>&1 &

Progress is saved to Drive after every batch, so stopping it at any point
keeps everything done so far. Re-running the same folder continues from
there and never re-pays for a file already in the master.
"""

import argparse
import hashlib
import os
import re
import shutil
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

from dotenv import load_dotenv

load_dotenv()

from drive_utils import (
    get_files_from_folder,
    download_file,
    extract_folder_id,
    local_name_for,
)
from extractor import (
    read_pdf,
    read_docx,
    extract_resume_data,
    ocr_pdf,
    MIN_USABLE_CHARS,
)
from master_store import (
    load_master,
    save_master,
    already_seen,
    to_master_rows,
    merge_into_master,
    extract_file_id,
    MASTER_FILE_NAME,
)
from duplicates import find_duplicate_groups, summarise
from skiplist import load_failed, save_failed, reason_for
from jobs import claim_next_job, update_job
from usage import tracker


WORK_DIR = "temp_resumes_cli"

DOWNLOAD_WORKERS = 5

SAVE_ATTEMPTS = 3
SAVE_RETRY_DELAY = 3


def log(message):
    """Timestamped, flushed — so a redirected log file stays current."""

    print(f"[{time.strftime('%H:%M:%S')}] {message}", flush=True)


def blank_safe(value):
    """Missing values as "", never the truthy string "None"."""

    if value is None:
        return ""

    text = str(value).strip()

    return "" if text.lower() in ("none", "nan", "null") else text


def hash_text(text):
    normalised = re.sub(r"\s+", " ", str(text or "")).strip().lower()

    return hashlib.sha256(normalised.encode("utf-8")).hexdigest()


def process_resume(file_info, existing_hashes):
    """Mirrors the app's process_resume, minus the UI."""

    file_id = file_info["id"]
    file_name = file_info["name"]

    file_path = os.path.join(WORK_DIR, local_name_for(file_id, file_name))

    try:
        ocr_used = False
        lower = file_name.lower()

        if lower.endswith(".pdf"):
            text = read_pdf(file_path)

            if len(text.strip()) < MIN_USABLE_CHARS:
                text = ocr_pdf(file_path)
                ocr_used = bool(text.strip())

        elif lower.endswith(".docx"):
            text = read_docx(file_path)

        else:
            return {"error": f"{file_name}: unsupported file type"}

        content_hash = hash_text(text)

        # Every unreadable file hashes identically; such a hash must never
        # be used to call two resumes the same.
        if len(text.strip()) < MIN_USABLE_CHARS:
            content_hash = ""

        if content_hash and content_hash in existing_hashes:
            return {
                "duplicate_of_content": True,
                "resume_file_name": file_name,
                "content_hash": content_hash,
            }

        data = extract_resume_data(text)
        data["content_hash"] = content_hash
        data["file_id"] = file_id
        data["resume_link"] = (
            f"https://drive.google.com/file/d/{file_id}/view"
        )
        data["resume_file_name"] = file_name
        data["duplicate_of_content"] = False
        data["ocr_used"] = ocr_used

        return data

    except Exception as e:
        return {"error": f"{file_name}: {e}"}

    finally:
        # Free the disk immediately; a thousand PDFs held at once is what
        # exhausts a small container.
        try:
            if os.path.exists(file_path):
                os.remove(file_path)

        except OSError:
            pass


def save_with_retries(folder_id, master_df, master_file_id):
    """Save, retrying a momentary Drive failure. False if it never lands."""

    for attempt in range(SAVE_ATTEMPTS):

        try:
            save_master(folder_id, master_df, master_file_id)
            return True

        except Exception as e:

            if attempt < SAVE_ATTEMPTS - 1:
                delay = SAVE_RETRY_DELAY * (attempt + 1)
                log(f"  save failed ({e}); retrying in {delay}s")
                time.sleep(delay)

            else:
                log(f"  SAVE FAILED after {SAVE_ATTEMPTS} attempts: {e}")

    return False


def watch(args):
    """
    Worker loop: take queued folders and parse them.

    Runs anywhere that stays awake — the EC2 box, or a desktop left on.
    Staff queue folders from the web app and never touch a terminal.
    """

    queue_folder = extract_folder_id(args.watch)

    worker_name = os.getenv("WORKER_NAME") or f"worker-{os.getpid()}"

    log(f"Watching the queue in folder {queue_folder} as {worker_name}.")
    log(f"Checking every {args.poll}s. Ctrl+C to stop.")

    while True:

        try:
            job = claim_next_job(queue_folder, worker_name)

        except Exception as e:
            log(f"Could not read the queue: {e}")
            time.sleep(args.poll)
            continue

        if job is None:
            time.sleep(args.poll)
            continue

        log(f"Job {job['id']}: {job.get('folder_link')}")

        try:
            job_args = argparse.Namespace(
                folder=job.get("folder_link"),
                master=job.get("master_link") or None,
                batch=args.batch,
                limit=0,
                no_subfolders=False,
                retry_failed=False,
                watch=None,
                poll=args.poll,
                _job=(queue_folder, job["id"]),
            )

            code = run_once(job_args)

            update_job(
                queue_folder, job["id"],
                status="done" if code == 0 else "failed",
                progress="finished" if code == 0 else "stopped early",
            )

            log(f"Job {job['id']} {'done' if code == 0 else 'failed'}.")

        except KeyboardInterrupt:
            update_job(queue_folder, job["id"], status="queued",
                       progress="interrupted, will retry")
            log("Interrupted; job returned to the queue.")
            raise

        except Exception as e:
            # Never let one bad job stop the worker.
            update_job(queue_folder, job["id"], status="failed", error=str(e))
            log(f"Job {job['id']} failed: {e}")


def main():

    parser = argparse.ArgumentParser(
        description="Parse a Google Drive folder of resumes."
    )
    parser.add_argument(
        "folder",
        nargs="?",
        default=None,
        help="Drive folder link (or set DRIVE_FOLDER_LINK in .env)"
    )
    parser.add_argument("--master", default=None, help="master workbook link")
    parser.add_argument("--batch", type=int, default=25)
    parser.add_argument("--limit", type=int, default=0,
                        help="stop after this many files (trial run)")
    parser.add_argument("--no-subfolders", action="store_true")
    parser.add_argument("--retry-failed", action="store_true")
    parser.add_argument(
        "--watch", metavar="FOLDER_LINK", default=None,
        help="run as a worker: poll that folder's queue and parse what "
             "staff submit from the app"
    )
    parser.add_argument("--poll", type=int, default=60,
                        help="seconds between queue checks in --watch mode")
    args = parser.parse_args()

    if args.watch:
        return watch(args)

    return run_once(args)


def run_once(args):
    """Parse one folder. Shared by the CLI and the worker."""

    if not os.getenv("OPENAI_API_KEY"):
        sys.exit("OPENAI_API_KEY is not set. Put it in .env next to this file.")

    # The link may be given on the command line or, for a folder parsed
    # regularly, left in .env so the command is just `python parse_folder.py`.
    folder_link = args.folder or os.getenv("DRIVE_FOLDER_LINK") or ""

    # Set when running a queued job, so progress reaches the app.
    job_ref = getattr(args, "_job", None)

    if not folder_link.strip():
        sys.exit(
            "No folder given.\n\n"
            "  python parse_folder.py \"https://drive.google.com/drive/folders/...\"\n\n"
            "or put DRIVE_FOLDER_LINK=... in .env and run it with no arguments."
        )

    try:
        folder_id = extract_folder_id(folder_link)

    except ValueError:
        sys.exit(
            f"That is not a Drive folder link: {folder_link}\n"
            "Copy the URL from your browser while the folder is open."
        )

    master_file_id = extract_file_id(
        args.master or os.getenv("MASTER_FILE_LINK") or ""
    )

    if os.path.exists(WORK_DIR):
        shutil.rmtree(WORK_DIR)

    os.makedirs(WORK_DIR)

    tracker.reset()

    # ---------------------------------------------------------------
    #  What to parse
    # ---------------------------------------------------------------
    log("Listing the folder...")

    files, skipped = get_files_from_folder(
        folder_link, recursive=not args.no_subfolders
    )

    files = [f for f in files if f["name"] != MASTER_FILE_NAME]

    log(f"{len(files)} downloadable file(s); {len(skipped)} Google-native skipped")

    master_df, existed = load_master(folder_id, master_file_id)

    log(f"Master database: {len(master_df)} candidates"
        + ("" if existed else " (new file will be created)"))

    seen_ids, seen_hashes = already_seen(master_df)

    groups = find_duplicate_groups(files, seen_ids)

    if groups:
        extra, sets_, involved, _ = summarise(groups, len(files))
        duplicate_ids = {d["id"] for g in groups for d in g["duplicates"]}
        files = [f for f in files if f["id"] not in duplicate_ids]
        log(f"{extra} byte-identical duplicate copies skipped ({sets_} sets)")

    before = len(files)
    files = [f for f in files if f["id"] not in seen_ids]
    log(f"{before - len(files)} already in the master, skipped")

    previously_failed = load_failed(folder_id)

    if previously_failed and not args.retry_failed:
        before = len(files)
        files = [f for f in files if f["id"] not in previously_failed]

        if before - len(files):
            log(f"{before - len(files)} failed previously, skipped "
                "(use --retry-failed to include them)")

    if args.limit:
        files = files[:args.limit]
        log(f"--limit {args.limit}: only parsing {len(files)} file(s)")

    if not files:
        log("Nothing new to parse. The master is up to date.")
        shutil.rmtree(WORK_DIR, ignore_errors=True)
        return 0

    # ---------------------------------------------------------------
    #  Parse, saving after each batch
    # ---------------------------------------------------------------
    existing_hashes = set(seen_hashes)
    existing_emails = set()
    existing_phones = set()

    workers = max(int(os.getenv("OPENAI_WORKERS", "2")), 1)

    total = len(files)
    total_batches = (total + args.batch - 1) // args.batch

    done = 0
    added_total = 0

    # Cumulative reasons, so "N parsed but only M saved" is answerable at
    # any point rather than being a mystery at the end.
    tally = {
        "added": 0, "already_in_master": 0, "same_person_this_run": 0,
        "identical_text": 0, "scans": 0, "failed": 0,
    }

    log(f"Parsing {total} file(s) in {total_batches} batch(es) of {args.batch}")

    for start in range(0, total, args.batch):

        batch = files[start:start + args.batch]
        batch_number = start // args.batch + 1

        downloader = ThreadPoolExecutor(max_workers=DOWNLOAD_WORKERS)
        parser_pool = ThreadPoolExecutor(max_workers=workers)

        parse_futures = {}

        download_futures = {
            downloader.submit(download_file, f["id"], f["name"], WORK_DIR): f
            for f in batch
        }

        for future in as_completed(download_futures):

            info = download_futures[future]

            try:
                future.result()
                parse_futures[
                    parser_pool.submit(process_resume, info, existing_hashes)
                ] = info

            except Exception as e:
                log(f"  download failed: {info['name']} ({e})")

        keepers = []

        for future in as_completed(parse_futures):

            done += 1
            result = future.result()

            if not result or "error" in result:
                if result:
                    log(f"  {result['error']}")
                continue

            if result.get("duplicate_of_content"):
                tally["identical_text"] += 1
                continue

            if result.get("extraction_failed"):

                if result.get("needs_ocr"):
                    tally["scans"] += 1
                else:
                    tally["failed"] += 1

                if reason_for(result) and result.get("file_id"):
                    previously_failed[str(result["file_id"])] = reason_for(result)

                continue

            email = blank_safe(result.get("email")).lower()
            phone = blank_safe(result.get("phone"))

            if (email and email in existing_emails) or (
                phone and phone in existing_phones
            ):
                tally["same_person_this_run"] += 1
                continue

            if blank_safe(result.get("content_hash")):
                existing_hashes.add(str(result["content_hash"]))

            if email:
                existing_emails.add(email)

            if phone:
                existing_phones.add(phone)

            keepers.append(result)

        downloader.shutdown(wait=True)
        parser_pool.shutdown(wait=True)

        save_failed(folder_id, previously_failed)

        batch_added = 0

        if keepers:

            master_df, batch_added, batch_report = merge_into_master(
                master_df, to_master_rows(keepers)
            )

            tally["already_in_master"] += len([
                item for item in batch_report
                if item.get("origin") != "master"
            ])

            if not save_with_retries(folder_id, master_df, master_file_id):
                log("Stopping so no more money is spent on unsavable results. "
                    "Everything from earlier batches is safe.")
                shutil.rmtree(WORK_DIR, ignore_errors=True)
                return 1

        added_total += batch_added
        tally["added"] += batch_added

        log(f"Batch {batch_number}/{total_batches}: "
            f"{len(batch)} files, {len(keepers)} parsed, {batch_added} added, "
            f"master now {len(master_df)} | {done}/{total} done")

        if job_ref:
            update_job(
                job_ref[0], job_ref[1],
                progress=(
                    f"batch {batch_number}/{total_batches}, "
                    f"{done}/{total} files, {added_total} added"
                )
            )

    shutil.rmtree(WORK_DIR, ignore_errors=True)

    log(f"Finished. {added_total} new candidate(s); "
        f"the master holds {len(master_df)}.")

    accounted = sum(tally.values())

    log("Where the parsed files went:")
    log(f"  added to the master           : {tally['added']}")
    log(f"  already there (same person)   : {tally['already_in_master']}")
    log(f"  repeated within this run      : {tally['same_person_this_run']}")
    log(f"  identical text to another CV  : {tally['identical_text']}")
    log(f"  unreadable scans              : {tally['scans']}")
    log(f"  extraction errors             : {tally['failed']}")
    log(f"  ---- accounted {accounted} of {done} parsed"
        + ("" if accounted == done else "  <-- MISMATCH, tell me"))

    summary = tracker.summary()

    if summary:
        log(summary)

    return 0


if __name__ == "__main__":
    sys.exit(main())
