"""
A job queue shared between the web app and the background worker.

WHY
---
Streamlit ties a running script to the browser session, so a long parse
dies when a laptop sleeps. The command line avoids that but needs Python,
a terminal and credential files — which staff won't have.

So the two are split. Anyone can queue a folder from the app in one click;
a worker running somewhere that stays awake picks the job up and parses it.
The browser can be closed the moment the job is queued.

The queue is a small JSON file in the Drive folder, for the same reason the
master workbook lives there: Streamlit Community Cloud gives the container
no persistent disk, and Drive is already authenticated.

A JOB'S LIFE
------------
    queued -> running -> done
                      -> failed

`claimed_at` is what stops two workers doing the same job: a worker only
takes a job that is queued, or one that is "running" but whose heartbeat
has gone stale (the worker died mid-job).
"""

import json
import time
import uuid

from drive_utils import download_cache_from_drive, upload_cache_to_drive


JOBS_FILE_NAME = "_parser_jobs.json"

# A running job whose heartbeat is older than this is considered abandoned
# and may be retried. Comfortably longer than one batch takes.
STALE_AFTER_SECONDS = 15 * 60


def _now():
    return int(time.time())


def load_jobs(folder_id):
    """All jobs, oldest first. Empty on any read problem."""

    if not folder_id:
        return []

    try:
        raw = download_cache_from_drive(folder_id, JOBS_FILE_NAME)

    except Exception:
        return []

    if not raw:
        return []

    try:
        data = json.loads(raw)

    except (ValueError, TypeError):
        return []

    return data if isinstance(data, list) else []


class QueueError(Exception):
    """A queue write failed, carrying Drive's own reason."""


def save_jobs(folder_id, jobs):
    """
    Write the queue back.

    Raises QueueError with Drive's message rather than returning a bare
    False. "Could not write to the queue file" told the user nothing —
    the actual cause is almost always a permission that can be fixed in
    ten seconds once it's named.
    """

    if not folder_id:
        raise QueueError("No Drive folder to write the queue to.")

    try:
        upload_cache_to_drive(
            folder_id, JOBS_FILE_NAME, json.dumps(jobs, indent=1)
        )
        return True

    except Exception as e:

        message = str(e)

        if "insufficientFilePermissions" in message or "403" in message:
            raise QueueError(
                "The service account can only READ this folder, so it "
                "can't create the queue file. In Drive: right-click the "
                "folder -> Share -> change the service account from "
                "Viewer to Editor."
            ) from e

        if "404" in message or "notFound" in message:
            raise QueueError(
                "Drive can't find that folder. Check the link, and that "
                "the folder is shared with the service account."
            ) from e

        raise QueueError(f"Drive refused the write: {message}") from e


def add_job(folder_id, folder_link, master_link="", requested_by=""):
    """
    Queue a folder. Returns (job, message).

    A folder already queued or running is not added twice — clicking the
    button again while a job is in flight should be harmless.
    """

    jobs = load_jobs(folder_id)

    for job in jobs:

        if (
            job.get("folder_link") == folder_link
            and job.get("status") in ("queued", "running")
        ):
            return job, "That folder is already queued."

    job = {
        "id": uuid.uuid4().hex[:12],
        "folder_link": folder_link,
        "master_link": master_link,
        "requested_by": requested_by,
        "status": "queued",
        "created_at": _now(),
        "claimed_at": None,
        "finished_at": None,
        "progress": "",
        "error": "",
    }

    jobs.append(job)

    try:
        save_jobs(folder_id, jobs)

    except QueueError as e:
        return None, str(e)

    return job, "Queued."


def claim_next_job(folder_id, worker_name=""):
    """
    Take the oldest job that needs doing, or None.

    Claiming writes the queue back immediately, so a second worker sees
    the job as taken. Stale "running" jobs are reclaimable because a
    worker killed mid-job would otherwise block the queue forever.
    """

    jobs = load_jobs(folder_id)

    for job in jobs:

        status = job.get("status")

        claimable = status == "queued" or (
            status == "running"
            and _now() - (job.get("claimed_at") or 0) > STALE_AFTER_SECONDS
        )

        if not claimable:
            continue

        job["status"] = "running"
        job["claimed_at"] = _now()
        job["worker"] = worker_name
        job["error"] = ""

        try:
            save_jobs(folder_id, jobs)
            return job

        except QueueError:
            # Someone else may have claimed it; try again next poll.
            return None

    return None


def update_job(folder_id, job_id, **fields):
    """
    Patch a job. Also refreshes claimed_at, which doubles as a heartbeat —
    a job reporting progress is alive and must not be reclaimed.
    """

    jobs = load_jobs(folder_id)

    for job in jobs:

        if job.get("id") == job_id:

            job.update(fields)

            if job.get("status") == "running":
                job["claimed_at"] = _now()

            if fields.get("status") in ("done", "failed"):
                job["finished_at"] = _now()

            try:
                return save_jobs(folder_id, jobs)

            except QueueError:
                # A progress update is not worth killing a running job for.
                return False

    return False


def describe(job):
    """One readable line for the app's status table."""

    status = job.get("status", "?")

    when = job.get("finished_at") or job.get("claimed_at") or job.get("created_at")

    stamp = time.strftime("%d %b %H:%M", time.localtime(when)) if when else ""

    return {
        "Status": status,
        "Folder": (job.get("folder_link") or "")[-28:],
        "Progress": job.get("progress") or "",
        "Updated": stamp,
        "Error": (job.get("error") or "")[:80],
    }
