"""
Find byte-identical duplicate files in a Drive folder.

Drive computes an md5Checksum for every uploaded file, so "cv.pdf" and
"cv (1).pdf" can be matched without downloading either one — no bandwidth,
no OpenAI tokens, no OCR.

Choosing which copy to KEEP matters more than finding the duplicates:

  1. a copy already recorded in the master database wins, because its file
     id is what the master uses to skip re-parsing. Trashing that one and
     keeping an unknown copy would make the next run re-parse the resume
     and pay for it again.
  2. otherwise the cleanest filename wins — "cv.pdf" over "cv (1).pdf" —
     judged by absence of a copy-suffix, then by length, then oldest.
"""

import re


# "cv (1).pdf", "cv - Copy.pdf", "cv(2).pdf", "copy of cv.pdf"
COPY_MARKERS = [
    re.compile(r"\(\s*\d+\s*\)"),
    re.compile(r"\bcopy\b", re.IGNORECASE),
    re.compile(r"\bcopie\b", re.IGNORECASE),
]


def looks_like_a_copy(name):
    """True if the filename carries a copy marker."""

    return any(marker.search(str(name or "")) for marker in COPY_MARKERS)


def _keep_rank(file_info, known_ids):
    """
    Sort key — lower is better, so the first entry is the one to keep.

    Tuple order encodes the priority described in the module docstring.
    """

    name = str(file_info.get("name") or "")

    return (
        # Already in the master: keep it, or the next run re-parses.
        0 if file_info.get("id") in known_ids else 1,
        # A name with "(1)" or "Copy" is the copy, not the original.
        1 if looks_like_a_copy(name) else 0,
        # Shorter name usually means the original.
        len(name),
        # Oldest wins any remaining tie.
        str(file_info.get("createdTime") or ""),
    )


def find_duplicate_groups(files, known_ids=None):
    """
    Group files by checksum.

    Returns a list of {checksum, keep, duplicates} for every set with more
    than one copy. Files with no checksum (Drive doesn't compute one for
    native Google formats) are ignored rather than guessed at.
    """

    known_ids = known_ids or set()

    by_checksum = {}

    for file_info in files:

        checksum = file_info.get("md5Checksum")

        if not checksum:
            continue

        by_checksum.setdefault(checksum, []).append(file_info)

    groups = []

    for checksum, matches in by_checksum.items():

        if len(matches) < 2:
            continue

        ordered = sorted(matches, key=lambda f: _keep_rank(f, known_ids))

        groups.append({
            "checksum": checksum,
            "keep": ordered[0],
            "duplicates": ordered[1:],
        })

    # Biggest pile-ups first — they're the ones worth acting on.
    groups.sort(key=lambda g: len(g["duplicates"]), reverse=True)

    return groups


def summarise(groups):
    """(number of duplicate files, number of groups) for the UI."""

    return sum(len(g["duplicates"]) for g in groups), len(groups)
