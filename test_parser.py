"""
Invariant tests for the parser.

Every bug in this project so far has been a silent data bug, not a crash:
rows multiplied by a label-based .loc, distinct candidates merged because
str(None) is the truthy string "none", "MBA" matched inside "Bombay",
blank rows counted as real. None of them raised an exception; all of them
were found by a human noticing a wrong number.

These tests assert the invariants those bugs violated.

    python test_parser.py
"""

import os
import sys

os.environ.setdefault("OPENAI_API_KEY", "test")
os.environ.setdefault("MOCK_MODE", "true")

import pandas as pd

from portal_export import (
    PORTAL_COLUMNS,
    to_portal_dataframe,
    drop_same_teacher,
)
from master_store import (
    MASTER_COLUMNS,
    empty_master,
    merge_into_master,
    to_master_rows,
    apply_edits,
    strip_tracking,
    cell,
)
from review import add_review_column, count_flagged, strip_review_column
from extractor import extract_qualifications
from duplicates import find_duplicate_groups, summarise
from skiplist import reason_for


FAILURES = []


def check(label, condition, detail=""):

    if condition:
        print(f"  PASS  {label}")

    else:
        print(f"  FAIL  {label} {detail}")
        FAILURES.append(label)


def record(name, phone, file_id, **extra):
    """A minimal parsed record."""

    base = {
        "full_name": name,
        "phone": phone,
        "file_id": file_id,
        "content_hash": f"hash-{file_id}",
        "resume_file_name": f"{file_id}.pdf",
        "subjects": ["Physics"],
    }
    base.update(extra)
    return base


# ---------------------------------------------------------------------------

def test_merge_never_invents_rows():
    """
    The bug: drop_blank_rows filtered with df.loc[labels]. After a merge the
    index has duplicate labels, so .loc returned each row once per matching
    label — 1814 + 11 became 1847.
    """

    master = pd.DataFrame([
        {**{c: "" for c in MASTER_COLUMNS},
         "Full Name": f"T{i}", "Phone": f"+9198{i:08d}",
         "Source File ID": f"O{i}"}
        for i in range(200)
    ])

    new = to_master_rows([
        record(f"New {i}", f"+9199{i:08d}", f"N{i}") for i in range(10)
    ])

    combined, added, _ = merge_into_master(master, new)

    check(
        "merge output never exceeds inputs",
        len(combined) <= len(master) + len(new),
        f"got {len(combined)}, max {len(master) + len(new)}",
    )
    check("merge count is exact", len(combined) == 210, f"got {len(combined)}")
    check("added is reported correctly", added == 10, f"got {added}")


def test_merge_is_idempotent():
    """Re-running the same folder must add nothing the second time."""

    rows = to_master_rows([record(f"T{i}", f"+9198{i:08d}", f"F{i}")
                           for i in range(5)])

    once, added_1, _ = merge_into_master(empty_master(), rows)
    twice, added_2, _ = merge_into_master(once, rows)

    check("first merge adds all", added_1 == 5, f"got {added_1}")
    check("second merge adds none", added_2 == 0, f"got {added_2}")
    check("row count stable", len(twice) == 5, f"got {len(twice)}")


def test_blank_fields_never_merge_distinct_people():
    """
    The bug: str(result.get("email", "")) returns "none" when the value is
    None, so every candidate without an email matched every other one.
    """

    rows = to_master_rows([
        record("Omprakash", "+918822114477", "F1", email=None),
        record("Sonukumar", "+919911772209", "F2", email=None),
        record("Yogesh", "+919928797360", "F3", email=None),
    ])

    combined, added, _ = merge_into_master(empty_master(), rows)

    check(
        "three candidates with no email all survive",
        len(combined) == 3,
        f"got {len(combined)}",
    )

    # Same name, different phone: different people.
    same_name = to_master_rows([
        record("Same Name", "+919811111111", "G1"),
        record("Same Name", "+919822222222", "G2"),
    ])

    combined2, _, _ = merge_into_master(empty_master(), same_name)

    check(
        "same name with different phones stays separate",
        len(combined2) == 2,
        f"got {len(combined2)}",
    )

    # Blank phone can never prove a match.
    blank_phone = to_master_rows([
        record("Blank One", "", "H1", email="a@x.com"),
        record("Blank One", "", "H2", email="b@x.com"),
    ])

    combined3, _, _ = merge_into_master(empty_master(), blank_phone)

    check(
        "blank phones are not treated as equal",
        len(combined3) == 2,
        f"got {len(combined3)}",
    )


def test_real_duplicates_are_merged():
    """The same person from two different files collapses to one row."""

    rows = to_master_rows([
        record("Gayatri Rukdikar", "+919823489490", "F1"),
        record("gayatri rukdikar", "9823489490", "F2"),
    ])

    combined, _, report = merge_into_master(empty_master(), rows)

    check("same person merged", len(combined) == 1, f"got {len(combined)}")
    check("the merge is explained", len(report) == 1, f"got {len(report)}")


def test_qualifications_use_word_boundaries():
    """The bug: substring matching found 'mba' inside 'Bombay'."""

    quals = extract_qualifications("M.Sc Physics from IIT Bombay, B.Ed")

    check("no phantom MBA from 'Bombay'", "MBA" not in quals, f"got {quals}")
    check("M.Sc found", "M.Sc" in quals, f"got {quals}")
    check("B.Ed found", "B.Ed" in quals, f"got {quals}")

    check(
        "spelled-out degrees recognised",
        "B.Tech" in extract_qualifications("Bachelor of Technology"),
    )

    ctet = extract_qualifications("CTET qualified")

    check("CTET does not also record TET", "TET" not in ctet, f"got {ctet}")

    check(
        "a school named Bombay is not a qualification",
        extract_qualifications("Teaches at Bombay Scottish School") == [],
    )


def test_portal_file_shape():
    """The download must always match the portal's import template."""

    df = to_portal_dataframe([record("A Teacher", "+919812345678", "F1")])

    check(
        "exactly the portal's columns",
        list(df.columns) == PORTAL_COLUMNS,
        f"got {len(df.columns)}",
    )

    # Grade Levels is the one column the portal splits on commas.
    rows = to_portal_dataframe([
        record("B", "+919812345679", "F2",
               grade_levels=["Senior Secondary (11-12)", "JEE Mains"])
    ])

    check(
        "grade levels are comma separated",
        "," in rows.iloc[0]["Grade Levels"],
        repr(rows.iloc[0]["Grade Levels"]),
    )

    check(
        "other list fields are semicolon separated",
        ";" in to_portal_dataframe([
            record("C", "+919812345670", "F3",
                   subjects=["Physics", "Chemistry"])
        ]).iloc[0]["Subjects"],
    )


def test_unusable_rows_never_stored():
    """A row with no name, email or phone is not a candidate."""

    junk = pd.DataFrame([
        {**{c: "" for c in MASTER_COLUMNS}, "Source File ID": f"J{i}"}
        for i in range(5)
    ])

    combined, added, _ = merge_into_master(empty_master(), junk)

    check("junk rows are not stored", len(combined) == 0, f"got {len(combined)}")
    check("junk adds nothing", added == 0, f"got {added}")

    check(
        "literal 'None' text counts as empty",
        cell({"x": "None"}, "x") == "" and cell({"x": None}, "x") == "",
    )


def test_edits_preserve_tracking():
    """Saving a correction must not break dedup."""

    master = to_master_rows([record("Meeta", "", "F1", email="m@x.com")])

    edited = strip_tracking(master, keep_id=True).copy()
    edited.loc[0, "Phone"] = "+916306209811"

    updated, changed, appended, _ = apply_edits(master, edited)

    check("the edit is applied", updated.iloc[0]["Phone"] == "+916306209811")
    check("one row changed", changed == 1, f"got {changed}")
    check("nothing appended", appended == 0, f"got {appended}")
    check(
        "tracking survives the edit",
        updated.iloc[0]["Source File ID"] == "F1"
        and updated.iloc[0]["Content Hash"] == "hash-F1",
    )


def test_review_column_round_trip():
    """The review column must never reach the downloaded file."""

    df = to_portal_dataframe([
        record("A", "+919812345678", "F1"),
        record("B", "", "F2", subjects=[]),
    ])

    flagged = add_review_column(df)

    check("review column is first", list(flagged.columns)[0] == "Needs review")
    check("incomplete rows are flagged", count_flagged(flagged) == 1,
          f"got {count_flagged(flagged)}")
    check(
        "review column is stripped for download",
        "Needs review" not in strip_review_column(flagged).columns,
    )
    check(
        "index is a clean range for the editor",
        list(flagged.index) == list(range(len(flagged))),
    )


def test_duplicate_summary_accounts_for_everything():
    """The bug: the summary reported two numbers that didn't add up."""

    files = (
        [{"id": f"a{i}", "name": f"a{i}.pdf", "md5Checksum": "AAA",
          "createdTime": "2026-01-01"} for i in range(3)]
        + [{"id": f"u{i}", "name": f"u{i}.pdf", "md5Checksum": f"U{i}",
            "createdTime": "2026-01-01"} for i in range(5)]
    )

    groups = find_duplicate_groups(files)
    extra, sets_, involved, appear_once = summarise(groups, len(files))

    check("extra copies counted", extra == 2, f"got {extra}")
    check(
        "every file is accounted for",
        involved + appear_once == len(files),
        f"{involved} + {appear_once} != {len(files)}",
    )


def test_keep_prefers_the_known_copy():
    """
    Trashing the copy recorded in the master would make the next run
    re-parse that resume and pay for it again.
    """

    files = [
        {"id": "A1", "name": "cv.pdf", "md5Checksum": "X",
         "createdTime": "2026-01-01"},
        {"id": "A2", "name": "cv (1).pdf", "md5Checksum": "X",
         "createdTime": "2026-02-01"},
    ]

    groups = find_duplicate_groups(files, known_ids={"A2"})

    check(
        "the copy already in the master is kept",
        groups[0]["keep"]["id"] == "A2",
        groups[0]["keep"]["id"],
    )


def test_only_real_failures_are_remembered():
    """Suppressing a duplicate or a same-person skip would lose candidates."""

    check("kept rows are not recorded", reason_for({"keep": True}) is None)
    check(
        "duplicates are not recorded",
        reason_for({"duplicate_of_content": True}) is None,
    )
    check(
        "same-person skips are not recorded",
        reason_for({"skipped_existing_contact": True}) is None,
    )
    check(
        "scans are recorded",
        reason_for({"needs_ocr": True, "extraction_failed": True}) is not None,
    )


def test_tracking_stays_with_the_right_candidate():
    """
    The bug: to_portal_dataframe treats the literal string "None" as blank
    and drops the row, but the tracking filter did not — so the two lists
    were different lengths and assigning file ids raised, killing the run.
    Earlier versions wrote that exact text, so such records really exist.
    """

    records = [
        record("None", "None", "F1", email="None", subjects=[]),
        record("Real Teacher", "+919812345678", "F2"),
        record("nan", "", "F3", email=None, subjects=[]),
    ]

    df = to_master_rows(records)

    check("unusable 'None' rows dropped", len(df) == 1, f"got {len(df)}")
    check(
        "the file id belongs to the surviving candidate",
        df.iloc[0]["Source File ID"] == "F2",
        df.iloc[0]["Source File ID"],
    )
    check(
        "the name belongs to the same candidate",
        df.iloc[0]["Full Name"] == "Real Teacher",
    )


def test_sanitizers_handle_real_resume_text():
    """Formats that actually appear in CVs, not just clean values."""

    from extractor import (sanitize_age, sanitize_phone,
                           sanitize_college_type)

    check("age written as '34 years'", sanitize_age("34 years") == 34)
    check("age written as 'Age: 34'", sanitize_age("Age: 34") == 34)
    check("a birth year is not an age", sanitize_age("1990") is None)
    check("an impossible age is rejected", sanitize_age("150") is None)

    check(
        "two phone numbers yields the first",
        sanitize_phone("+91-9812345678 / 9876543210") == "+919812345678",
    )
    check(
        "a phone buried in text is found",
        sanitize_phone("Mob: 9812345678 (R) 0212-2345678") == "+919812345678",
    )
    check(
        "a non-mobile is rejected",
        sanitize_phone("5812345678") is None,
    )

    check("IIT recognised", sanitize_college_type("IIT Bombay") == "IIT")
    check(
        "IIIT is not an IIT",
        sanitize_college_type("IIIT Hyderabad") is None,
    )
    check(
        "MANIT is not an NIT",
        sanitize_college_type("MANIT Bhopal") is None,
    )


def test_salary_agrees_with_its_unit():
    """
    The portal reads 'lpa' as a figure IN LAKHS, so 450000 lpa would be
    about 45 billion rupees landing in the CRM.
    """

    from extractor import (sanitize_salary, sanitize_salary_unit,
                           reconcile_salary)

    def final(amount, unit):
        a = sanitize_salary(amount)
        u = sanitize_salary_unit(unit, a)
        return reconcile_salary(a, u)

    check("annual rupees convert to lakhs", final("450000", "per annum") == (4.5, "lpa"))
    check("540000/annum becomes 5.4 lpa", final("540000", "annum") == (5.4, "lpa"))
    check("lakhs stay as lakhs", final("4.5", "lpa") == (4.5, "lpa"))
    check("monthly stays monthly", final("35000", "per month") == (35000, "pm"))
    check("a tiny 'pm' figure is really lpa", final("6", "pm") == (6, "lpa"))
    check("zero is rejected", final("0", None) == (None, None))


def test_vocabulary_has_no_false_positives():
    """Short aliases must not match inside unrelated words."""

    from portal_vocab import (snap_subjects, snap_grade_levels,
                              snap_qualification)

    for text in ["Interior Design", "Peace Studies",
                 "Artificial Intelligence", "Bioinformatics"]:
        canonical, _ = snap_subjects([text])
        check(f"{text!r} is not a subject", canonical == [], f"got {canonical}")

    for text in ["Internet Safety", "Board of Directors", "Magnet School"]:
        canonical, _ = snap_grade_levels([text])
        check(f"{text!r} is not a grade level", canonical == [],
              f"got {canonical}")

    # And the real ones still work.
    check("IT maps to Information Technology",
          snap_subjects(["IT"])[0] == ["Information Technology"])
    check("board exams still match",
          snap_grade_levels(["CBSE board exams"])[0] == ["Board Exams"])
    check("spelled-out degrees map",
          snap_qualification("Bachelor of Arts") == "B.A")


def test_grid_deletions_are_scoped_to_what_was_visible():
    """
    Deleting rows in the grid did nothing at all, despite the editor
    allowing it. Implementing it naively is worse: the run panel and any
    search or filter show a SUBSET of the master, so treating absent rows
    as deleted would wipe everything the user simply wasn't looking at.
    """

    master = to_master_rows([
        record(f"T{i}", f"+9198{i:08d}", f"F{i}") for i in range(5)
    ])

    visible = strip_tracking(master, keep_id=True)

    # Full view, two rows deleted.
    updated, _, _, deleted = apply_edits(
        master, visible.iloc[[0, 2, 4]].copy(), visible
    )

    check("deletions are applied", deleted == 2, f"got {deleted}")
    check("master shrinks correctly", len(updated) == 3, f"got {len(updated)}")

    # Filtered view showing 2 of 5, nothing deleted by the user.
    filtered = visible.iloc[[0, 1]]

    updated2, _, _, deleted2 = apply_edits(master, filtered.copy(), filtered)

    check(
        "rows outside a filtered view are NOT deleted",
        deleted2 == 0 and len(updated2) == 5,
        f"deleted {deleted2}, left {len(updated2)}",
    )

    # Deleting within a filtered view removes only that row.
    updated3, _, _, deleted3 = apply_edits(
        master, filtered.iloc[[0]].copy(), filtered
    )

    check(
        "deleting inside a filter removes only that row",
        deleted3 == 1 and len(updated3) == 4,
        f"deleted {deleted3}, left {len(updated3)}",
    )


def test_none_never_becomes_the_text_none():
    """
    str(None) is "None", which is truthy. That single mistake has caused
    three separate bugs: distinct candidates merged, a run-killing row
    misalignment, and the literal word written into stored fields.
    """

    from extractor import flatten_value, flatten_list_field

    check("None flattens to None", flatten_value(None) is None)
    check("the text 'None' flattens to None", flatten_value("None") is None)
    check("the text 'nan' flattens to None", flatten_value("nan") is None)
    check("real values survive", flatten_value("M.Sc") == "M.Sc")

    cleaned = flatten_list_field([{"degree": "M.Sc"}, "B.Ed", None, "", "nan"])

    check(
        "list fields carry no junk",
        cleaned == ["degree: M.Sc", "B.Ed"],
        f"got {cleaned}",
    )


def test_folder_links_are_parsed_robustly():
    """A trailing slash yielded '1ABC/', and the run reported an empty folder."""

    from drive_utils import extract_folder_id

    for url in [
        "https://drive.google.com/drive/folders/1ABC",
        "https://drive.google.com/drive/folders/1ABC/",
        "https://drive.google.com/drive/folders/1ABC?usp=sharing",
        "https://drive.google.com/drive/u/0/folders/1ABC/",
        "https://drive.google.com/drive/folders/1ABC#gid=0",
    ]:
        check(f"parsed {url[-24:]!r}", extract_folder_id(url) == "1ABC",
              extract_folder_id(url))


def test_same_teacher_matches_the_portal():
    """drop_same_teacher must agree with the portal's isSameTeacher."""

    df = pd.DataFrame([
        {"Full Name": "A", "Phone": "+919812345678"},
        {"Full Name": "a", "Phone": "9812345678"},
        {"Full Name": "B", "Phone": ""},
        {"Full Name": "B", "Phone": ""},
        {"Full Name": "", "Phone": "+919812345678"},
    ])

    kept = drop_same_teacher(df)

    check(
        "only the true duplicate is dropped",
        len(kept) == 4,
        f"got {len(kept)}",
    )


def main():

    tests = [
        test_merge_never_invents_rows,
        test_merge_is_idempotent,
        test_blank_fields_never_merge_distinct_people,
        test_real_duplicates_are_merged,
        test_qualifications_use_word_boundaries,
        test_portal_file_shape,
        test_unusable_rows_never_stored,
        test_edits_preserve_tracking,
        test_review_column_round_trip,
        test_duplicate_summary_accounts_for_everything,
        test_keep_prefers_the_known_copy,
        test_only_real_failures_are_remembered,
        test_same_teacher_matches_the_portal,
        test_tracking_stays_with_the_right_candidate,
        test_sanitizers_handle_real_resume_text,
        test_salary_agrees_with_its_unit,
        test_vocabulary_has_no_false_positives,
        test_grid_deletions_are_scoped_to_what_was_visible,
        test_none_never_becomes_the_text_none,
        test_folder_links_are_parsed_robustly,
    ]

    for test in tests:
        print(f"\n{test.__name__}")
        test()

    print("\n" + "=" * 60)

    if FAILURES:
        print(f"{len(FAILURES)} FAILING CHECK(S):")

        for name in FAILURES:
            print(f"  - {name}")

        return 1

    print("All checks passed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
