"""
Flag parsed rows that a human should look at before they reach the portal.

Extraction is never perfect, and a bad row is much cheaper to fix here than
after it has been imported and merged into the CRM. Three problems make a row
genuinely unusable rather than merely incomplete:

  * no name   — nothing to call the person
  * no phone  — the consultancy can't contact them, and the portal's
                duplicate rule (isSameTeacher) needs a phone to work at all,
                so phone-less rows quietly become permanent duplicates
  * no subjects — they can't be matched to any vacancy

Everything else (missing age, no salary, no college) is normal for a resume
and is not worth interrupting the user over.
"""

REVIEW_COLUMN = "Needs review"


def review_issues(row):
    """Return the list of problems with one portal-format row."""

    def blank(column):
        value = row.get(column, "")
        return str(value).strip() == "" or str(value).strip().lower() == "nan"

    issues = []

    if blank("Full Name"):
        issues.append("no name")

    if blank("Phone"):
        issues.append("no phone")

    if blank("Subjects"):
        issues.append("no subjects")

    return issues


def review_label(row):
    """The cell text shown in the review column ("" when the row is clean)."""

    issues = review_issues(row)

    return ", ".join(issues) if issues else ""


def add_review_column(df):
    """
    Add the review column, put it first, and sort the rows that need
    attention to the top (most problems first). Row order within each group
    is preserved so a run stays reproducible.
    """

    if df.empty:
        df[REVIEW_COLUMN] = []
        return df

    working = df.copy()

    working[REVIEW_COLUMN] = [
        review_label(row) for _, row in working.iterrows()
    ]

    # Sort key: more problems first, clean rows last. mergesort is stable, so
    # rows with the same number of problems keep their original order.
    problem_count = [
        len(review_issues(row)) for _, row in working.iterrows()
    ]

    working["_problems"] = problem_count

    working = working.sort_values(
        "_problems",
        ascending=False,
        kind="mergesort"
    ).drop(columns=["_problems"])

    columns = [REVIEW_COLUMN] + [
        c for c in working.columns if c != REVIEW_COLUMN
    ]

    # Sorting leaves the original index behind. st.data_editor needs a plain
    # range index to let the user add rows, so renumber before handing it over.
    return working[columns].reset_index(drop=True)


def count_flagged(df):
    """How many rows currently have at least one problem."""

    if df.empty:
        return 0

    return sum(1 for _, row in df.iterrows() if review_issues(row))


def strip_review_column(df):
    """
    Remove the review column before writing the file. The portal ignores
    unknown headers, so leaving it in would be harmless — but the download
    should match the import template exactly, with no stray columns.
    """

    if REVIEW_COLUMN in df.columns:
        return df.drop(columns=[REVIEW_COLUMN])

    return df
