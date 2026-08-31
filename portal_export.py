"""
Convert parsed resume records into the exact column format that the
Guru Setu teacher portal's "Import teachers" screen expects.

The portal (src/services/teacherService.js -> normalizeImportedTeacher)
matches headers loosely but the SEPARATORS matter:

  * list fields  -> "; " separated        (Subjects, Grade Levels, Languages,
                                           Skills, College, Tags)
  * Work Experience -> one position per LINE, fields inside a line separated
                       by " | " in this order:
                       Institution | Role | Subject | From | To | Description
                       (a line with no "|" is read as just the institution)

Two record shapes are accepted:
  1. dicts straight out of extractor.extract_resume_data() — list fields are
     real Python lists
  2. rows read back from the Google master sheet — the same keys, but list
     fields are already ", " joined strings

Both are handled by _as_list().
"""

import pandas as pd


# Header order used in the generated file. Kept identical to the portal's
# downloadImportTemplate() so the two files look the same side by side.
PORTAL_COLUMNS = [
    "Full Name",
    "Email",
    "Phone",
    "Gender",
    "Age",
    "City",
    "State",
    "Subjects",
    "Grade Levels",
    "Languages",
    "Skills",
    "Qualification",
    "Extra Qualifications",
    "College Type",
    "College",
    "Experience (years)",
    "Current Institution",
    "Work Experience",
    "Current Salary",
    "Current Salary Unit",
    "Expected Salary",
    "Expected Salary Unit",
    "Preferred Job Type",
    "Availability",
    "Status",
    "CV URL",
    "Tags",
    "Notes",
    "Extra Col",
]


def _as_list(value, split_commas=True):
    """
    Normalize a field that may be a real list, a joined string, or None into
    a clean list of non-empty strings.

    split_commas=False matters for fields whose entries CONTAIN commas —
    "B.Sc Physics, Fergusson College, 2012" is one education entry, and
    "Delhi Public School, Noida" is one employer. Splitting those on commas
    shreds them into fragments, so those fields only ever split on ";".
    """

    if value is None:
        return []

    if isinstance(value, list):
        items = value

    else:
        text = str(value).strip()

        if not text or text.lower() in ("nan", "none"):
            return []

        # Sheet rows come back joined; hand-edited cells may use either mark.
        if split_commas:
            items = text.replace(";", ",").split(",")
        else:
            items = text.split(";")

    cleaned = []

    for item in items:

        if item is None:
            continue

        text = str(item).strip()

        if text and text.lower() not in ("nan", "none"):
            cleaned.append(text)

    return cleaned


def _join(value, separator="; ", split_commas=True):
    """List field -> the separator the portal importer splits on."""

    return separator.join(_as_list(value, split_commas=split_commas))


def _text(value):
    """Scalar field -> a clean string, or "" for missing values."""

    if value is None:
        return ""

    text = str(value).strip()

    if text.lower() in ("nan", "none"):
        return ""

    return text


def _number(value):
    """Age / experience -> an int for Excel, or "" so the cell stays blank."""

    if value is None or value == "":
        return ""

    try:
        return int(float(str(value).strip()))

    except (ValueError, TypeError):
        return ""


def _number_or_decimal(value):
    """
    Salary -> a number Excel stores as a number. Unlike _number() this keeps
    decimals, because an LPA figure is legitimately "4.5".
    """

    if value is None or value == "":
        return ""

    try:
        amount = float(str(value).strip())

    except (ValueError, TypeError):
        return ""

    if amount <= 0:
        return ""

    return int(amount) if amount == int(amount) else round(amount, 2)


def _salary_unit(amount, unit):
    """Only emit a unit when there is actually an amount to qualify."""

    if _number_or_decimal(amount) == "":
        return ""

    text = str(unit or "").strip().lower()

    return "lpa" if text == "lpa" else "pm"


def build_work_experience(record):
    """
    The portal has no current_designation or previous_institutions column —
    both belong in Work Experience.

    The current job becomes the first line (with "Present" as the To date so
    the portal shows it as ongoing); every previous institution follows on
    its own line. Designation is only known for the current role, so previous
    lines carry the institution name alone.
    """

    lines = []

    current_institution = _text(record.get("current_institution"))
    current_designation = _text(record.get("current_designation"))

    if current_institution:
        lines.append(
            f"{current_institution} | {current_designation} |  |  | Present"
        )

    elif current_designation:
        # Designation with no employer named — keep it rather than lose it.
        lines.append(f" | {current_designation} |  |  | Present")

    for institution in _as_list(
        record.get("previous_institutions"), split_commas=False
    ):

        if institution == current_institution:
            continue

        lines.append(institution)

    return "\n".join(lines)


def build_extra_col(record):
    """
    Everything the parser found that the portal has no field for. This lands
    in the teacher profile's "Extra info" box, so it stays searchable instead
    of being dropped on import.
    """

    parts = []

    education = _as_list(record.get("education_history"), split_commas=False)

    if education:
        parts.append("Education: " + "; ".join(education))

    resume_file = _text(record.get("resume_file_name"))

    if resume_file:
        parts.append(f"Resume file: {resume_file}")

    return " | ".join(parts)


def to_portal_row(record):
    """One parsed resume -> one dict keyed by the portal's column headers."""

    return {
        "Full Name": _text(record.get("full_name")),
        "Email": _text(record.get("email")),
        "Phone": _text(record.get("phone")),
        "Gender": _text(record.get("gender")),
        "Age": _number(record.get("age")),
        "City": _text(record.get("city")),
        "State": _text(record.get("state")),

        "Subjects": _join(record.get("subjects")),
        # COMMA, not semicolon: the portal's normalizeGradeLevels() splits this
        # column on "," only (utils/gradeLevels.js), unlike every other list
        # field. No canonical grade level contains a comma, so this is safe.
        "Grade Levels": _join(record.get("grade_levels"), ", "),
        "Languages": _join(record.get("languages")),
        "Skills": _join(record.get("skills")),

        "Qualification": _text(record.get("qualification")),
        # extraQualifications is a plain text field in the portal, not a list.
        "Extra Qualifications": _join(record.get("extra_qualifications"), ", "),

        "College Type": _text(record.get("college_type")),
        "College": _join(record.get("college"), split_commas=False),

        "Experience (years)": _number(record.get("experience_years")),
        "Current Institution": _text(record.get("current_institution")),
        "Work Experience": build_work_experience(record),

        # The portal reads the unit column to decide whether the amount is
        # monthly rupees or lakhs per annum, so a blank amount must leave the
        # unit blank too — never send a unit on its own.
        "Current Salary": _number_or_decimal(record.get("current_salary")),
        "Current Salary Unit": _salary_unit(
            record.get("current_salary"), record.get("current_salary_unit")
        ),
        "Expected Salary": _number_or_decimal(record.get("expected_salary")),
        "Expected Salary Unit": _salary_unit(
            record.get("expected_salary"), record.get("expected_salary_unit")
        ),

        "Preferred Job Type": _text(record.get("preferred_job_type")),
        "Availability": _text(record.get("availability")),
        "Status": "New",

        "CV URL": _text(record.get("resume_link")),
        # Grade levels / languages the extractor couldn't place. The portal
        # keeps unrecognised level text as tags too, so this matches it.
        "Tags": _join(record.get("tags")),
        "Notes": "",
        "Extra Col": build_extra_col(record),
    }


def to_portal_dataframe(records):
    """
    Build the import-ready DataFrame. Rows with no name, email AND phone are
    dropped because the portal's isImportRowUsable() would reject them anyway.
    """

    rows = []

    for record in records:

        row = to_portal_row(record)

        if not (row["Full Name"] or row["Email"] or row["Phone"]):
            continue

        rows.append(row)

    return pd.DataFrame(rows, columns=PORTAL_COLUMNS)


def drop_same_teacher(df):
    """
    Collapse rows the portal would treat as the same person, using the
    portal's own rule (utils/helpers.js isSameTeacher):

        same person  <=>  both have a name AND both have a phone
                          AND the names match AND the last 10 phone digits match

    A blank phone therefore never makes two rows duplicates. This matters:
    pandas' drop_duplicates treats blanks as equal, so deduping on
    ["Full Name", "Phone"] would silently delete distinct teachers whose
    number wasn't extracted — exactly the rows that most need a human look.
    """

    if df.empty:
        return df

    seen = set()
    keep = []

    for index, row in df.iterrows():

        raw_name = row.get("Full Name", "")
        name = "" if raw_name is None else str(raw_name).strip().lower()

        if name in ("none", "nan", "null"):
            name = ""

        digits = "".join(c for c in str(row.get("Phone", "")) if c.isdigit())
        mobile = digits[-10:] if len(digits) >= 10 else ""

        if not name or not mobile:
            # Can't be judged a duplicate — always kept.
            keep.append(index)
            continue

        key = (name, mobile)

        if key in seen:
            continue

        seen.add(key)
        keep.append(index)

    # Renumber so the caller gets a clean range index (st.data_editor
    # needs one, and a gappy index is confusing in any downstream use).
    return df.loc[keep].reset_index(drop=True)


def write_portal_excel(records, path="guru_setu_teacher_import.xlsx"):
    """Write the portal-format file and return (path, row_count)."""

    df = to_portal_dataframe(records)

    df.to_excel(path, index=False)

    return path, len(df)
