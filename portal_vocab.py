"""
The portal's fixed option lists, mirrored from the portal repo at
src/data/options.js.

Keeping these here lets the extractor ask the LLM for values that are ALREADY
in the portal's vocabulary, instead of emitting free text and hoping the
portal's importer can map it afterwards.

Two layers of safety:
  1. the prompt (extractor.py) lists the allowed values explicitly
  2. snap_* below re-checks every value the model returned, maps common
     drift via the SYNONYM tables, and hands back anything it can't place
     as "leftover" so the caller can keep it as a tag rather than lose it

The portal's own normalizer still runs at import time, so this is a first
pass, not the last line of defence.

IF THE ADMIN EDITS A LIST IN Settings -> Lists, UPDATE THE MATCHING LIST HERE.
"""


GRADE_LEVELS = [
    "Pre-Primary", "Primary (1-5)", "Middle (6-8)", "Secondary (9-10)",
    "Senior Secondary (11-12)", "Dropper / Repeater", "College / UG",
    "Post Graduate", "Foundation (6-10)", "JEE Mains", "JEE Advanced",
    "NEET", "State CET", "CUET", "Olympiad / NTSE", "Defence (NDA/Sainik)",
    "GATE / NET", "Board Exams", "Other Competitive Exams",
]

SUBJECTS = [
    "Mathematics", "Physics", "Chemistry",
    "Inorganic Chemistry", "Organic Chemistry", "Physical Chemistry",
    "Biology", "Zoology", "Botany", "Science (General)",
    "English", "Hindi", "Sanskrit", "French", "Social Studies", "History",
    "Geography", "Economics", "Political Science", "Commerce", "Accountancy",
    "Business Studies", "Computer Science", "Information Technology",
    "Physical Education", "Art & Craft", "Music", "Dance",
    "Environmental Studies", "Pre-Primary / Montessori", "Special Education",
]

LANGUAGES = [
    "English", "Hindi", "Marathi", "Gujarati", "Punjabi", "Urdu", "Bengali",
    "Tamil", "Telugu", "Kannada", "Malayalam", "Odia", "Assamese", "Sanskrit",
    "French", "German", "Spanish",
]

QUALIFICATIONS = [
    "B.Ed", "M.Ed", "D.El.Ed", "NTT", "B.A", "M.A", "B.Sc", "M.Sc",
    "B.Com", "M.Com", "B.Tech", "M.Tech", "BCA", "MCA", "PhD",
    "CTET Qualified", "State TET Qualified", "Diploma", "Other",
]

JOB_TYPES = [
    "Full-time", "Part-time", "Substitute", "Visiting Faculty",
    "Online / Remote", "Home Tuition",
]

AVAILABILITY = [
    "Immediate", "Within 15 days", "Within 1 month",
    "Within 2 months", "Notice period", "Not actively looking",
]

GENDERS = ["Female", "Male", "Other", "Prefer not to say"]

TEACHER_STATUSES = [
    "New", "Active", "Shortlisted", "Placed", "On Hold", "Archived",
]

SKILLS = [
    "Classroom Management", "Lesson Planning", "Curriculum Development",
    "Smart Board / EdTech", "Online Teaching", "Student Counselling",
    "Exam Preparation", "Special Needs Education", "Extracurricular Activities",
    "Communication", "Leadership", "Assessment & Evaluation",
]

COLLEGE_TYPES = ["IIT", "NIT", "Other"]


# ---------------------------------------------------------------------------
#  Synonym tables — only for drift the model actually produces. The prompt
#  already asks for canonical values, so these stay deliberately small.
# ---------------------------------------------------------------------------

GRADE_LEVEL_SYNONYMS = {
    "pre primary": "Pre-Primary",
    "preprimary": "Pre-Primary",
    "nursery": "Pre-Primary",
    "kg": "Pre-Primary",
    "montessori": "Pre-Primary",
    "primary": "Primary (1-5)",
    "class 1-5": "Primary (1-5)",
    "middle": "Middle (6-8)",
    "class 6-8": "Middle (6-8)",
    "upper primary": "Middle (6-8)",
    "secondary": "Secondary (9-10)",
    "high school": "Secondary (9-10)",
    "class 9-10": "Secondary (9-10)",
    "tgt": "Secondary (9-10)",
    "senior secondary": "Senior Secondary (11-12)",
    "higher secondary": "Senior Secondary (11-12)",
    "class 11-12": "Senior Secondary (11-12)",
    "intermediate": "Senior Secondary (11-12)",
    "hsc": "Senior Secondary (11-12)",
    "puc": "Senior Secondary (11-12)",
    "pgt": "Senior Secondary (11-12)",
    "10+2": "Senior Secondary (11-12)",
    "dropper": "Dropper / Repeater",
    "repeater": "Dropper / Repeater",
    "college": "College / UG",
    "ug": "College / UG",
    "undergraduate": "College / UG",
    "graduation": "College / UG",
    "pg": "Post Graduate",
    "postgraduate": "Post Graduate",
    "post graduate": "Post Graduate",
    "foundation": "Foundation (6-10)",
    "jee main": "JEE Mains",
    "jee mains": "JEE Mains",
    "jee advance": "JEE Advanced",
    "jee advanced": "JEE Advanced",
    "iit jee": ["JEE Mains", "JEE Advanced"],
    "iit-jee": ["JEE Mains", "JEE Advanced"],
    "engineering entrance": ["JEE Mains", "JEE Advanced"],
    "neet": "NEET",
    "aipmt": "NEET",
    "medical entrance": "NEET",
    "cet": "State CET",
    "mht cet": "State CET",
    "mhcet": "State CET",
    "kcet": "State CET",
    "cuet": "CUET",
    "olympiad": "Olympiad / NTSE",
    "ntse": "Olympiad / NTSE",
    "nda": "Defence (NDA/Sainik)",
    "sainik": "Defence (NDA/Sainik)",
    "defence": "Defence (NDA/Sainik)",
    "gate": "GATE / NET",
    "net": "GATE / NET",
    "ugc net": "GATE / NET",
    "csir net": "GATE / NET",
    # NOT a bare "board": it matched "Board of Directors" and "notice
    # board". The exam sense always carries a qualifying word.
    "board exam": "Board Exams",
    "board exams": "Board Exams",
    "boards": "Board Exams",
    "board pattern": "Board Exams",
    "cbse": "Board Exams",
    "icse": "Board Exams",
    "state board": "Board Exams",
    "competitive exams": "Other Competitive Exams",
    "competitive": "Other Competitive Exams",
}

SUBJECT_SYNONYMS = {
    "math": "Mathematics", "maths": "Mathematics", "mathematic": "Mathematics",
    "algebra": "Mathematics", "calculus": "Mathematics", "geometry": "Mathematics",
    "trigonometry": "Mathematics", "statistics": "Mathematics",
    "phy": "Physics", "physic": "Physics", "mechanics": "Physics",
    "optics": "Physics", "thermodynamics": "Physics",
    "chem": "Chemistry", "biochemistry": "Chemistry",
    "bio": "Biology", "life science": "Biology", "genetics": "Biology",
    "microbiology": "Biology", "biotechnology": "Biology",
    "science": "Science (General)", "general science": "Science (General)",
    "sst": "Social Studies", "social science": "Social Studies",
    "civics": "Political Science", "polity": "Political Science",
    "evs": "Environmental Studies", "environmental science": "Environmental Studies",
    "computer": "Computer Science", "computers": "Computer Science",
    "cs": "Computer Science", "programming": "Computer Science",
    "it": "Information Technology",
    "pe": "Physical Education", "sports": "Physical Education",
    "accounts": "Accountancy", "accounting": "Accountancy",
    "bst": "Business Studies",
    "eco": "Economics",
    "montessori": "Pre-Primary / Montessori",
    "pre primary": "Pre-Primary / Montessori",
    "special needs": "Special Education",
    "art": "Art & Craft", "craft": "Art & Craft", "drawing": "Art & Craft",
}

QUALIFICATION_SYNONYMS = {
    "ph.d": "PhD", "ph. d": "PhD", "phd": "PhD",
    "doctor of philosophy": "PhD", "doctorate": "PhD",
    "ctet": "CTET Qualified",
    "tet": "State TET Qualified",
    "state tet": "State TET Qualified",
    "d.t.ed": "Diploma", "d.ed": "Diploma", "dted": "Diploma",
    "m.sc": "M.Sc", "msc": "M.Sc",
    "b.sc": "B.Sc", "bsc": "B.Sc",
    "m.a": "M.A", "ma": "M.A",
    "b.a": "B.A", "ba": "B.A",
    "m.com": "M.Com", "b.com": "B.Com",
    "m.ed": "M.Ed", "b.ed": "B.Ed",
    "m.tech": "M.Tech", "b.tech": "B.Tech",
    "mca": "MCA", "bca": "BCA",
    "d.el.ed": "D.El.Ed", "deled": "D.El.Ed",
    "ntt": "NTT",

    "bachelor of arts": "B.A",
    "bachelor of science": "B.Sc",
    "bachelor of commerce": "B.Com",
    "bachelor of education": "B.Ed",
    "bachelor of technology": "B.Tech",
    "bachelor of engineering": "B.Tech",
    "master of arts": "M.A",
    "master of science": "M.Sc",
    "master of commerce": "M.Com",
    "master of education": "M.Ed",
    "master of technology": "M.Tech",
    "master of computer applications": "MCA",
    "master of business administration": "MBA",
}

JOB_TYPE_SYNONYMS = {
    "full time": "Full-time", "fulltime": "Full-time",
    "permanent": "Full-time", "regular": "Full-time",
    "part time": "Part-time", "parttime": "Part-time",
    "substitute": "Substitute", "temporary": "Substitute",
    "visiting": "Visiting Faculty", "visiting faculty": "Visiting Faculty",
    "guest faculty": "Visiting Faculty", "contract": "Visiting Faculty",
    "online": "Online / Remote", "remote": "Online / Remote",
    "work from home": "Online / Remote",
    "home tuition": "Home Tuition", "tuition": "Home Tuition",
    "private tutor": "Home Tuition",
}

AVAILABILITY_SYNONYMS = {
    "immediate": "Immediate",
    "immediately": "Immediate",
    "available immediately": "Immediate",
    "15 days": "Within 15 days",
    "within 15 days": "Within 15 days",
    "2 weeks": "Within 15 days",
    "1 month": "Within 1 month",
    "30 days": "Within 1 month",
    "one month": "Within 1 month",
    "2 months": "Within 2 months",
    "60 days": "Within 2 months",
    "two months": "Within 2 months",
    "3 months": "Notice period",
    "90 days": "Notice period",
    "notice period": "Notice period",
    "serving notice": "Notice period",
    "not looking": "Not actively looking",
    "not actively looking": "Not actively looking",
}

GENDER_SYNONYMS = {
    "m": "Male", "male": "Male", "man": "Male",
    "f": "Female", "female": "Female", "woman": "Female",
    "other": "Other", "transgender": "Other",
}


def _key(value):
    """Lowercase and strip punctuation so 'Ph.D.' and 'phd' compare equal."""

    text = str(value or "").strip().lower()

    # Collapse everything that isn't a letter or digit into single spaces.
    cleaned = []
    previous_was_space = False

    for char in text:

        if char.isalnum():
            cleaned.append(char)
            previous_was_space = False

        elif not previous_was_space:
            cleaned.append(" ")
            previous_was_space = True

    return "".join(cleaned).strip()


def _build_index(canonical_list, synonyms):
    """
    Map every accepted spelling onto its canonical value (or values — a bare
    "IIT-JEE" legitimately means both JEE papers).

    Returns (exact_index, ordered_aliases). ordered_aliases is sorted longest
    first so the containment fallback prefers "jee advanced" over "jee".
    """

    index = {}

    for value in canonical_list:
        index[_key(value)] = [value]

    for alias, value in (synonyms or {}).items():
        targets = value if isinstance(value, list) else [value]
        index.setdefault(_key(alias), targets)

    ordered = sorted(index.keys(), key=len, reverse=True)

    return index, ordered


def _lookup(text, index, ordered, allow_multiple):
    """
    Exact match first, then a word-boundary containment scan so free-text
    phrasings still resolve. Padding both sides with spaces keeps short
    aliases like "ma" from matching inside "mathematics".
    """

    key = _key(text)

    if not key:
        return []

    if key in index:
        return list(index[key])

    padded = f" {key} "
    hits = []

    for alias in ordered:

        if f" {alias} " in padded:

            for target in index[alias]:

                if target not in hits:
                    hits.append(target)

            if not allow_multiple:
                break

    return hits


_GRADE_INDEX = _build_index(GRADE_LEVELS, GRADE_LEVEL_SYNONYMS)
_SUBJECT_INDEX = _build_index(SUBJECTS, SUBJECT_SYNONYMS)
_LANGUAGE_INDEX = _build_index(LANGUAGES, {})
_QUALIFICATION_INDEX = _build_index(QUALIFICATIONS, QUALIFICATION_SYNONYMS)
_JOB_TYPE_INDEX = _build_index(JOB_TYPES, JOB_TYPE_SYNONYMS)
_AVAILABILITY_INDEX = _build_index(AVAILABILITY, AVAILABILITY_SYNONYMS)
_GENDER_INDEX = _build_index(GENDERS, GENDER_SYNONYMS)


def _snap_list(values, built, canonical_order):
    """
    Returns (canonical_values, leftover). Canonical values come back in the
    portal's own list order so the output is stable regardless of the order
    the model happened to emit them in. One raw value may yield several
    canonical ones ("IIT-JEE" -> both JEE papers).
    """

    index, ordered_aliases = built

    if not values:
        return [], []

    if not isinstance(values, list):
        values = [values]

    matched = set()
    leftover = []

    for raw in values:

        text = str(raw or "").strip()

        if not text:
            continue

        hits = _lookup(text, index, ordered_aliases, allow_multiple=True)

        if hits:
            matched.update(hits)
        elif text not in leftover:
            leftover.append(text)

    ordered = [v for v in canonical_order if v in matched]

    return ordered, leftover


def _snap_one(value, built, keep_unknown=False):
    """
    Single-value snap. keep_unknown=True returns the original text when there
    is no canonical equivalent — right for qualification, where "MBA" or
    "UGC NET" is still useful free text, but wrong for a strict dropdown
    like job type, where an unknown value would never match a filter.
    """

    index, ordered_aliases = built

    text = str(value or "").strip()

    if not text:
        return None

    hits = _lookup(text, index, ordered_aliases, allow_multiple=False)

    if hits:
        return hits[0]

    return text if keep_unknown else None


def snap_grade_levels(values):
    return _snap_list(values, _GRADE_INDEX, GRADE_LEVELS)


def snap_subjects(values):
    return _snap_list(values, _SUBJECT_INDEX, SUBJECTS)


def snap_languages(values):
    return _snap_list(values, _LANGUAGE_INDEX, LANGUAGES)


def snap_qualification(value):
    return _snap_one(value, _QUALIFICATION_INDEX, keep_unknown=True)


def snap_job_type(value):
    return _snap_one(value, _JOB_TYPE_INDEX)


def snap_availability(value):
    return _snap_one(value, _AVAILABILITY_INDEX)


def snap_gender(value):
    return _snap_one(value, _GENDER_INDEX)


def format_allowed(values):
    """Render a list for the prompt: 'A' | 'B' | 'C'."""

    return " | ".join(f'"{v}"' for v in values)
