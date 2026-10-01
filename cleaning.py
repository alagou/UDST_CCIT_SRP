"""Pure cleaning functions for Microsoft Forms free-text fields.

Each function returns ``(cleaned_value, warning_or_None)``. A warning means
the value was odd and should be fixed at the source. The cleaned value is
still returned so nothing is silently discarded.
"""

from __future__ import annotations

import math
import re

_EMAIL_RE = re.compile(r"^[a-z0-9._%+\-]+@[a-z0-9.\-]+\.[a-z]{2,}$")
_YEAR_WORDS = {
    "1": "Year 1",
    "first": "Year 1",
    "1st": "Year 1",
    "2": "Year 2",
    "second": "Year 2",
    "2nd": "Year 2",
    "3": "Year 3",
    "third": "Year 3",
    "3rd": "Year 3",
    "4": "Year 4",
    "fourth": "Year 4",
    "4th": "Year 4",
}
_YES = {"y", "yes", "true", "t", "1", "confirm", "confirmed"}
_NO = {"n", "no", "false", "f", "0"}
_BLANK_TOKENS = {"", "nan", "none", "null", "nat"}


def _is_nan(value) -> bool:
    if value is None:
        return True
    if isinstance(value, float) and math.isnan(value):
        return True
    return False


def clean_text(value):
    """Strip text, collapse whitespace, and turn blanks into None."""
    if _is_nan(value):
        return None, None
    text = str(value).replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text).strip()
    if text.lower() in _BLANK_TOKENS:
        return None, None
    return text, None


def clean_multiline(value):
    """Trim each line of a long answer and keep paragraph breaks."""
    if _is_nan(value):
        return None, None
    text = str(value).replace("\xa0", " ").replace("\r\n", "\n").replace("\r", "\n")
    lines = []
    for line in text.split("\n"):
        cleaned = re.sub(r"[ \t]+", " ", line).strip()
        if cleaned and cleaned.lower() not in _BLANK_TOKENS:
            lines.append(cleaned)
    if not lines:
        return None, None
    return "\n".join(lines), None


def clean_email(value):
    """Lowercase an email. Flag values that are not a normal address."""
    text, _ = clean_text(value)
    if text is None:
        return None, None
    text = text.lower()
    if not _EMAIL_RE.match(text):
        return text, "malformed email"
    if re.match(r"^\d+@", text):
        return text, "numeric local-part"
    return text, None


def parse_capacity(value):
    """Parse a capacity string to its largest integer.

    ``6`` stays 6 with no warning. ``1 or 2``, ``1 or 2 maximum``, and
    ``3+`` become the largest number, with a warning, because the source
    cell was not a plain integer.
    """
    text, _ = clean_text(value)
    if text is None:
        return None, "missing capacity"
    if re.fullmatch(r"\d+", text):
        return int(text), None
    if re.fullmatch(r"\d+\.0+", text):
        return int(float(text)), None
    numbers = [int(match) for match in re.findall(r"-?\d+", text)]
    if not numbers:
        return None, "non-numeric capacity with no digits"
    return max(numbers), "non-numeric capacity interpreted as maximum"


def split_year_program(value):
    """Split a year/program answer into two normalized lists.

    Examples: ``4/AI, 4/Cybersecurity`` -> years ``['Year 4']``,
    programs ``['AI', 'Cybersecurity']``. ``Open to all`` -> ``Any`` / ``Any``.
    ``3rd or 4th year`` -> ``['Year 3', 'Year 4']`` and programs ``['Any']``.
    Unparseable text is kept as a one-item year list plus a warning.
    """
    text, _ = clean_text(value)
    if text is None:
        return ([], []), None
    low = text.lower()
    if low in {"open to all", "open", "any", "all", "all years", "any year"} or low.startswith("open to all"):
        return (["Any"], ["Any"]), None

    years = []
    programs = []
    for match in re.finditer(r"\b([1-4])\s*/\s*([^,;]+)", text):
        years.append(f"Year {match.group(1)}")
        program = re.sub(r"\s+", " ", match.group(2)).strip(" .")
        if program:
            programs.append(program)

    if not years:
        for match in re.finditer(r"\b(1st|2nd|3rd|4th|first|second|third|fourth)\b", low):
            years.append(_YEAR_WORDS[match.group(1)])
        for match in re.finditer(r"\b([1-4])\s*(?:st|nd|rd|th)?\s*year\b", low):
            years.append(f"Year {match.group(1)}")
        for match in re.finditer(r"\byear\s*([1-4])\b", low):
            years.append(f"Year {match.group(1)}")

    years = _uniq(years)
    programs = _uniq(programs)
    if years and not programs:
        programs = ["Any"]
    if programs and not years:
        years = ["Any"]
    if not years and not programs:
        return ([text], []), "unparseable year/program"
    return (years, programs), None


def split_multi_value(value, delimiters=(";", ",")):
    """Split a free-text list into trimmed, non-empty strings.

    When a semicolon is present it is the only delimiter, so a value such
    as ``Mohamed, Aly; Shaat, Omar`` stays two people. Newlines are always
    treated as separators. Commas are used only when there is no semicolon.
    """
    if _is_nan(value):
        return [], None
    text = str(value).replace("\xa0", " ").replace("\r\n", "\n").replace("\r", "\n")
    if text.strip().lower() in _BLANK_TOKENS:
        return [], None
    if ";" in text and ";" in delimiters:
        pattern = r"[;\n]+"
    elif "," in delimiters:
        pattern = r"[,\n]+"
    else:
        pattern = "|".join(re.escape(item) for item in delimiters) or r"\n+"
    parts = []
    for part in re.split(pattern, text):
        cleaned = re.sub(r"\s+", " ", part).strip(" \t;")
        if cleaned and cleaned.lower() not in _BLANK_TOKENS:
            parts.append(cleaned)
    return parts, None


def pair_names_ids(names_list, ids_list):
    """Zip names and IDs when the counts match.

    If the counts differ, do not guess a pairing. Both lists are kept,
    each row is marked incomplete, and a warning is returned.
    """
    names = list(names_list or [])
    ids = list(ids_list or [])
    if len(names) == len(ids):
        pairs = [
            {"name": name, "student_id": student_id, "incomplete": False}
            for name, student_id in zip(names, ids)
        ]
        return pairs, None
    pairs = [{"name": name, "student_id": None, "incomplete": True} for name in names]
    pairs.extend({"name": None, "student_id": student_id, "incomplete": True} for student_id in ids)
    detail = f"{len(names)} names, {len(ids)} IDs"
    return pairs, detail


def normalize_student_id(value, expected_length=8):
    """Strip internal spaces and require a numeric ID of the expected length."""
    text, _ = clean_text(value)
    if text is None:
        return None, None
    compact = re.sub(r"\s+", "", text)
    if re.fullmatch(r"\d+\.0+", compact):
        compact = str(int(float(compact)))
    if not compact.isdigit() or len(compact) != expected_length:
        return compact, "student id is not numeric of expected length"
    return compact, None


def normalize_project_id(value):
    """Normalize a Project_ID, including a choice label such as ``SRP26-001 | Title``."""
    text, _ = clean_text(value)
    if text is None:
        return None, None
    match = re.search(r"SRP\s*(\d{2})\s*-\s*(\d{1,3})", text, re.IGNORECASE)
    if not match:
        compact = re.sub(r"\s+", "", text).upper()
        return compact, "Project_ID does not match SRP<YY>-<NNN>"
    return f"SRP{match.group(1)}-{int(match.group(2)):03d}", None


def parse_registration_action(value):
    """Map the registration form's Action choice to ``register`` or ``unregister``."""
    text, _ = clean_text(value)
    if text is None:
        return None, "missing registration action"
    low = text.lower()
    if "unregister" in low:
        return "unregister", None
    if "register" in low:
        return "register", None
    return None, "unmapped registration action"


def parse_yes_no(value):
    """Map common yes/no spellings to True or False.

    Unmapped text returns ``(None, warning)``. Callers still have the raw
    cell, so the original answer is not dropped.
    """
    text, _ = clean_text(value)
    if text is None:
        return None, None
    key = text.lower().rstrip(".")
    if key in _YES:
        return True, None
    if key in _NO:
        return False, None
    return None, "unmapped yes/no value"


def truncate(value, limit):
    """Collapse whitespace and cut text to about ``limit`` characters."""
    text, _ = clean_text(value)
    if text is None:
        return None
    if len(text) <= limit:
        return text
    cut = text[:limit]
    if " " in cut:
        cut = cut.rsplit(" ", 1)[0]
    return cut.rstrip(" ,;:-") + "\u2026"


def _uniq(items):
    seen = set()
    result = []
    for item in items:
        if item not in seen:
            seen.add(item)
            result.append(item)
    return result
