import re

REVISION_REGEX_FALLBACK = re.compile(r"^[A-Z]{1,2}\d{1,2}[a-zA-Z]?$", re.IGNORECASE)
DATE_REGEX = re.compile(r"""
    (?:\d{1,2}[/-]\d{1,2}[/-]\d{2,4}) |
    (?:\d{1,2}\s*[-]?\s*[A-Za-z]{3,9}\s*[-]?\s*\d{2,4})
""", re.VERBOSE | re.IGNORECASE)

DESC_KEYWORDS = [
    "issued for","issue","submission","schematic","detailed","concept","design",
    "construction","revised","resubmission","ifc","tender","addendum"
]
