#app/domain/revision_patterns.py

REVISION_PATTERNS = {
    "XX": {
        "pattern": r"^\d{2}$",
        "examples": ["00", "01", "02", "10", "11", "12", "99"]
    },
    "Alphabet Only": {
        "pattern": r"^[A-Z]$",
        "examples": ["A", "B", "C"]
    },
    "Design (DAE)": {
        "pattern": r"^[A-Z]\d{0,2}[a-zA-Z]?$",
        "examples": ["A", "B1", "C3a", "D12"]
    },
    "IFC (DAE)": {
        "pattern": r"^C\d{2}$",
        "examples": ["C00", "C01", "C02"]
    },
    "P0x": {
        "pattern": r"^P\d{2}$",
        "examples": ["P00", "P01", "P02"]
    },
    "[l][n][n]": {
        "pattern": r"^[A-Z]\d{2}$",
        "examples": ["A01", "Z05"]
    },
    "Fuzzy Match": {
        "pattern": r"(^[A-Z]{1,2}\d{1,2}[a-zA-Z]?$)|(^\d{2}$)",
        "examples": ["A", "B1", "C3a", "00", "12"]
    }
}
