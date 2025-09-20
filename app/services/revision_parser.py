from __future__ import annotations
from typing import List, Optional, Tuple
import re
from app.domain.revision_rules import DATE_REGEX, DESC_KEYWORDS, REVISION_REGEX_FALLBACK

class RevisionParser:
    def __init__(self, revision_regex: str | None):
        self.revision_regex = re.compile(revision_regex, re.IGNORECASE) if revision_regex else REVISION_REGEX_FALLBACK

    def detect_column_indices(self, rows: List[List[str]], max_rows: int = 3) -> Tuple[Optional[int], Optional[int], Optional[int]]:
        col_scores = {}
        for row in rows[:max_rows]:
            for idx, cell in enumerate(row):
                text = str(cell).strip()
                if not text:
                    continue
                col_scores.setdefault(idx, {"rev": 0, "date": 0, "desc": 0})
                if self.revision_regex.fullmatch(text.upper()):
                    col_scores[idx]["rev"] += 1
                elif DATE_REGEX.search(text):
                    col_scores[idx]["date"] += 1
                elif any(kw in text.lower() for kw in DESC_KEYWORDS):
                    col_scores[idx]["desc"] += 1

        scores = {k: v.copy() for k, v in col_scores.items()}
        def pick(metric):
            if not scores: return None
            best = max(scores, key=lambda c: scores[c][metric])
            if scores[best][metric] == 0:
                return None
            scores.pop(best)
            return best

        return (pick("rev"), pick("desc"), pick("date"))

    def is_footer_or_header_row(self, row: List[str], rev_idx: Optional[int]) -> bool:
        normalized = [str(c or "").strip().lower() for c in row]
        if sum(1 for c in normalized if c == "") >= max(1, int(len(normalized) * 0.75)):
            return True

        rev_col_val = normalized[rev_idx] if rev_idx is not None and rev_idx < len(normalized) else ""
        if self.revision_regex.fullmatch(rev_col_val.upper()):
            return False

        unwanted = ["revision", "rev date description", "by chk'd", "date", "description", "no.", "rev", "checked by"]
        return any(kw in cell for cell in normalized for kw in unwanted)

    def parse_row(self, row: List[str], rev_idx, desc_idx, date_idx):
        rev  = row[rev_idx].strip()  if rev_idx  is not None and rev_idx  < len(row) else None
        desc = row[desc_idx].strip() if desc_idx is not None and desc_idx < len(row) else None
        date = row[date_idx].strip() if date_idx is not None and date_idx < len(row) else None

        if rev and not self.revision_regex.fullmatch(rev.strip().upper()):
            rev = None
        if date and not DATE_REGEX.search(date):
            date = None
        return rev, desc, date

    def parse_table_rows(self, rows: List[List[str]]) -> List[dict]:
        if not rows:
            return []
        # bottom-up (latest last)
        rows = rows[::-1]
        r_idx, d_idx, dt_idx = self.detect_column_indices(rows)
        filtered = [r for r in rows if not self.is_footer_or_header_row(r, r_idx)]
        r_idx, d_idx, dt_idx = self.detect_column_indices(filtered)

        out = []
        for row in filtered:
            if not any(row): continue
            rev, desc, date = self.parse_row(row, r_idx, d_idx, dt_idx)
            if rev and desc and date:
                out.append({"rev": rev, "desc": desc, "date": date})
        return out
