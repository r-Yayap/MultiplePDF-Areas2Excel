# app/services/revision_parser.py
from __future__ import annotations
from typing import List, Optional, Tuple
import re
from app.domain.revision_rules import DATE_REGEX, DESC_KEYWORDS, REVISION_REGEX_FALLBACK

WORD_REV = re.compile(r"\brev(?:ision)?\b", re.IGNORECASE)

class RevisionParser:
    def __init__(self, revision_regex: str | None):
        self.rev_re = re.compile(revision_regex, re.IGNORECASE) if revision_regex else REVISION_REGEX_FALLBACK

    # ---------- helpers ----------
    @staticmethod
    def _norm(text: str) -> str:
        if text is None:
            return ""
        t = str(text)
        # normalize dashes & whitespace
        t = t.replace("—", "-").replace("–", "-")
        t = re.sub(r"\s+", " ", t)
        return t.strip()

    def _extract_rev_token(self, text: str) -> Optional[str]:
        """
        Pull a revision token from within a messy cell.
        Examples that should match: 'Rev A', 'A – IFC', 'REVISED TO C03', '(P02) Issued for Construction'
        Strategy: scan all word-ish tokens and also look at bigrams around 'rev' headers.
        """
        if not text:
            return None
        t = self._norm(text)

        # try all tokens & dash/paren-stripped chunks
        # e.g., 'A-IFC' -> check 'A', 'IFC'; '(C03)' -> 'C03'
        candidates = []
        # split on non-alnum to get tokens
        candidates.extend([c for c in re.split(r"[^A-Za-z0-9]+", t) if c])

        # also look for sequences like "rev A", "rev: A01"
        m = re.search(r"\brev(?:ision)?\s*[:\-]?\s*([A-Za-z0-9]{1,4})", t, re.IGNORECASE)
        if m:
            candidates.insert(0, m.group(1))

        # scan candidates against the chosen rev regex
        for cand in candidates:
            if self.rev_re.fullmatch(cand.upper()):
                return cand.upper()
        return None

    @staticmethod
    def _looks_like_desc(text: str) -> bool:
        if not text:
            return False
        t = RevisionParser._norm(text)
        low = t.lower()
        # positive signals
        kw_hit = any(kw in low for kw in DESC_KEYWORDS)
        has_space = " " in t
        has_letters = re.search(r"[A-Za-z]", t) is not None
        longish = len(t) >= 6
        # negative signals (avoid short pure codes)
        mostly_code = re.fullmatch(r"[A-Za-z]{0,3}\d{0,3}", t) is not None
        return (kw_hit or (has_space and has_letters and longish)) and not mostly_code

    # ---------- detection ----------
    def detect_column_indices(self, rows: List[List[str]], max_rows: int = 4) -> Tuple[Optional[int], Optional[int], Optional[int]]:
        scores = {}  # idx -> dict(rev=, date=, desc=)

        def bump(i, k, v=1):
            scores.setdefault(i, {"rev": 0, "date": 0, "desc": 0})
            scores[i][k] += v

        for row in rows[:max_rows]:
            for idx, raw in enumerate(row):
                txt = self._norm(raw)
                if not txt:
                    continue

                # revision signal (heaviest)
                if self._extract_rev_token(txt):
                    bump(idx, "rev", 3)
                if WORD_REV.search(txt):
                    bump(idx, "rev", 1)   # header-like hint

                # date signal
                if DATE_REGEX.search(txt):
                    bump(idx, "date", 2)

                # description signal
                if self._looks_like_desc(txt):
                    bump(idx, "desc", 1)

        def pick(metric: str) -> Optional[int]:
            if not scores:
                return None
            best_idx, best_val = None, -1
            for i, sc in scores.items():
                v = sc[metric]
                if v > best_val:
                    best_idx, best_val = i, v
            # require at least one hit for that metric
            if best_val <= 0:
                return None
            # remove picked column to avoid reusing it for other roles
            scores.pop(best_idx, None)
            return best_idx

        r_idx = pick("rev")
        d_idx = pick("desc")
        dt_idx = pick("date")
        return (r_idx, d_idx, dt_idx)

    # ---------- row filtering ----------
    def is_footer_or_header_row(self, row: List[str], rev_idx: Optional[int]) -> bool:
        normed = [self._norm(c).lower() for c in row]
        # mostly empty -> skip
        if sum(1 for c in normed if not c) >= max(1, int(len(normed) * 0.75)):
            return True

        # if it contains a valid rev token anywhere, treat as data
        if any(self._extract_rev_token(c) for c in normed if c):
            return False

        # classic header junk, but avoid killing real descriptions that say "revised"
        headerish = any(
            re.search(r"\b(revision|rev\b|date\b|description|chk.?d|checked|approved|no\.)", c)
            for c in normed if c
        )
        return headerish

    # ---------- parsing ----------
    def parse_row(self, row: List[str], rev_idx, desc_idx, date_idx):
        def get(idx):
            return self._norm(row[idx]) if idx is not None and idx < len(row) else ""

        rev  = get(rev_idx)
        desc = get(desc_idx)
        date = get(date_idx)

        # salvage rev/date/desc if column guess failed
        if not rev:
            # scan entire row for a rev token
            for c in row:
                tok = self._extract_rev_token(self._norm(c))
                if tok:
                    rev = tok
                    break
        else:
            # clean to the token if cell had extra words
            tok = self._extract_rev_token(rev)
            if tok:
                rev = tok

        if date and not DATE_REGEX.search(date):
            date = ""
        if not date:
            for c in row:
                c2 = self._norm(c)
                if DATE_REGEX.search(c2):
                    date = c2
                    break

        if not desc:
            # pick the longest “sentence-like” cell that isn't clearly a date or rev token
            candidates = [self._norm(c) for c in row if c]
            candidates = [c for c in candidates if not DATE_REGEX.search(c) and not self._extract_rev_token(c)]
            if candidates:
                # prefer one that "looks like description"
                desc_like = [c for c in candidates if self._looks_like_desc(c)]
                desc = max(desc_like or candidates, key=len)

        rev  = rev or None
        desc = desc or None
        date = date or None
        return rev, desc, date

    # ---------- table API ----------
    def parse_table_rows(self, rows: List[List[str]]) -> List[dict]:
        if not rows:
            return []

        # work bottom-up (latest last)
        rows_rev = rows[::-1]

        # first pass: ignore obvious headers/footers
        tmp_r_idx, tmp_d_idx, tmp_dt_idx = self.detect_column_indices(rows_rev)
        filtered = [r for r in rows_rev if not self.is_footer_or_header_row(r, tmp_r_idx)]

        # re-detect using filtered sample
        r_idx, d_idx, dt_idx = self.detect_column_indices(filtered)

        out: List[dict] = []
        for row in filtered:
            if not any(row):
                continue
            rev, desc, date = self.parse_row(row, r_idx, d_idx, dt_idx)

            # keep partials: prefer rows with a rev; else keep (desc+date)
            if rev or (desc and date):
                item = {}
                if rev:  item["rev"]  = rev
                if desc: item["desc"] = desc
                if date: item["date"] = date
                out.append(item)
        return out
