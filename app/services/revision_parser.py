# app/services/revision_parser.py
from __future__ import annotations
from typing import List, Optional, Tuple
import re
from app.domain.revision_rules import DATE_REGEX, DESC_KEYWORDS, REVISION_REGEX_FALLBACK

WORD_REV = re.compile(r"\brev(?:ision)?\b", re.IGNORECASE)

class RevisionParser:
    def __init__(self, revision_regex: str | None, certainty_lock: bool = True, fill_missing: bool = True):
        self.rev_re = re.compile(revision_regex, re.IGNORECASE) if revision_regex else REVISION_REGEX_FALLBACK
        self.certainty_lock = certainty_lock
        self.fill_missing = fill_missing


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
        Extract a plausible revision token, prioritizing tokens that start with a letter.
        Avoid numeric-only tokens when the same cell looks like a date, unless the
        cell explicitly contains 'rev'/'revision'.
        """
        if not text:
            return None
        t = self._norm(text)

        # tokenization
        tokens = [c for c in re.split(r"[^A-Za-z0-9]+", t) if c]
        if not tokens:
            return None

        # explicit "rev: X" capture gets highest priority
        m = re.search(r"\brev(?:ision)?\s*[:\-]?\s*([A-Za-z0-9]{1,4})", t, re.IGNORECASE)
        explicit = [m.group(1)] if m else []

        # prioritize tokens that begin with a letter (A, B2, C03, P01...)
        letter_first = [tok for tok in tokens if tok[0].isalpha()]

        # other alphanumerics that contain letters but don’t start with one
        alnum_with_letters = [tok for tok in tokens
                              if any(ch.isalpha() for ch in tok) and not tok[0].isalpha()]

        # numeric-only tokens – only allow if the cell is NOT a date,
        # or if the cell explicitly mentions "rev"
        looks_like_date = bool(DATE_REGEX.search(t))
        has_rev_word = bool(WORD_REV.search(t))
        numeric_only = [tok for tok in tokens if tok.isdigit()]
        numeric_ok = numeric_only if (not looks_like_date or has_rev_word) else []

        # build priority list and test against the chosen rev regex
        candidates = explicit + letter_first + alnum_with_letters + numeric_ok
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
        scores: dict[int, dict[str, int]] = {}  # idx -> {'rev': int, 'date': int, 'desc': int}

        def bump(i, k, v=1):
            scores.setdefault(i, {"rev": 0, "date": 0, "desc": 0})
            scores[i][k] += v

        # ----- SCORING PASS -----
        for row in rows[:max_rows]:
            for idx, raw in enumerate(row):
                txt = self._norm(raw)
                if not txt:
                    continue

                has_date = bool(DATE_REGEX.search(txt))
                has_rev_word = bool(WORD_REV.search(txt))
                rev_tok = self._extract_rev_token(txt)

                # REV signal (heavy) – don't count from date-looking cells unless labeled “rev”
                if rev_tok and (not has_date or has_rev_word):
                    bump(idx, "rev", 3)

                # header-like 'rev' word gives a small hint
                if has_rev_word:
                    bump(idx, "rev", 1)

                # DATE signal
                if has_date:
                    bump(idx, "date", 2)

                # DESC signal
                if self._looks_like_desc(txt):
                    bump(idx, "desc", 1)

        if not scores:
            return (None, None, None)

        # keep an immutable copy for fallbacks
        scores_all = {i: sc.copy() for i, sc in scores.items()}

        # ----- OPTIONAL CERTAINTY LOCK (before picking) -----
        if getattr(self, "certainty_lock", False):
            DATE_STRONG = 4  # (e.g., 2 rows * weight 2)
            REV_STRONG = 6  # (e.g., 2 rows * weight 3)
            for i, sc in scores.items():
                # strong date column should not be considered rev unless it also had strong rev
                if sc["date"] >= DATE_STRONG and sc["rev"] <= 1:
                    sc["rev"] = -999
                # strong rev column (with no date hints) should not be considered date
                if sc["rev"] >= REV_STRONG and sc["date"] == 0:
                    sc["date"] = -999

        # ----- PRIMARY EXCLUSIVE PICKS -----
        # Work on a mutable copy so we can pop selected columns
        work = {i: sc.copy() for i, sc in scores.items()}

        def pick_exclusive(metric: str) -> Optional[int]:
            if not work:
                return None
            best_idx, best_val = None, -1_000_000
            for i, sc in work.items():
                v = sc[metric]
                if v > best_val:
                    best_idx, best_val = i, v
            if best_val <= 0:
                return None
            work.pop(best_idx, None)
            return best_idx

        r_idx = pick_exclusive("rev")
        d_idx = pick_exclusive("desc")
        dt_idx = pick_exclusive("date")

        # ----- FALLBACK FILL (avoid blanks if possible) -----
        if getattr(self, "fill_missing", False):
            used = {i for i in (r_idx, d_idx, dt_idx) if i is not None}

            DATE_STRONG = 4
            REV_STRONG = 6

            def fallback(metric: str, current_idx: Optional[int]) -> Optional[int]:
                if current_idx is not None:
                    return current_idx
                # candidates = remaining columns not used yet
                candidates = [i for i in scores_all.keys() if i not in used]
                if not candidates:
                    return None

                # Filter out strongly conflicting columns for the requested metric
                pruned: list[int] = []
                for i in candidates:
                    sc = scores_all[i]
                    if metric == "rev":
                        # avoid columns that are strongly date
                        if sc["date"] >= DATE_STRONG and sc["rev"] <= 1:
                            continue
                    elif metric == "date":
                        # avoid columns that are strongly rev
                        if sc["rev"] >= REV_STRONG and sc["date"] == 0:
                            continue
                    else:  # desc
                        # avoid columns that are both strongly rev or strongly date (acts as a weak filter)
                        if sc["rev"] >= REV_STRONG or sc["date"] >= DATE_STRONG:
                            # still allow if it also has good desc evidence
                            if sc["desc"] <= 1:
                                continue
                    pruned.append(i)

                pool = pruned or candidates  # if pruning removed all, fall back to all
                # choose the one with the highest score for that metric
                best_idx = max(pool, key=lambda i: scores_all[i][metric])
                if scores_all[best_idx][metric] > 0:
                    used.add(best_idx)
                    return best_idx

                # if truly no signal for this metric, leave it None (too risky to force)
                return None

            r_idx = fallback("rev", r_idx)
            d_idx = fallback("desc", d_idx)
            dt_idx = fallback("date", dt_idx)

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
