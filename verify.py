"""Check every claim against the paper's own text, without asking the model anything.

For each quote: is it really in the paper (exactly or nearly), on which page, under
which heading, and what are the sentences around it? Context and section come from
the OCR text, so a reviewer never has to trust the model for them.
"""

import re
from bisect import bisect_left, bisect_right

from rapidfuzz import fuzz

MATCH_THRESHOLD = 90  # minimum fuzzy score (0-100) to count a quote as found
SHORT_QUOTE_WORDS = 8  # shorter quotes prove little, so they get flagged
CONTEXT_SENTENCES = 1  # sentences of context on each side of the quote

# Characters that can differ between the model's copy and the OCR text.
REPLACE = {"‘": "'", "’": "'", "“": '"', "”": '"', "–": "-", "—": "-", " ": " "}
DROP = set("*_#`\\")  # markdown and LaTeX markup


def normalize(text: str) -> tuple[str, list[int]]:
    """Lowercase, unify quotes and dashes, drop markup, rejoin words split by a
    line-break hyphen, and collapse whitespace.

    Also returns, for each character of the result, its position in the original
    text, so a match found in the normalized text can be traced back to the original.
    """
    chars, positions = [], []
    i = 0

    while i < len(text):
        ch = REPLACE.get(text[i], text[i])

        if ch == "-" and text[i + 1 : i + 2] == "\n":  # "exam-\nple" -> "example"
            i += 2

            while text[i : i + 1] in (" ", "\t"):
                i += 1

            continue

        if ch.isspace():
            if chars and chars[-1] != " ":
                chars.append(" ")
                positions.append(i)

        elif ch not in DROP:
            for lowered in ch.lower():
                chars.append(lowered)
                positions.append(i)

        i += 1

    return "".join(chars), positions


class SourceText:
    """The OCR text of one paper, prepared once so many quotes can be checked against it."""

    def __init__(self, ocr: dict):
        self.page_starts, self.page_numbers, parts = [], [], []
        offset = 0
        for page in ocr["pages"]:
            self.page_starts.append(offset)
            self.page_numbers.append(page["page"])
            parts.append(page["markdown"])
            offset += len(page["markdown"]) + 2  # +2 for the "\n\n" between pages
        self.text = "\n\n".join(parts)
        self.normalized, self.positions = normalize(self.text)

        # Where each sentence starts: after . ! or ? plus whitespace, or after a blank line.
        self.sentence_starts = [m.end() for m in re.finditer(r"[.!?][\"')\]]*\s+|\n\s*\n", self.text)]
        headings = list(re.finditer(r"^#{1,6}\s+(.+)$", self.text, re.MULTILINE))
        self.heading_starts = [m.start() for m in headings]
        self.heading_titles = [m.group(1).replace("*", "").strip() for m in headings]

    def page_at(self, offset: int) -> int:
        return self.page_numbers[bisect_right(self.page_starts, offset) - 1]

    def section_at(self, offset: int) -> str | None:
        index = bisect_right(self.heading_starts, offset) - 1
        return self.heading_titles[index] if index >= 0 else None

    def find(self, quote: str, page: int) -> tuple[float, int, int] | None:
        """Score and (start, end) of the quote in the original text, or None if not found."""
        needle = normalize(quote)[0].strip()

        if not needle:
            return None

        # Exact match first. Short quotes can occur several times, so prefer the cited page.
        hits = [m.start() for m in re.finditer(re.escape(needle), self.normalized)]

        if hits:
            start = next((h for h in hits if self.page_at(self.positions[h]) == page), hits[0])
            score, end = 100.0, start + len(needle)

        else:
            match = fuzz.partial_ratio_alignment(needle, self.normalized, score_cutoff=MATCH_THRESHOLD)
            if match is None:
                return None
            score, start, end = match.score, match.dest_start, match.dest_end

        return score, self.positions[start], self.positions[end - 1] + 1

    def context(self, start: int, end: int) -> str:
        """The quote plus CONTEXT_SENTENCES whole sentences on each side, from the paper itself."""
        before = bisect_right(self.sentence_starts, start) - 1 - CONTEXT_SENTENCES
        after = bisect_left(self.sentence_starts, end) + CONTEXT_SENTENCES
        context_start = self.sentence_starts[before] if before >= 0 else 0
        context_end = self.sentence_starts[after] if after < len(self.sentence_starts) else len(self.text)
        snippet = self.text[context_start:context_end]
        # Tidy for reading: rejoin "frame-\nwork", drop markdown, put it all on one line.
        snippet = re.sub(r"-\n\s*", "", snippet).replace("*", "").replace("#", "")
        return " ".join(snippet.split())


def verify_claims(rows: list[dict], ocr: dict) -> list[dict]:
    """Add quote_found, match_score, found_page, context and flags to each reported claim.

    When the quote is found, the section is taken from the paper's headings
    instead of from the model. Rows are updated in place and also returned.
    """
    source = SourceText(ocr)

    for row in rows:
        if row["status"] != "reported":
            continue

        flags = []
        match = source.find(row["quote"], row["page"])

        if match is None:
            row.update(quote_found=False, match_score=None, found_page=None, context=None)
            flags.append("quote not found")

        else:
            score, start, end = match
            row.update(
                quote_found=True,
                match_score=round(score),
                found_page=source.page_at(start),
                context=source.context(start, end),
                section=source.section_at(start) or row["section"],
            )

            if row["found_page"] != row["page"]:
                flags.append(f"cited page {row['page']}, found on page {row['found_page']}")

        if len(row["quote"].split()) < SHORT_QUOTE_WORDS:
            flags.append("short quote")

        if "..." in row["quote"] or "…" in row["quote"]:
            flags.append("shortened with ...")

        row["flags"] = "; ".join(flags)
        
    return rows
