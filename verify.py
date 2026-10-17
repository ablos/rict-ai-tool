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
MAX_GAP = 500  # most characters a "..." in a quote may leave out
ELLIPSIS = re.compile(r"\.\.\.|…")

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
        """Score and (start, end) of the quote in the original text, or None if not found.

        A quote shortened with "..." is matched piece by piece: the pieces must appear
        in order and close together, and the match then covers the left-out text too.
        """
        pieces = [normalize(piece)[0].strip() for piece in ELLIPSIS.split(quote)]
        pieces = [piece for piece in pieces if piece]
        if not pieces:
            return None

        first = self._find_piece(pieces[0], page, 0, len(self.normalized))
        if first is None:
            return None
        score, start, end = first
        for piece in pieces[1:]:
            match = self._find_piece(piece, page, end, end + MAX_GAP + len(piece))
            if match is None:
                return None
            score, end = min(score, match[0]), match[2]

        return score, self.positions[start], self.positions[end - 1] + 1

    def find_interrupted(self, quote: str, page: int) -> tuple[float, int, int] | None:
        """A quote that the paper's text interrupts, e.g. with footnotes that OCR placed in
        the middle of a sentence at a page break.

        The quote is cut in two (each part at least 3 words); both parts must appear
        exactly, the second within MAX_GAP characters after the first. Text joined
        from different places in another order still does not match.
        """
        words = normalize(quote)[0].strip().split(" ")

        for cut in range(3, len(words) - 2):
            head, tail = " ".join(words[:cut]), " ".join(words[cut:])

            for hit in re.finditer(re.escape(head), self.normalized):
                window_end = hit.end() + MAX_GAP + len(tail)
                rest = self.normalized.find(tail, hit.end(), window_end)

                if rest >= 0:
                    return 100.0, self.positions[hit.start()], self.positions[rest + len(tail) - 1] + 1

        return None

    def _find_piece(self, needle: str, page: int, low: int, high: int) -> tuple[float, int, int] | None:
        """Best match of a normalized needle within normalized[low:high], as (score, start, end)."""
        window = self.normalized[low:high]

        # Exact match first. Short quotes can occur several times, so prefer the cited page.
        hits = [low + m.start() for m in re.finditer(re.escape(needle), window)]
        if hits:
            start = next((h for h in hits if self.page_at(self.positions[h]) == page), hits[0])
            return 100.0, start, start + len(needle)

        match = fuzz.partial_ratio_alignment(needle, window, score_cutoff=MATCH_THRESHOLD)
        if match is None:
            return None
        return match.score, low + match.dest_start, low + match.dest_end

    def excerpt(self, start: int, end: int) -> str:
        """The paper's own text between start and end, tidied for reading."""
        return _tidy(self.text[start:end])

    def context(self, start: int, end: int) -> str:
        """The quote plus CONTEXT_SENTENCES whole sentences on each side, from the paper itself."""
        before = bisect_right(self.sentence_starts, start) - 1 - CONTEXT_SENTENCES
        after = bisect_left(self.sentence_starts, end) + CONTEXT_SENTENCES
        context_start = self.sentence_starts[before] if before >= 0 else 0
        context_end = self.sentence_starts[after] if after < len(self.sentence_starts) else len(self.text)
        return _tidy(self.text[context_start:context_end])


def _tidy(snippet: str) -> str:
    """Rejoin "frame-\\nwork", drop markdown, and put the text on one line."""
    snippet = re.sub(r"-\n\s*", "", snippet).replace("*", "").replace("#", "")
    return " ".join(snippet.split())


def verify_claims(rows: list[dict], ocr: dict) -> list[dict]:
    """Add quote_found, match_score, context and flags to each reported claim.

    When the quote is found, quote, page and section are taken from the paper itself
    instead of from the model; the model's versions are kept as model_quote and
    model_page. Rows are updated in place and also returned.
    """
    source = SourceText(ocr)

    for row in rows:
        if row["status"] != "reported":
            continue

        flags = []
        # setdefault: verifying the same rows twice must not overwrite the model's versions.
        row.setdefault("model_quote", row["quote"])
        row.setdefault("model_page", row["page"])
        match = source.find(row["model_quote"], row["model_page"])

        if match is None:
            match = source.find_interrupted(row["model_quote"], row["model_page"])

            if match is not None:
                flags.append("quote interrupted by other text in the paper (e.g. footnotes)")

        if match is None:
            row.update(quote_found=False, match_score=None, context=None)
            flags.append("quote not found")

        else:
            score, start, end = match
            row.update(
                quote_found=True,
                match_score=round(score),
                quote=source.excerpt(start, end),
                page=source.page_at(start),
                context=source.context(start, end),
                section=source.section_at(start) or row["section"],
            )

        if len(row["quote"].split()) < SHORT_QUOTE_WORDS:
            flags.append("short quote")

        row["flags"] = "; ".join(flags)
        
    return rows
