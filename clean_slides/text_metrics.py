"""
Text measurement helpers.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

EMU_PER_PT = 12_700  # 1 typographic point = 12,700 EMU
WORD_RE = re.compile(r"[\w''-]+", re.UNICODE)

# Character-width classes (proportion of em-width at given font size)
CHAR_NARROW = set("ilI|.,'`:")
CHAR_WIDE = set("mwMW")

# Proportional width factors for each character class
_WIDTH_WIDE = 0.78
_WIDTH_NARROW = 0.30
_WIDTH_UPPER = 0.62
_WIDTH_DIGIT = 0.55
_WIDTH_DEFAULT = 0.50
_WIDTH_SPACE = 0.25  # inter-word space

# Font-family width adjustments (Georgia is ~5% wider than Arial)
_FONT_FACTOR_GEORGIA = 1.05
_FONT_FACTOR_DEFAULT = 1.0


@dataclass
class FontSpec:
    """Font definition."""

    name: str
    path: str


class TextMetrics:
    """Estimate text dimensions using font metrics heuristics."""

    def __init__(self, fonts: dict[str, FontSpec] | None = None, fudge: float = 1.12):
        """*fudge* is a global safety multiplier on all width/height estimates,
        compensating for differences between our heuristic and PowerPoint's
        actual text shaping engine."""
        self.fonts = fonts or {}
        self.fudge = fudge

    def word_width(self, word: str, font: str, size_pt: int) -> int:
        """Return word width in EMU (heuristic character-width model)."""
        if not word:
            return 0

        font_lower = (font or "").lower()
        font_factor = _FONT_FACTOR_GEORGIA if "georgia" in font_lower else _FONT_FACTOR_DEFAULT

        total = 0.0
        for ch in word:
            if ch in CHAR_WIDE:
                total += _WIDTH_WIDE
            elif ch in CHAR_NARROW:
                total += _WIDTH_NARROW
            elif ch.isupper():
                total += _WIDTH_UPPER
            elif ch.isdigit():
                total += _WIDTH_DIGIT
            else:
                total += _WIDTH_DEFAULT

        width = total * size_pt * EMU_PER_PT * font_factor
        return int(width * self.fudge)

    def longest_word(self, text: str) -> str:
        """Return longest token that cannot wrap."""
        words = WORD_RE.findall(text or "")
        if not words:
            return ""
        return max(words, key=len)

    def lines_needed(self, text: str | None, width_emu: int, font: str, size_pt: int) -> int:
        """Estimate how many lines text will wrap to."""
        if width_emu <= 0:
            return 0
        if text is None:
            return 0

        font_lower = (font or "").lower()
        font_factor = _FONT_FACTOR_GEORGIA if "georgia" in font_lower else _FONT_FACTOR_DEFAULT
        space_width = int(size_pt * EMU_PER_PT * _WIDTH_SPACE * font_factor * self.fudge)

        total_lines = 0
        for raw_line in str(text).splitlines() or [""]:
            if not raw_line:
                total_lines += 1
                continue

            current = 0
            line_count = 1
            for word in WORD_RE.findall(raw_line):
                word_width = self.word_width(word, font, size_pt)
                if current == 0:
                    current = word_width
                    continue

                if current + space_width + word_width <= width_emu:
                    current += space_width + word_width
                else:
                    line_count += 1
                    current = word_width

            total_lines += line_count

        return total_lines

    def text_height(
        self,
        text: str,
        width_emu: int,
        font: str,
        size_pt: int,
        line_spacing: float = 1.14,  # PowerPoint default single-spacing multiplier
    ) -> int:
        """Return estimated text height in EMU."""
        lines = self.lines_needed(text, width_emu, font, size_pt)
        line_height = int(size_pt * EMU_PER_PT * line_spacing)
        return int(lines * line_height * self.fudge)

    def text_width_no_wrap(self, text: str | None, font: str, size_pt: int) -> int:
        """Return estimated width in EMU if *text* is kept on one line.

        Uses the same tokenization/spacing assumptions as :meth:`lines_needed`.
        For multi-line strings, returns the width of the widest line.
        """
        if not text:
            return 0

        font_lower = (font or "").lower()
        font_factor = _FONT_FACTOR_GEORGIA if "georgia" in font_lower else _FONT_FACTOR_DEFAULT
        space_width = int(size_pt * EMU_PER_PT * _WIDTH_SPACE * font_factor * self.fudge)

        max_w = 0
        for raw_line in str(text).splitlines() or [""]:
            words = WORD_RE.findall(raw_line)
            if not words:
                continue
            w = sum(self.word_width(word, font, size_pt) for word in words)
            if len(words) > 1:
                w += (len(words) - 1) * space_width
            max_w = max(max_w, w)

        return int(max_w)
