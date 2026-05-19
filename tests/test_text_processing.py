"""Unit tests for the markdown-level helpers in rtf_to_docx_endnotes.py."""

from __future__ import annotations

import unittest

# pylint: disable=unused-import  # _helpers import has sys.path side effect
from tests._helpers import REPO_ROOT  # noqa: F401  (ensures sys.path)
from rtf_to_docx_endnotes import (
    INLINE_RE,
    NOTES_HEADER_RE,
    NOTE_START_RE,
    build_footnoted_markdown,
    parse_notes,
    rewrite_body,
    split_body_notes,
    strip_paragraph_bold,
)


class TestSplitBodyNotes(unittest.TestCase):
    def test_basic_split(self):
        md = "# Title\n\nBody paragraph one.\n\n## Endnotes\n\n[1] First note.\n"
        body, notes = split_body_notes(md)
        self.assertIn("Body paragraph one.", body)
        self.assertNotIn("Endnotes", body)
        self.assertIn("[1] First note.", notes)

    def test_alternative_headers_caseinsensitive(self):
        for header in (
            "## Notes",
            "### footnotes",
            "## Anmerkungen",
            "## ENDNOTES",
            "## References",
        ):
            md = f"Body text.\n\n{header}\n\n[1] note"
            body, notes = split_body_notes(md)
            self.assertNotIn(
                header.lstrip("# ").strip(),
                body,
                f"header leaked into body for {header!r}",
            )
            self.assertIn("[1] note", notes)

    def test_missing_header_raises(self):
        with self.assertRaises(SystemExit):
            split_body_notes("just body, no notes section")


class TestParseNotes(unittest.TestCase):
    def test_simple_notes(self):
        notes_md = "[1] First note text.\n\n[2] Second note text.\n"
        d = parse_notes(notes_md)
        self.assertEqual(d[1], "First note text.")
        self.assertEqual(d[2], "Second note text.")

    def test_bold_wrapped_notes_stripped(self):
        # Source RTF often wraps each note in **...**
        notes_md = "**[1] Hansjörg Schmid, *Soziale Konflikte*, 2024.**\n"
        d = parse_notes(notes_md)
        # The leading **\[N\] and trailing ** are stripped; intra-paragraph
        # *italic* stays as-is.
        self.assertEqual(d[1], "Hansjörg Schmid, *Soziale Konflikte*, 2024.")

    def test_escaped_brackets(self):
        # pandoc emits `\[1\]` in some configurations.
        notes_md = "\\[3\\] Note three.\n"
        d = parse_notes(notes_md)
        self.assertEqual(d[3], "Note three.")

    def test_multiline_note_collapses(self):
        notes_md = "[1] Line one of note\ncontinuing on line two.\n"
        d = parse_notes(notes_md)
        self.assertEqual(d[1], "Line one of note continuing on line two.")

    def test_no_notes_returns_empty(self):
        self.assertEqual(parse_notes(""), {})
        self.assertEqual(parse_notes("just some intro text\nbut nothing numbered"), {})

    def test_out_of_order_numbers_preserved(self):
        # Numbers in non-sequential order: keep as-is.
        notes_md = "[3] three\n\n[1] one\n\n[2] two\n"
        d = parse_notes(notes_md)
        self.assertEqual(d[1], "one")
        self.assertEqual(d[2], "two")
        self.assertEqual(d[3], "three")


class TestRewriteBody(unittest.TestCase):
    def test_replaces_numeric_markers_only(self):
        body = "Quote.[1] More text.[42] Slide marker [S2] stays. Author tag [Liedhegener] stays."
        out = rewrite_body(body, valid_nums={1, 42})
        self.assertIn("[^1]", out)
        self.assertIn("[^42]", out)
        self.assertIn("[S2]", out)
        self.assertIn("[Liedhegener]", out)
        # No raw [1] / [42] left
        self.assertNotIn("[1]", out)
        self.assertNotIn("[42]", out)

    def test_unknown_number_left_alone(self):
        body = "Cite.[99] More."
        out = rewrite_body(body, valid_nums={1, 2})
        self.assertEqual(out, "Cite.[99] More.")  # 99 not a defined note

    def test_escaped_brackets_supported(self):
        body = "Quote.\\[5\\] more"
        out = rewrite_body(body, valid_nums={5})
        self.assertIn("[^5]", out)


class TestStripParagraphBold(unittest.TestCase):
    def test_whole_paragraph_unwrapped(self):
        md = "**Whole paragraph wrapped in bold.**\n\nSecond paragraph plain."
        out, n = strip_paragraph_bold(md)
        self.assertEqual(n, 1)
        self.assertNotIn("**Whole", out)
        self.assertIn("Whole paragraph wrapped in bold.", out)
        self.assertIn("Second paragraph plain.", out)

    def test_intra_paragraph_bold_kept(self):
        md = "A paragraph with **a bold span** inside, not wrapped overall."
        out, n = strip_paragraph_bold(md)
        self.assertEqual(n, 0)
        self.assertIn("**a bold span**", out)

    def test_mixed_paragraphs(self):
        md = (
            "**Bold-wrapped one.**\n\n"
            "Normal paragraph.\n\n"
            "**Bold-wrapped two with *italic* inside.**\n\n"
            "Another normal."
        )
        out, n = strip_paragraph_bold(md)
        self.assertEqual(n, 2)
        # italic survives
        self.assertIn("*italic*", out)
        # bold wrappers gone but text content kept
        self.assertIn("Bold-wrapped one.", out)
        self.assertIn("Bold-wrapped two with *italic* inside.", out)

    def test_paragraph_with_double_bold_skipped(self):
        # `**foo** bar **baz**` is two emphasis spans, not a whole-paragraph
        # wrap; must not be naively unwrapped.
        md = "**foo** bar **baz**"
        out, n = strip_paragraph_bold(md)
        self.assertEqual(n, 0)
        self.assertEqual(out, md)


class TestBuildFootnotedMarkdown(unittest.TestCase):
    def test_definitions_in_numeric_order(self):
        body = "Lead text.[^2] More.[^1]"
        notes = {2: "second", 1: "first"}
        out = build_footnoted_markdown(body, notes)
        self.assertIn(body, out)
        # Definitions appear in ascending numeric order.
        idx1 = out.find("[^1]: first")
        idx2 = out.find("[^2]: second")
        self.assertGreater(idx1, 0)
        self.assertGreater(idx2, idx1)


class TestRegexes(unittest.TestCase):
    def test_inline_re_digit_only(self):
        # INLINE_RE only matches `[<digits>]` — slide markers like [S2] and
        # bracketed names like [Liedhegener] are skipped at the regex level.
        self.assertEqual(
            INLINE_RE.findall("[1] [22] [Liedhegener] [S2]"),
            ["1", "22"],
        )

    def test_notes_header_recognized(self):
        self.assertIsNotNone(NOTES_HEADER_RE.match("## Endnotes"))
        self.assertIsNotNone(NOTES_HEADER_RE.match("### footnotes"))
        self.assertIsNotNone(NOTES_HEADER_RE.match("# Anmerkungen"))
        self.assertIsNone(NOTES_HEADER_RE.match("## End of section"))

    def test_note_start_recognized(self):
        self.assertEqual(NOTE_START_RE.match("[7] something").group("num"), "7")
        self.assertEqual(NOTE_START_RE.match("**[7]** something").group("num"), "7")
        self.assertEqual(NOTE_START_RE.match("\\[7\\] something").group("num"), "7")
        self.assertIsNone(NOTE_START_RE.match("regular text"))


if __name__ == "__main__":
    unittest.main()
