"""Unit tests for docxtool's standalone bits."""

# pylint: disable=protected-access  # tests intentionally probe private helpers
from __future__ import annotations

import io
import subprocess
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path

from tests._helpers import (  # noqa: F401
    REPO_ROOT,
    W_DOC_HEAD,
    W_DOC_TAIL,
    write_docx,
    part,
)
import docxtool


# Document with both rsid attrs and textutil-style non-standard tags.
DOC_XML_DIRTY = (
    W_DOC_HEAD
    + "<w:body>"
    + '<w:p w14:paraId="A1B2" w:rsidR="1234">'
    + '<w:r><w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="28"/><w:sz-cs w:val="28"/></w:rPr>'
    + '<w:t xml:space="preserve">Hello</w:t></w:r></w:p>'
    + "</w:body>"
    + W_DOC_TAIL
)


class TestNormalizeNonstandardTags(unittest.TestCase):
    def test_rename_sz_cs(self):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "t.docx"
            write_docx(path, DOC_XML_DIRTY)
            n = docxtool._normalize_nonstandard_tags(path)
            # Fixture has one <w:sz-cs .../> self-closing element → 1 rename.
            self.assertEqual(n, 1)
            after = part(path, "word/document.xml")
            self.assertNotIn("w:sz-cs", after)
            self.assertIn("w:szCs", after)

    def test_noop_on_clean_docx(self):
        clean = W_DOC_HEAD + "<w:body><w:p><w:r><w:t>x</w:t></w:r></w:p></w:body>" + W_DOC_TAIL
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "t.docx"
            write_docx(path, clean)
            n = docxtool._normalize_nonstandard_tags(path)
            self.assertEqual(n, 0)


class TestCleanSubcommand(unittest.TestCase):
    def test_dry_run_reports_count_without_changes(self):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "t.docx"
            write_docx(path, DOC_XML_DIRTY)
            before = path.read_bytes()

            buf = io.StringIO()
            with redirect_stdout(buf):
                rc = docxtool.cmd_clean([str(path), "--dry-run"])
            self.assertEqual(rc, 0)
            self.assertIn("would strip", buf.getvalue())
            # nothing written
            self.assertEqual(path.read_bytes(), before)

    def test_in_place_strips_rsids(self):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "t.docx"
            write_docx(path, DOC_XML_DIRTY)
            buf = io.StringIO()
            with redirect_stdout(buf):
                rc = docxtool.cmd_clean([str(path)])
            self.assertEqual(rc, 0)
            after = part(path, "word/document.xml")
            self.assertNotIn("paraId", after)
            self.assertNotIn("rsidR", after)

    def test_also_normalize_tags(self):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "t.docx"
            write_docx(path, DOC_XML_DIRTY)
            buf = io.StringIO()
            with redirect_stdout(buf):
                rc = docxtool.cmd_clean([str(path), "--also-normalize-tags"])
            self.assertEqual(rc, 0)
            after = part(path, "word/document.xml")
            self.assertNotIn("w:sz-cs", after)
            self.assertIn("w:szCs", after)
            self.assertNotIn("paraId", after)

    def test_output_to_new_path_leaves_original(self):
        with tempfile.TemporaryDirectory() as td:
            src = Path(td) / "in.docx"
            dst = Path(td) / "out.docx"
            write_docx(src, DOC_XML_DIRTY)
            src_bytes = src.read_bytes()
            buf = io.StringIO()
            with redirect_stdout(buf):
                rc = docxtool.cmd_clean([str(src), "-o", str(dst)])
            self.assertEqual(rc, 0)
            self.assertEqual(src.read_bytes(), src_bytes)
            self.assertTrue(dst.exists())
            self.assertNotIn("paraId", part(dst, "word/document.xml"))


class TestDispatcher(unittest.TestCase):
    """End-to-end smoke tests of the docxtool CLI."""

    def _run(self, *argv: str) -> subprocess.CompletedProcess:
        return subprocess.run(
            [sys.executable, str(REPO_ROOT / "docxtool.py"), *argv],
            capture_output=True,
            text=True,
            check=False,
        )

    def test_top_level_help_lists_subcommands(self):
        r = self._run("--help")
        self.assertEqual(r.returncode, 0)
        out = r.stdout + r.stderr
        for cmd in ("notes", "transplant", "translate", "debug", "clean"):
            self.assertIn(cmd, out)

    def test_unknown_subcommand_errors(self):
        r = self._run("not-a-real-cmd")
        self.assertNotEqual(r.returncode, 0)
        self.assertIn("unknown subcommand", r.stderr)

    def test_clean_help_is_local(self):
        # `clean` is built into docxtool itself, not delegated.
        r = self._run("clean", "--help")
        self.assertEqual(r.returncode, 0)
        self.assertIn("rsid", r.stdout.lower() + r.stderr.lower())


if __name__ == "__main__":
    unittest.main()
