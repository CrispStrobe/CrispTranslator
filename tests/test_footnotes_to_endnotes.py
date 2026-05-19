"""Unit tests for the footnotes → endnotes converter."""

from __future__ import annotations

import tempfile
import unittest
import zipfile
from pathlib import Path

# pylint: disable=unused-import  # _helpers import has sys.path side effect
from tests._helpers import (  # noqa: F401
    REPO_ROOT,
    W_DOC_HEAD,
    W_DOC_TAIL,
    write_docx,
    part,
)
from rtf_to_docx_endnotes import footnotes_to_endnotes


# A document with one inline footnote reference.
DOC_XML = (
    W_DOC_HEAD
    + "<w:body><w:p>"
    + '<w:r><w:t xml:space="preserve">Body. </w:t></w:r>'
    + '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>'
    + '<w:footnoteReference w:id="1"/></w:r>'
    + '<w:r><w:t xml:space="preserve"> after.</w:t></w:r>'
    + "</w:p></w:body>"
    + W_DOC_TAIL
)

FOOTNOTES_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:footnote w:id="-1" w:type="separator"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>'
    '<w:footnote w:id="0" w:type="continuationSeparator"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>'
    '<w:footnote w:id="1"><w:p>'
    '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>'
    '<w:r><w:t xml:space="preserve"> the note body</w:t></w:r>'
    "</w:p></w:footnote>"
    "</w:footnotes>"
)


class TestFootnotesToEndnotes(unittest.TestCase):
    def test_full_conversion(self):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "t.docx"
            write_docx(path, DOC_XML, footnotes_xml=FOOTNOTES_XML)

            footnotes_to_endnotes(path)

            with zipfile.ZipFile(path) as zf:
                names = set(zf.namelist())

            # Part renamed
            self.assertNotIn("word/footnotes.xml", names)
            self.assertIn("word/endnotes.xml", names)

            # endnotes.xml uses endnote tags
            en = part(path, "word/endnotes.xml")
            self.assertIn("<w:endnotes", en)
            self.assertIn('<w:endnote w:id="1"', en)
            self.assertIn("<w:endnoteRef", en)
            self.assertNotIn("w:footnote", en)

            # document.xml references rewritten
            doc = part(path, "word/document.xml")
            self.assertIn("w:endnoteReference", doc)
            self.assertNotIn("w:footnoteReference", doc)

            # rels rewired
            rels = part(path, "word/_rels/document.xml.rels")
            self.assertIn("endnotes.xml", rels)
            self.assertNotIn("footnotes.xml", rels)
            self.assertIn("relationships/endnotes", rels)

            # content types rewired
            ct = part(path, "[Content_Types].xml")
            self.assertIn("/word/endnotes.xml", ct)
            self.assertIn("wordprocessingml.endnotes+xml", ct)
            self.assertNotIn("/word/footnotes.xml", ct)

    def test_noop_when_no_footnotes_part(self):
        # If there are no footnotes, the function must be a no-op.
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "t.docx"
            write_docx(path, DOC_XML)  # no footnotes_xml argument
            footnotes_to_endnotes(path)  # must not raise
            with zipfile.ZipFile(path) as zf:
                names = set(zf.namelist())
            self.assertNotIn("word/endnotes.xml", names)


if __name__ == "__main__":
    unittest.main()
