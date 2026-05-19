"""Unit tests for the rsid/paraId stripping pass."""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from tests._helpers import REPO_ROOT, W_DOC_HEAD, W_DOC_TAIL, write_docx, part  # noqa: F401
from rtf_to_docx_endnotes import _strip_rsids_in_xml, strip_rsids_from_docx


class TestStripRsidsInXml(unittest.TestCase):
    def test_removes_known_rsid_attrs(self):
        xml = (
            W_DOC_HEAD
            + '<w:body>'
            + '<w:p w14:paraId="A1B2" w14:textId="C3D4" '
              'w:rsidR="00112233" w:rsidRDefault="44556677" w:rsidRPr="DEADBEEF">'
            + '<w:r w:rsidR="11223344" w:rsidRPr="99887766">'
            + '<w:t xml:space="preserve">hello</w:t></w:r></w:p>'
            + '</w:body>' + W_DOC_TAIL
        )
        out_bytes, removed = _strip_rsids_in_xml(xml.encode("utf-8"))
        out = out_bytes.decode("utf-8")

        self.assertEqual(removed, 7)
        for attr in ("paraId", "textId", "rsidR", "rsidRDefault", "rsidRPr"):
            self.assertNotIn(attr, out)
        # Content preserved
        self.assertIn("<w:t", out)
        self.assertIn("hello", out)

    def test_no_rsids_is_noop(self):
        xml = (
            W_DOC_HEAD
            + '<w:body><w:p><w:r><w:t>plain</w:t></w:r></w:p></w:body>'
            + W_DOC_TAIL
        )
        out_bytes, removed = _strip_rsids_in_xml(xml.encode("utf-8"))
        self.assertEqual(removed, 0)
        # Bytes returned unchanged when nothing to strip.
        self.assertEqual(out_bytes, xml.encode("utf-8"))

    def test_malformed_xml_passthrough(self):
        garbage = b"<not xml at all <<>>"
        out, removed = _strip_rsids_in_xml(garbage)
        self.assertEqual(removed, 0)
        self.assertEqual(out, garbage)


class TestStripRsidsFromDocx(unittest.TestCase):
    def test_in_place_strip(self):
        doc_xml = (
            W_DOC_HEAD
            + '<w:body>'
            + '<w:p w14:paraId="ABCD" w:rsidR="11">'
            + '<w:r w:rsidRPr="22"><w:t>x</w:t></w:r></w:p>'
            + '</w:body>' + W_DOC_TAIL
        )
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "t.docx"
            write_docx(path, doc_xml)
            removed = strip_rsids_from_docx(path)
            self.assertEqual(removed, 3)
            after = part(path, "word/document.xml")
            self.assertNotIn("paraId", after)
            self.assertNotIn("rsidR", after)
            self.assertNotIn("rsidRPr", after)
            self.assertIn("<w:t", after)

    def test_also_strips_in_footnotes(self):
        doc_xml = W_DOC_HEAD + "<w:body><w:p><w:r><w:t>x</w:t></w:r></w:p></w:body>" + W_DOC_TAIL
        fn_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
            '<w:footnote w:id="1">'
            '<w:p w14:paraId="DEAD" w:rsidR="BEEF"><w:r><w:t>note</w:t></w:r></w:p>'
            '</w:footnote></w:footnotes>'
        )
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "t.docx"
            write_docx(path, doc_xml, footnotes_xml=fn_xml)
            removed = strip_rsids_from_docx(path)
            self.assertEqual(removed, 2)
            after_fn = part(path, "word/footnotes.xml")
            self.assertNotIn("paraId", after_fn)
            self.assertNotIn("rsidR", after_fn)
            self.assertIn("note", after_fn)


if __name__ == "__main__":
    unittest.main()
