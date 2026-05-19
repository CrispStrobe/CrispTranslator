"""Shared test helpers: make repo-root scripts importable and build tiny
docx fixtures in-memory without needing pandoc or textutil."""

from __future__ import annotations

import io
import sys
import zipfile
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))


# Minimal Open Packaging Conventions skeleton for a single-page docx.
_CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  {extra}
</Types>"""

_PKG_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

_DOC_RELS_BASE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{extra}</Relationships>"""


def make_minimal_docx(
    document_xml: str,
    footnotes_xml: str | None = None,
) -> bytes:
    """Build a docx (as bytes) containing the given document.xml and optional
    footnotes.xml. Returns the raw zip payload."""
    extra_ct = ""
    extra_rels = ""
    if footnotes_xml is not None:
        extra_ct = (
            '<Override PartName="/word/footnotes.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.'
            'wordprocessingml.footnotes+xml"/>'
        )
        extra_rels = (
            '<Relationship Id="rId10" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/'
            'relationships/footnotes" Target="footnotes.xml"/>'
        )

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", _CONTENT_TYPES.format(extra=extra_ct))
        zf.writestr("_rels/.rels", _PKG_RELS)
        zf.writestr(
            "word/_rels/document.xml.rels",
            _DOC_RELS_BASE.format(extra=extra_rels),
        )
        zf.writestr("word/document.xml", document_xml)
        if footnotes_xml is not None:
            zf.writestr("word/footnotes.xml", footnotes_xml)
    return buf.getvalue()


def write_docx(path: Path, document_xml: str, footnotes_xml: str | None = None) -> Path:
    path.write_bytes(make_minimal_docx(document_xml, footnotes_xml))
    return path


def part(docx_path: Path, name: str) -> str:
    with zipfile.ZipFile(docx_path, "r") as zf:
        return zf.read(name).decode("utf-8")


W_DOC_HEAD = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
)
W_DOC_TAIL = "</w:document>"
