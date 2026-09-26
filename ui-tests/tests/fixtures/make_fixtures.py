"""Regenerate the documents the integration tests open.

The tests assert on the text written here, so change both together.
Needs python-docx and python-pptx: `python make_fixtures.py` in this folder.
"""
from pathlib import Path

from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from pptx import Presentation

HERE = Path(__file__).parent


def add_link(paragraph, text, target):
    """Append a hyperlink run; python-docx has no public API for it"""
    rel = paragraph.part.relate_to(target, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    link = OxmlElement("w:hyperlink")
    link.set(qn("r:id"), rel)
    run = OxmlElement("w:r")
    label = OxmlElement("w:t")
    label.text = text
    run.append(label)
    link.append(run)
    paragraph._p.append(link)


def make_docx():
    doc = Document()
    doc.add_heading("Quarterly Report", level=1)
    para = doc.add_paragraph("Revenue grew in ")
    para.add_run("every region").bold = True
    table = doc.add_table(rows=2, cols=2)
    for row, (region, revenue) in zip(table.rows, [("Region", "Revenue"), ("North", "120")]):
        row.cells[0].text, row.cells[1].text = region, revenue
    links = doc.add_paragraph()
    add_link(links, "Jupyter website", "https://jupyter.org")
    links.add_run(" and ")
    add_link(links, "Unsafe link", "javascript:alert(1)")
    doc.add_page_break()
    doc.add_heading("Appendix", level=1)
    doc.add_paragraph("Second page text.")
    doc.save(HERE / "sample.docx")


def make_altchunk_docx():
    """A DOCX embedding an HTML part whose script marks the page if it runs"""
    doc = Document()
    doc.add_paragraph("Text around the chunk")
    html = (
        "<html><body><p>Embedded chunk text</p>"
        "<script>parent.document.body.dataset.altchunk = 'ran'</script>"
        "</body></html>"
    ).encode()
    part = Part(PackURI("/word/chunk.html"), "text/html", html, doc.part.package)
    chunk = OxmlElement("w:altChunk")
    chunk.set(qn("r:id"), doc.part.relate_to(part, RELATIONSHIP_TYPE.A_F_CHUNK))
    doc.paragraphs[-1]._p.addnext(chunk)
    doc.save(HERE / "altchunk.docx")


def make_pptx():
    deck = Presentation()
    texts = [
        ("Alpha slide", "First slide"),
        ("Beta slide", "Second slide"),
        ("Gamma slide", "The zebra appears only here"),
    ]
    for title, subtitle in texts:
        slide = deck.slides.add_slide(deck.slide_layouts[0])
        slide.shapes.title.text = title
        slide.placeholders[1].text = subtitle
    deck.save(HERE / "sample.pptx")


def make_others():
    (HERE / "sample.rtf").write_text(
        r"{\rtf1\ansi\deff0{\fonttbl{\f0 Arial;}}\f0\fs24 Plain text and {\b bold text} in RTF.\par}"
    )
    # OLE compound-file signature: what a real .doc or .ppt starts with
    ole = bytes.fromhex("d0cf11e0a1b11ae1") + bytes(504)
    (HERE / "legacy.doc").write_bytes(ole)
    (HERE / "legacy.ppt").write_bytes(ole)
    (HERE / "broken.docx").write_bytes(b"This is not a DOCX file.")
    (HERE / "empty.pptx").write_bytes(b"")


if __name__ == "__main__":
    make_docx()
    make_altchunk_docx()
    make_pptx()
    make_others()
