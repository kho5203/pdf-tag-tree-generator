"""Check Word footer accounting and repaginated PDF artifacts."""

from copy import copy
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from docx import Document
from docx.oxml import parse_xml
from pypdf import PdfReader
from pypdf.generic import ContentStream, NameObject

from test_milestone import descendants, tagged_text, tagpdf_command


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def artifact_text(page):
    stream = ContentStream(None, page.pdf)
    operations = []
    stack = []
    for args, op in ContentStream(page["/Contents"], page.pdf, "bytes").operations:
        if op in (b"BMC", b"BDC"):
            stack.append(args[0])
        elif op == b"EMC":
            stack.pop()
        if op in (b"Tj", b"TJ", b"'", b'"') and "/Artifact" not in stack:
            continue
        operations.append((args, op))
    stream.operations = operations
    artifact_page = copy(page)
    artifact_page[NameObject("/Contents")] = stream
    return artifact_page.extract_text().strip()


class FooterTest(unittest.TestCase):
    def test_running_title_and_page_field_are_already_accounted_for(self):
        doc = Document()
        doc.add_paragraph("Example Guide", style="Title")
        doc.add_page_break()
        doc.add_paragraph("Read this guidance.")
        section = doc.sections[0]
        section.different_first_page_header_footer = True
        footer = section.footer.paragraphs[0]
        footer.text = "Example Guide\tPage | "
        footer.add_run()._r.addnext(parse_xml(rf"""
            <w:r xmlns:w="{W}">
              <w:fldChar w:fldCharType="begin"/><w:instrText> PAGE \* MERGEFORMAT </w:instrText>
              <w:fldChar w:fldCharType="separate"/><w:t>987</w:t><w:fldChar w:fldCharType="end"/>
            </w:r>"""))

        with TemporaryDirectory() as directory:
            source = Path(directory) / "Guide.docx"
            doc.save(source)
            result = tagpdf_command(source)
            review = source.with_suffix(".review.txt").read_text(encoding="utf-8")
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr + review)
            self.assertIn("Conversion: No content loss detected.", review)
            pdf = PdfReader(source.with_suffix(".pdf"))
            self.assertEqual(len(pdf.pages), 2)
            self.assertEqual(pdf.pages[1].extract_text().splitlines(), ["Read this guidance.", "2"])
            self.assertEqual([artifact_text(page) for page in pdf.pages], ["", "2"])
            self.assertEqual(sum(page.extract_text().count("Example Guide") for page in pdf.pages), 1)
            tagged = tagged_text(pdf, pdf.root_object["/StructTreeRoot"])
            self.assertNotIn("987", tagged)
            self.assertNotIn("2", tagged)

    def test_unique_first_footer_is_tagged_and_unknown_fields_are_reported(self):
        doc = Document()
        doc.add_paragraph("Report body.")
        section = doc.sections[0]
        section.different_first_page_header_footer = True
        section.first_page_footer.paragraphs[0].text = "For assistance contact the service desk."
        section.even_page_footer.paragraphs[0].text = "Inactive even-page footer."
        footer = section.footer.paragraphs[0]
        footer._p.append(parse_xml(f'<w:fldSimple xmlns:w="{W}" w:instr="NUMPAGES">'
                                   '<w:r><w:t>999</w:t></w:r></w:fldSimple>'))

        with TemporaryDirectory() as directory:
            source = Path(directory) / "Contacts.docx"
            doc.save(source)
            result = tagpdf_command(source)
            review = source.with_suffix(".review.txt").read_text(encoding="utf-8")
            self.assertEqual(result.returncode, 2, result.stdout + result.stderr + review)
            self.assertIn("unsupported field “NUMPAGES”", review)
            pdf = PdfReader(source.with_suffix(".pdf"))
            tree = pdf.root_object["/StructTreeRoot"]
            nodes = [tagged_text(pdf, node) for node in descendants(tree) if node.get("/S") == "/text"]
            self.assertEqual(nodes.count("For assistance contact the service desk."), 1)
            self.assertNotIn("Inactive even-page footer.", nodes)
            self.assertNotIn("999", pdf.pages[0].extract_text())
