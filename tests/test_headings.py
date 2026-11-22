"""Numbered Word headings must remain headings in the generated PDF."""

from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from docx import Document
from pypdf import PdfReader

from test_milestone import descendants, tagged_text, tagpdf_command


class HeadingTest(unittest.TestCase):
    def test_numbered_headings_keep_heading_roles_and_separate_list(self):
        doc = Document()
        heading = doc.styles["Heading 2"]
        heading.element.get_or_add_pPr().get_or_add_numPr().get_or_add_numId().val = 5
        doc.add_heading("Text Formatting", level=1)
        doc.add_heading("Are lists formatted correctly?", level=2)
        doc.add_paragraph("Select a list style.", style="List Bullet")
        doc.add_paragraph("Enter a list item.", style="List Bullet")
        doc.add_heading("Are columns formatted correctly?", level=2)

        with TemporaryDirectory() as temp:
            source = Path(temp) / "Numbered Headings.docx"
            doc.save(source)
            result = tagpdf_command(source)
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

            pdf = PdfReader(source.with_suffix(".pdf"))
            tree = pdf.root_object["/StructTreeRoot"]
            role_map = tree["/RoleMap"]
            actual = []
            for node in descendants(tree):
                role = role_map.get(node.get("/S"), node.get("/S"))
                if role in ("/H1", "/H2", "/LBody"):
                    actual.append((role, tagged_text(pdf, node)))
            self.assertEqual(actual, [
                ("/H1", "Text Formatting"),
                ("/H2", "1. Are lists formatted correctly?"),
                ("/LBody", "Select a list style."),
                ("/LBody", "Enter a list item."),
                ("/H2", "2. Are columns formatted correctly?"),
            ])
