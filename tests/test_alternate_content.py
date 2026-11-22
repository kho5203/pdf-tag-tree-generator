"""Exercise branch selection and textbox fallback through the installed CLI."""

from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from docx import Document
from docx.oxml import parse_xml
from pypdf import PdfReader

from test_milestone import descendants, tagged_text, tagpdf_command


def alternate(content):
    return parse_xml(
        '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
        'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:drawing="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:unknown="urn:unsupported">'
        + content + '</mc:AlternateContent>'
    )


class AlternateContentTest(unittest.TestCase):
    def convert(self, doc, exit_code=0):
        with TemporaryDirectory() as directory:
            source = Path(directory) / "Compatibility Cover.docx"
            doc.save(source)
            result = tagpdf_command(source)
            self.assertEqual(result.returncode, exit_code, result.stdout + result.stderr)
            pdf = PdfReader(source.with_suffix(".pdf"))
            tree = pdf.root_object["/StructTreeRoot"]
            role_map = tree["/RoleMap"]
            tagged = [tagged_text(pdf, node) for node in descendants(tree)
                      if role_map.get(node.get("/S"), node.get("/S")) in ("/Title", "/P")]
            visible = " ".join(" ".join(page.extract_text().split()) for page in pdf.pages)
            return tagged, visible, source.with_suffix(".review.txt").read_text(encoding="utf-8")

    def test_selects_first_choice_with_all_required_namespaces(self):
        doc = Document()
        before = doc.add_paragraph("Keep this repeated line.")
        before._p.addnext(alternate("""
            <mc:Choice Requires="w unknown"><w:p><w:r><w:t>Unsupported</w:t></w:r></w:p></mc:Choice>
            <mc:Choice Requires="w drawing"><w:p><w:r><w:t>Selected choice</w:t></w:r></w:p></mc:Choice>
            <mc:Choice Requires="w"><w:p><w:r><w:t>Later choice</w:t></w:r></w:p></mc:Choice>
            <mc:Fallback><w:p><w:r><w:t>Unused fallback</w:t></w:r></w:p></mc:Fallback>
        """))
        doc.add_paragraph("Keep this repeated line.")

        tagged, visible, review = self.convert(doc)
        self.assertEqual(tagged, ["Keep this repeated line.", "Selected choice", "Keep this repeated line."])
        self.assertEqual(visible.count("Selected choice"), 1)
        self.assertEqual(visible.count("Keep this repeated line."), 2)
        self.assertIn("Conversion: No content loss detected.", review)

    def test_unsupported_choice_uses_textbox_fallback(self):
        doc = Document()
        run = doc.add_paragraph().add_run()
        run._r.append(alternate("""
            <mc:Choice Requires="unknown"><w:t>Unsupported version</w:t></mc:Choice>
            <mc:Fallback>
              <w:pict><v:roundrect><v:textbox><w:txbxContent>
                <w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr>
                  <w:r><w:t>Cover title</w:t></w:r></w:p>
                <w:p><w:r><w:t>Cover subtitle</w:t></w:r></w:p>
                <w:p><w:r><w:t>October 2018</w:t></w:r></w:p>
              </w:txbxContent></v:textbox></v:roundrect></w:pict>
            </mc:Fallback>
        """))
        doc.add_paragraph("Report body.")

        tagged, visible, review = self.convert(doc)
        self.assertEqual(tagged, ["Cover title", "Cover subtitle", "October 2018", "Report body."])
        for line in tagged:
            self.assertEqual(visible.count(line), 1)
        self.assertNotIn("Unsupported version", visible)
        self.assertIn("Conversion: No content loss detected.", review)

    def test_selected_branch_still_reports_unsupported_material(self):
        doc = Document()
        doc.add_paragraph("Before the object.").add_run()._r.append(alternate("""
            <mc:Choice Requires="w"><w:pict><v:shape><v:imagedata/></v:shape></w:pict></mc:Choice>
            <mc:Fallback><w:t>Unused fallback</w:t></mc:Fallback>
        """))

        _, visible, review = self.convert(doc, exit_code=2)
        self.assertNotIn("Unused fallback", visible)
        self.assertIn("Restore the unsupported legacy Word picture or shape from Word.", review)
