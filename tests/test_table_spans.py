"""PDF regression for a cell spanning both axes and a forward header reference."""

from pathlib import Path
import tempfile
import unittest

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from pypdf import PdfReader

from test_milestone import descendants, items, tagged_text, tagpdf_command


class TableSpanTest(unittest.TestCase):
    def test_spanned_data_uses_later_row_header(self):
        doc = Document()
        doc.core_properties.language = "en-US"
        table = doc.add_table(rows=3, cols=4)
        values = [
            ["District", "January", "February", "Code"],
            ["North", "", "", "17"],
            ["South", "", "", "23"],
        ]
        for row, contents in zip(table.rows, values):
            for cell, content in zip(row.cells, contents):
                cell.text = content
        table.cell(1, 1).merge(table.cell(2, 2)).text = "Combined"
        table.rows[0]._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))
        table._tbl.tblPr.find(qn("w:tblLook")).set(qn("w:firstColumn"), "1")

        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "Districts.docx"
            doc.save(source)
            result = tagpdf_command(source)
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            pdf = PdfReader(source.with_suffix(".pdf"))
            tree = pdf.root_object["/StructTreeRoot"]
            cells = [o for o in descendants(tree) if o.get("/S") in ("/TH", "/TD")]
            self.assertEqual(len(cells), 9)
            names = {o["/ID"]: tagged_text(pdf, o) for o in cells}
            cell = next(o for o in cells if names[o["/ID"]] == "Combined")
            attrs = [o.get_object() for o in items(cell["/A"])]
            attrs.extend(tree["/ClassMap"][name] for name in items(cell["/C"]))
            self.assertIn(2, [a.get("/RowSpan") for a in attrs])
            self.assertIn(2, [a.get("/ColSpan") for a in attrs])
            headers = next(a["/Headers"] for a in attrs if "/Headers" in a)
            self.assertCountEqual([names[i] for i in headers], ["January", "February", "North", "South"])
