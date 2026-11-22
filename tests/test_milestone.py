"""End-to-end checks for the installed tagpdf command and the controlled report."""

import copy
import os
import re
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from docx import Document
from pypdf import PdfReader
from pypdf.generic import ContentStream, NameObject


SAMPLE = Path(__file__).resolve().parents[1] / "examples" / "Quarterly Report.docx"


def items(value):
    return value if isinstance(value, list) else [value]


def descendants(node):
    yield node
    for child in items(node.get("/K", [])):
        obj = child.get_object()
        if isinstance(obj, dict):
            yield from descendants(obj)


def tagged_text(reader, node):
    """Read the actual text paints belonging to this element's marked content."""
    by_page = {}
    for child in descendants(node):
        if "/MCID" in child:
            page_num = reader.get_page_number(child["/Pg"])
            by_page.setdefault(page_num, set()).add(child["/MCID"])
    chunks = []
    for page_num, mcids in by_page.items():
        page = copy.copy(reader.pages[page_num])
        stream = ContentStream(None, reader)
        stack = []
        operations = []
        for args, op in ContentStream(page["/Contents"], reader, "bytes").operations:
            if op in (b"BMC", b"BDC"):
                props = args[1] if len(args) > 1 else {}
                stack.append(props.get("/MCID") if isinstance(props, dict) else None)
            elif op == b"EMC":
                stack.pop()
            if op in (b"Tj", b"TJ", b"'", b'"', b"Do") and not mcids.intersection(stack):
                continue
            operations.append((args, op))
        stream.operations = operations
        page[NameObject("/Contents")] = stream
        chunks.append(page.extract_text())
    return " ".join(" ".join(chunks).split())


def table_rows(reader, table):
    rows = []
    for row in descendants(table):
        if row.get("/S") == "/TR":
            rows.append([tagged_text(reader, cell) for cell in descendants(row)
                         if cell.get("/S") in ("/TD", "/TH")])
    return rows


def tagpdf_command(source):
    executable = Path(sys.executable).with_name("tagpdf.exe" if os.name == "nt" else "tagpdf")
    return subprocess.run(
        [executable if executable.is_file() else "tagpdf", source],
        cwd=source.parent, text=True, capture_output=True, timeout=180,
    )


class MilestoneTest(unittest.TestCase):
    def test_quarterly_report(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / SAMPLE.name
            shutil.copyfile(SAMPLE, source)
            result = tagpdf_command(source)
            self.assertEqual(result.returncode, 2, result.stdout + result.stderr)
            self.assertEqual(source.read_bytes(), SAMPLE.read_bytes())
            pdf = source.with_suffix(".pdf")
            review = source.with_suffix(".review.txt")
            self.assertEqual(set(source.parent.iterdir()), {source, pdf, review})
            self.check_outputs(pdf, review)

            # Protect each existing output even if its sibling is absent.
            pdf.unlink()
            review.write_text("employee review", encoding="utf-8")
            self.assertEqual(tagpdf_command(source).returncode, 1)
            self.assertEqual(review.read_text(encoding="utf-8"), "employee review")
            self.assertFalse(pdf.exists())
            review.unlink()
            pdf.write_bytes(b"employee-repaired PDF")
            self.assertEqual(tagpdf_command(source).returncode, 1)
            self.assertEqual(pdf.read_bytes(), b"employee-repaired PDF")
            self.assertFalse(review.exists())
            self.assertEqual(source.read_bytes(), SAMPLE.read_bytes())

    def test_described_map_passes_automated_validation(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / SAMPLE.name
            doc = Document(SAMPLE)
            doc.inline_shapes[-1]._inline.docPr.set("descr", "The office and library share a route.")
            doc.save(source)
            result = tagpdf_command(source)
            report_path = source.with_suffix(".review.txt")
            report = report_path.read_text(encoding="utf-8") if report_path.exists() else ""
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr + report)
            self.assertIn("Automated validation: PDF/UA-2 — PASS", report)
            figures = [n for n in descendants(PdfReader(source.with_suffix(".pdf")).root_object["/StructTreeRoot"])
                       if n.get("/S") == "/Figure"]
            self.assertIn("The office and library share a route.", [f.get("/Alt") for f in figures])

    def check_outputs(self, pdf, review):
        reader = PdfReader(pdf)
        tree = reader.root_object["/StructTreeRoot"]
        nodes = list(descendants(tree))
        role_map = tree["/RoleMap"]
        roles = {role_map.get(node["/S"], node["/S"]) for node in nodes if "/S" in node}
        self.assertTrue({"/Table", "/Figure", "/Link", "/L", "/LI", "/Caption"} <= roles, roles)
        self.assertTrue("/H1" in roles)
        figures = [node for node in nodes if node.get("/S") == "/Figure"]
        self.assertEqual(len(figures), 2)
        self.assertEqual(figures[0].get("/Alt"),
                         "Opened and closed service requests both increased in the second quarter.")
        self.assertIsNone(figures[1].get("/Alt"))
        captions = [tagged_text(reader, node) for figure in figures
                    for node in descendants(figure) if node.get("/S") == "/Caption"]
        self.assertEqual(captions, ["Figure 1. Requests by quarter", "Figure 2. Service area map"])
        tables = [node for node in nodes if node.get("/S") == "/Table"]
        spending = tables[0]
        self.assertEqual(table_rows(reader, spending), [
            ["Program", "FY 2025"], ["Budget", "Actual"],
            ["Outreach", "125", "108"], ["Training", "84", "79"],
        ])
        cells = [n for n in descendants(spending) if n.get("/S") in ("/TH", "/TD")]
        by_id = {cell["/ID"]: tagged_text(reader, cell) for cell in cells if cell["/S"] == "/TH"}
        for name, span in [("Program", "/RowSpan"), ("FY 2025", "/ColSpan")]:
            header = next(cell for cell in cells if tagged_text(reader, cell) == name)
            class_map = tree["/ClassMap"]
            attrs = [class_map[k] for k in items(header.get("/C", []))]
            attrs.extend(v.get_object() for v in items(header.get("/A", [])))
            self.assertIn(2, [attr.get(span) for attr in attrs], name)
        actual = next(cell for cell in cells if tagged_text(reader, cell) == "108")
        attributes = [v.get_object() for v in items(actual["/A"])]
        headers = next(attr["/Headers"] for attr in attributes if "/Headers" in attr)
        self.assertCountEqual([by_id[header] for header in headers], ["FY 2025", "Actual", "Outreach"])

        chart_rows = [table_rows(reader, table) for table in tables[1:]]
        self.assertEqual(chart_rows, [
            [["Point", "Category", "Opened"], ["1", "Q1", "31"], ["2", "Q2", "42"]],
            [["Point", "Category", "Closed"], ["1", "Q1", "28"], ["2", "Q2", "39"]],
        ])
        document_text = " ".join(" ".join(page.extract_text().split()) for page in reader.pages)
        for phrase in ["Opened", "Closed", "Office staff handle intake.", "service directory"]:
            self.assertIn(phrase, document_text)
        self.assertNotIn("Needs human review", document_text)
        uris = [annot.get_object()["/A"].get("/URI") for page in reader.pages
                for annot in page.get("/Annots", []) if "/A" in annot.get_object()]
        self.assertIn("https://example.org/service", uris)

        # Locate the actual rendered map XObject independently of the review report.
        map_pages = []
        for i, page in enumerate(reader.pages, 1):
            objects = page["/Resources"].get("/XObject", {})
            for args, op in page.get_contents().operations:
                if op == b"Do":
                    image = objects.get_object()[args[0]]
                    if (image.get("/Width"), image.get("/Height")) == (540, 132):
                        map_pages.append(i)
        self.assertEqual(len(map_pages), 1)
        report = review.read_text(encoding="utf-8")
        self.assertIn("Needs human review", report)
        self.assertIn("PDF/UA-2", report)
        self.assertRegex(report, r"(?i)veraPDF")
        self.assertIn("Conversion: No content loss detected.", report)
        self.assertIn("Automated validation: PDF/UA-2 — FAIL", report)
        rules = [line for line in report.splitlines() if "§" in line]
        self.assertEqual(len(rules), 1, report)
        self.assertIn("§8.2.5.28.2, test 1 (1 failed)", rules[0])
        image_issue = [line.lower() for line in report.splitlines()
                       if "image" in line.lower() and re.search(r"descri|alt text", line, re.I)]
        self.assertTrue(any(f"page {map_pages[0]}" in line and "service area" in line
                            for line in image_issue), report)


if __name__ == "__main__":
    unittest.main()
