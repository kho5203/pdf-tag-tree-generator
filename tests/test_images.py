"""Verify embedded GIF pixels and reading order through the installed CLI."""

from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from docx import Document
from PIL import Image
from pypdf import PdfReader

from test_milestone import descendants, tagged_text, tagpdf_command


class ImageTest(unittest.TestCase):
    def test_animated_gif_reports_missing_frames(self):
        buffer = BytesIO()
        first = Image.new("RGB", (16, 12), "red")
        first.save(buffer, format="GIF", save_all=True,
                   append_images=[Image.new("RGB", first.size, "blue")], duration=200)
        buffer.seek(0)
        doc = Document()
        picture = doc.add_picture(buffer)
        picture._inline.docPr.set("descr", "Rectangle changes color.")

        with TemporaryDirectory() as temp:
            source = Path(temp) / "Animation.docx"
            doc.save(source)
            result = tagpdf_command(source)
            review = source.with_suffix(".review.txt").read_text(encoding="utf-8")
            self.assertEqual(result.returncode, 2, result.stdout + result.stderr + review)
            self.assertIn("Conversion: INCOMPLETE", review)
            self.assertIn("Only the first frame of this animated GIF was preserved.", review)
            self.assertIn("Automated validation: PDF/UA-2 — PASS", review)
            page = PdfReader(source.with_suffix(".pdf")).pages[0]
            image, = page.images
            self.assertEqual(image.image.convert("RGB").tobytes(), first.tobytes())

    def test_static_gif_is_painted_once_with_source_description(self):
        buffer = BytesIO()
        source_pixels = Image.new("RGBA", (16, 12), (0, 0, 0, 0))
        source_pixels.paste((255, 128, 0, 255), (3, 2, 13, 10))
        source_pixels.save(buffer, format="GIF")
        buffer.seek(0)
        doc = Document()
        doc.add_paragraph("Before the symbol.")
        picture = doc.add_picture(buffer)
        picture._inline.docPr.set("descr", "Orange rectangle on a transparent background.")
        doc.add_paragraph("After the symbol.")

        with TemporaryDirectory() as temp:
            source = Path(temp) / "Symbol.docx"
            doc.save(source)
            result = tagpdf_command(source)
            review = source.with_suffix(".review.txt").read_text(encoding="utf-8")
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr + review)
            self.assertIn("Conversion: No content loss detected.", review)

            pdf = PdfReader(source.with_suffix(".pdf"))
            page = pdf.pages[0]
            paints = [args[0] for args, op in page.get_contents().operations if op == b"Do"]
            self.assertEqual(len(paints), 1)
            image = page.images[paints[0]].image.convert("RGBA")
            self.assertEqual(image.size, source_pixels.size)
            self.assertEqual(image.tobytes(), source_pixels.tobytes())

            nodes = list(descendants(pdf.root_object["/StructTreeRoot"]))
            figures = [n for n in nodes if n.get("/S") == "/Figure"]
            self.assertEqual([n.get("/Alt") for n in figures], ["Orange rectangle on a transparent background."])
            role_map = pdf.root_object["/StructTreeRoot"]["/RoleMap"]
            paragraphs = [n for n in nodes if role_map.get(n.get("/S"), n.get("/S")) == "/P"]
            before = next(n for n in paragraphs if tagged_text(pdf, n) == "Before the symbol.")
            after = next(n for n in paragraphs if tagged_text(pdf, n) == "After the symbol.")
            positions = [n.indirect_reference.idnum for n in nodes if n.get("/Type") == "/StructElem"]
            before, figure, after = [positions.index(n.indirect_reference.idnum) for n in [before, *figures, after]]
            self.assertLess(before, figure)
            self.assertLess(figure, after)
