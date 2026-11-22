"""Regenerate the synthetic Word sample with python-docx, python-pptx, and Pillow."""

from io import BytesIO
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.opc.constants import CONTENT_TYPE as CT, RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part, XmlPart
from docx.shared import Inches, Pt, RGBColor
from lxml import etree
from PIL import Image, ImageDraw
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE


SAMPLE = Path(__file__).resolve().parents[1] / "examples" / "Quarterly Report.docx"


def add_chart(doc):
    data = CategoryChartData()
    data.categories = ["Q1", "Q2"]
    data.add_series("Opened", (31, 42))
    data.add_series("Closed", (28, 39))
    chart_xml = etree.fromstring(data.xml_bytes(XL_CHART_TYPE.COLUMN_CLUSTERED))
    chart_xml.find(".//" + qn("a:defRPr")).set("sz", "1000")
    chart = chart_xml.find(qn("c:chart"))
    plot_area = chart.find(qn("c:plotArea"))
    legend = etree.Element(qn("c:legend"))
    etree.SubElement(legend, qn("c:legendPos"), val="b")
    plot_area.addnext(legend)
    chart_part = XmlPart(PackURI("/word/charts/chart1.xml"), CT.DML_CHART, chart_xml, doc.part.package)
    workbook_part = Part(
        PackURI("/word/embeddings/requests.xlsx"),
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        data.xlsx_blob,
        doc.part.package,
    )
    external_data = etree.SubElement(chart_xml, qn("c:externalData"))
    external_data.set(qn("r:id"), chart_part.relate_to(workbook_part, RT.PACKAGE))
    etree.SubElement(external_data, qn("c:autoUpdate"), val="0")
    relationship = doc.part.relate_to(chart_part, RT.CHART)
    drawing = parse_xml(f"""
        <w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
          xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <wp:inline>
            <wp:extent cx="{Inches(5.5)}" cy="{Inches(1.7)}"/>
            <wp:docPr id="1" name="Requests by quarter"
              descr="Opened and closed service requests both increased in the second quarter."/>
            <wp:cNvGraphicFramePr/>
            <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart r:id="{relationship}"/>
            </a:graphicData></a:graphic>
          </wp:inline>
        </w:drawing>""")
    doc.add_paragraph().add_run()._r.append(drawing)


def add_table(doc):
    table = doc.add_table(rows=4, cols=3)
    table.style = "Table Grid"
    table.cell(0, 0).merge(table.cell(1, 0)).text = "Program"
    table.cell(0, 1).merge(table.cell(0, 2)).text = "FY 2025"
    table.cell(1, 1).text = "Budget"
    table.cell(1, 2).text = "Actual"
    for row in table.rows[:2]:
        row._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))
        for cell in row.cells:
            for run in cell.paragraphs[0].runs:
                run.bold = True
    look = table._tbl.tblPr.find(qn("w:tblLook"))
    look.set(qn("w:firstColumn"), "1")
    for row, values in zip(table.rows[2:], [("Outreach", "125", "108"), ("Training", "84", "79")]):
        for cell, value in zip(row.cells, values):
            cell.text = value


def add_map(doc):
    # A two-location schematic. Its missing alt text is intentional.
    image = Image.new("RGB", (540, 132), "#f0f4f3")
    draw = ImageDraw.Draw(image)
    draw.line([(0, 80), (160, 52), (330, 88), (540, 52)], fill="#b7ced7", width=16)
    draw.line([(84, 75), (254, 32), (445, 65)], fill="#425448", width=3)
    for x, y, label in [(84, 75, "Office"), (445, 65, "Library")]:
        draw.ellipse((x-6, y-6, x+6, y+6), fill="#1f4941", outline="white")
        draw.text((x-26, y+11), label, fill="#23322b", font_size=18)
    png = BytesIO()
    image.save(png, format="PNG")
    picture = doc.add_picture(png, width=Inches(3.6))
    picture._inline.docPr.set("name", "Service area map")


def build(path):
    doc = Document()
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = Inches(0.65)
    normal = doc.styles["Normal"]
    normal.font.size = Pt(10)
    normal.paragraph_format.space_after = Pt(3)
    title = doc.styles["Title"]
    title.font.color.rgb = RGBColor(0, 0, 0)
    border = title.element.find(".//" + qn("w:pBdr"))
    if border is not None:
        border.getparent().remove(border)
    doc.core_properties.title = "Quarterly Report"
    doc.core_properties.author = "PDF Tag Tree Generator"
    doc.core_properties.language = "en-US"
    doc.add_paragraph("Quarterly Report", style="Title")
    doc.add_paragraph("Both programs stayed within budget as service requests increased.")
    doc.add_heading("Spending by year", level=1)
    doc.add_paragraph("Table 1. Spending by year", style="Caption")
    add_table(doc)
    doc.add_heading("Service requests", level=1)
    add_chart(doc)
    doc.add_paragraph("Figure 1. Requests by quarter", style="Caption")
    doc.add_heading("Service area", level=1)
    add_map(doc)
    doc.add_paragraph("Figure 2. Service area map", style="Caption")
    doc.add_paragraph("Office staff handle intake.", style="List Bullet")
    doc.add_paragraph("Library staff handle referrals.", style="List Bullet")
    paragraph = doc.add_paragraph("See the ")
    link = OxmlElement("w:hyperlink")
    link.set(qn("r:id"), doc.part.relate_to("https://example.org/service", RT.HYPERLINK, is_external=True))
    run = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.text = "service directory"
    run.append(text)
    link.append(run)
    paragraph._p.append(link)
    paragraph.add_run(".")
    doc.save(path)


if __name__ == "__main__":
    build(SAMPLE)
