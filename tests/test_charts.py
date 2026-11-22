"""Regression cases for data omissions and chart cache indexing."""

import unittest
from xml.etree.ElementTree import fromstring

from tagpdf.charts import C, render_chart


class ChartContext:
    label = "Chart 1"
    anchor = ""
    caption = ""
    text = staticmethod(str)

    def __init__(self):
        self.reviews = []
        self.losses = []

    def review(self, action):
        self.reviews.append(action)

    def loss(self, action):
        self.losses.append(action)


def convert(data, kind="barChart", properties=""):
    root = fromstring(f'<chartSpace xmlns="{C[1:-1]}"><chart><plotArea>'
                      f'<{kind}>{properties}<ser>{data}</ser></{kind}></plotArea></chart></chartSpace>')
    ctx = ChartContext()
    return render_chart(root, ctx, description="Source description"), ctx


class ChartTests(unittest.TestCase):
    def test_bar_direction_preserves_orientation_and_tagged_table_data(self):
        data = """
            <cat><strLit><pt idx="0"><v>Category 1</v></pt>
                         <pt idx="1"><v>Category 2</v></pt></strLit></cat>
            <val><numLit><pt idx="0"><v>-2</v></pt>
                         <pt idx="1"><v>4</v></pt></numLit></val>
        """
        horizontal, ctx = convert(data, properties='<barDir val="bar"/>')
        vertical, _ = convert(data, properties='<barDir val="col"/>')
        # The negative bar starts left of zero and the positive bar starts at zero.
        self.assertIn(r"\put(72.000,37.000){\rule{56.000pt}{10.000pt}}", horizontal)
        self.assertIn(r"\put(128.000,19.000){\rule{112.000pt}{10.000pt}}", horizontal)
        self.assertIn(r"\rule{18.000pt}{28.000pt}", vertical)
        for tex in (horizontal, vertical):
            self.assertIn("1 & Category 1 & -2", tex)
            self.assertIn("2 & Category 2 & 4", tex)
        self.assertEqual(ctx.losses, [])

    def test_sparse_index_preserves_category_relationship_and_custom_label(self):
        tex, ctx = convert("""
            <tx><v>Cases</v></tx>
            <cat><multiLvlStrRef><multiLvlStrCache>
                <lvl><pt idx="0"><v>Q1</v></pt><pt idx="1"><v>Q2</v></pt></lvl>
                <lvl><pt idx="0"><v>2025</v></pt></lvl>
            </multiLvlStrCache></multiLvlStrRef></cat>
            <val><numLit><pt idx="1"><v>23</v></pt></numLit></val>
            <dLbls><dLbl><idx val="1"/><tx><v>Peak</v></tx></dLbl></dLbls>
        """)
        self.assertIn("1 & 2025 / Q1 &  & ", tex)
        self.assertIn("2 & 2025 / Q2 & 23 & Peak", tex)
        self.assertTrue(ctx.reviews)
        self.assertEqual(ctx.losses, [])

    def test_formula_only_values_report_incomplete(self):
        _, ctx = convert("<val><numRef><f>Sheet1!A1:A4</f></numRef></val>")
        self.assertTrue(any("no usable cached" in entry for entry in ctx.losses))

    def test_unreconstructed_plot_preserves_xy_and_bubble_size(self):
        def numeric(tag, value):
            return f'<{tag}><numLit><pt idx="0"><v>{value}</v></pt></numLit></{tag}>'

        tex, ctx = convert(numeric("xVal", 30) + numeric("yVal", 2) + numeric("bubbleSize", 8),
                           kind="bubbleChart")
        self.assertIn("1 & 30 & 2 & 8", tex)
        self.assertTrue(any("bubbleChart plot was omitted" in message for message in ctx.losses))
