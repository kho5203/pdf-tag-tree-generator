"""Translate document-order WordprocessingML to the project's controlled template."""

from collections import Counter
from contextlib import contextmanager
from dataclasses import dataclass
from importlib.resources import files
from io import BytesIO
from pathlib import Path
import re
from urllib.parse import quote, urlsplit

from docx import Document
from docx.oxml.ns import qn
from lxml import etree
from PIL import Image

from .charts import render_chart
from .numbering import Numbering, val
from .tables import render_table

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
CHART = "{http://schemas.openxmlformats.org/drawingml/2006/chart}chart"
DECORATIVE = "{http://schemas.microsoft.com/office/drawing/2017/decorative}decorative"
LATEX = {c: "\\" + c for c in "&%$#_{}"}
LATEX.update({"\\": r"\textbackslash{}", "^": r"\textasciicircum{}", "~": r"\textasciitilde{}",
              "<": r"\textless{}", ">": r"\textgreater{}", "\u00a0": "~"})

def text(string):
    return "".join(LATEX.get(c, c if c >= " " else " ") for c in string)


def words(node):
    return "".join(t.text or "" for t in node.iter(W+"t")).strip()


def local(node):
    return etree.QName(node).localname


@dataclass(frozen=True)
class Finding:
    marker: int
    label: str
    action: str
    lost: bool


class Report:
    text = staticmethod(text)

    def __init__(self, source: Path, work: Path):
        self.doc = Document(source)
        self.work = work
        self.title = self.doc.core_properties.title or source.stem
        self.styles = {s.style_id: s for s in self.doc.styles}
        self.numbering = Numbering(self.doc)
        self.counts = Counter()
        self.findings = []
        self.heading = ""
        self.label = "Document"
        self.marker = 0
        self.next_marker = 0
        self.caption = ""
        self.bookmarks = {node.get(W+"name"): f"bookmark{i}" for i, node in
                          enumerate(self.doc.element.body.iter(W+"bookmarkStart"))}
        lang_node = self.doc.styles.element.find(".//"+W+"docDefaults//"+W+"lang")
        lang = self.doc.core_properties.language or (lang_node.get(W+"val") if lang_node is not None else "")
        self.language = lang if re.fullmatch(r"[a-zA-Z]{2,8}(?:-[A-Za-z0-9]{1,8})*", lang or "") else "en-US"
        if self.language != lang:
            self.review("Confirm the document language in Acrobat. The template used en-US.")

    @property
    def anchor(self):
        return r"\TagPage{" + str(self.marker) + "}"

    def review(self, action):
        self.findings.append(Finding(self.marker, self.label, action, False))

    def loss(self, action):
        self.findings.append(Finding(self.marker, self.label, action, True))

    @contextmanager
    def location(self, kind, name="", caption=""):
        original = self.marker, self.label, self.caption
        self.counts[kind] += 1
        self.next_marker += 1
        self.marker = self.next_marker
        self.label = f"{kind} {self.counts[kind]}"
        if name:
            self.label += f", ‘{name}’"
        if self.heading and self.heading != name:
            self.label += f" under ‘{self.heading}’"
        self.caption = caption
        try:
            yield
        finally:
            self.marker, self.label, self.caption = original

    def properties(self, p):
        props = [p.find(W+"pPr")]
        style = self.styles.get(val(props[0], W+"pStyle", "Normal"))
        names = []
        while style is not None and style.style_id not in names:
            names.append(style.style_id)
            props.append(style.element.find(W+"pPr"))
            style = style.base_style
        return props, names

    def inline_paragraph(self, p):
        return "".join(self.inline(n) for n in p)

    def inline(self, n):
        tag = n.tag
        if tag == W+"t":
            return text(n.text or "")
        if tag in (W+"r", W+"hyperlink", W+"ins", W+"moveTo", W+"sdtContent", W+"smartTag", W+"fldSimple", W+"customXml"):
            content = "".join(self.inline(child) for child in n)
            if tag == W+"hyperlink":
                rid = n.get(qn("r:id"))
                rel = self.doc.part.rels.get(rid)
                mark = n.get(W+"anchor")
                if rel is not None and rel.is_external and urlsplit(rel.target_ref).scheme.lower() in {"https", "http", "mailto"}:
                    url = quote(rel.target_ref, safe=":/?&=+#%~@!$'()*,-._")
                    # Percent-encode braces, backslashes and TeX's superscript character before
                    # constructing the href. hyperref decodes escaped percent/hash characters.
                    url = url.replace("%", r"\%").replace("#", r"\#")
                    return r"\href{" + url + "}{" + content + "}"
                if not rid and mark in self.bookmarks:
                    return r"\hyperlink{" + self.bookmarks[mark] + "}{" + content + "}"
                self.loss("Restore the hyperlink destination in Acrobat or Word. Its visible text was retained.")
            if tag == W+"r" and content and n.find(W+"drawing") is None:
                props = n.find(W+"rPr")
                for key, macro in (("b", "textbf"), ("i", "textit")):
                    el = props.find(W+key) if props is not None else None
                    if el is not None and el.get(W+"val","1") not in {"false","0","off"}:
                        content = "\\" + macro + "{" + content + "}"
                vertical = val(props, W+"vertAlign")
                if vertical in ("superscript", "subscript"):
                    content = r"\text" + vertical + "{" + content + "}"
            return content
        if tag == W+"drawing":
            return self.drawing(n)
        if tag == W+"tab":
            return r"\quad{}"
        if tag in (W+"br", W+"cr"):
            return r"\newpage{}" if n.get(W+"type") == "page" else r"\newline{}"
        if tag == W+"noBreakHyphen":
            return r"\mbox{-}"
        if tag == W+"softHyphen":
            return r"\-"
        if tag == W+"bookmarkStart" and n.get(W+"name") in self.bookmarks:
            return r"\hypertarget{" + self.bookmarks[n.get(W+"name")] + "}{}"
        if tag == W+"sdt":
            body = n.find(W+"sdtContent")
            return self.inline_paragraph(body) if body is not None else ""
        if tag in {W+t for t in ("pPr","rPr","bookmarkEnd","proofErr","lastRenderedPageBreak",
                                 "instrText","fldChar","del","moveFrom","commentRangeStart","commentRangeEnd",
                                 "commentReference","permStart","permEnd","renderedPageBreak")}:
            return ""
        self.loss(f"Restore or compare unsupported Word content ({local(n)}) with the source.")
        return text(words(n))

    def drawing(self, n):
        prop = n.find(".//"+WP+"docPr")
        name = prop.get("name","") if prop is not None else ""
        description = prop.get("descr","").strip() if prop is not None else ""
        chart = n.find(".//"+CHART)
        caption = self.caption
        with self.location("Chart" if chart is not None else "Image", name, caption):
            if n.find(WP+"anchor") is not None:
                self.review("Confirm the floating object's position in the tag reading order.")
            if chart is not None:
                rel = self.doc.part.rels.get(chart.get(qn("r:id")))
                if rel is None or rel.is_external:
                    self.loss("Restore the native chart. Its chart part is missing or external.")
                    return ""
                root = etree.fromstring(rel.target_part.blob, etree.XMLParser(resolve_entities=False, no_network=True))
                return "\n\n" + render_chart(root, self, description=description) + "\n\n"
            data = n.find(".//"+A+"graphicData")
            images = n.findall(".//"+A+"blip")
            if data is None or data.get("uri") != "http://schemas.openxmlformats.org/drawingml/2006/picture" or len(images) != 1:
                self.loss("Restore the unsupported drawing or diagram from Word.")
                return text(words(n))
            rel = self.doc.part.rels.get(images[0].get(qn("r:embed")))
            if rel is None or rel.is_external:
                self.loss("Embed the linked image in Word and convert again.")
                return ""
            blob = rel.target_part.blob
            if blob.startswith((b"GIF87a", b"GIF89a")):
                try:
                    with Image.open(BytesIO(blob), formats=("GIF",)) as image:
                        with BytesIO() as png:
                            image.convert("RGBA").save(png, format="PNG")
                            blob = png.getvalue()
                        if image.is_animated:
                            self.loss("Only the first frame of this animated GIF was preserved. Restore the remaining information from Word.")
                except (OSError, Image.DecompressionBombError):
                    self.loss("The embedded GIF could not be decoded. Reinsert it as PNG in Word and convert again.")
                    return ""
            if blob.startswith(b"\x89PNG\r\n\x1a\n"):
                ext = "png"
            elif blob.startswith(b"\xff\xd8\xff"):
                ext = "jpg"
            else:
                self.loss("Reinsert this image as PNG or JPEG. Its original format was omitted.")
                return ""
            asset = f"picture-{self.marker}.{ext}"
            (self.work / asset).write_bytes(blob)
            decorative = any(d.get("val", "1") in {"1","true","on"} for d in n.iter(DECORATIVE))
            if decorative:
                option = "artifact"
            elif description:
                option = "alt={" + text(description) + "}"
            else:
                option = "tagging-setup={mode=undescribed}"
                self.review("Add a description (alt text) in Acrobat that conveys the image’s purpose, or mark it decorative after inspecting it.")
            if n.find(".//"+A+"srcRect") is not None or any(e.get("rot") or e.get("flipH") or e.get("flipV") for e in n.iter(A+"xfrm")):
                self.review("Compare the original crop and orientation with the full embedded image.")
            extent = n.find(".//"+WP+"extent")
            inches = min(6.5, max(.01,int(extent.get("cx")) / 914400)) if extent is not None else 6.5
            size = (rf"width=\dimexpr\ifdim{inches:.3f}in>\linewidth\linewidth"
                    rf"\else{inches:.3f}in\fi\relax,height=.65\textheight,keepaspectratio")
            if caption and not decorative:
                # Place the source caption directly under the Figure. The default
                # graphicx plug would create a second, nested Figure for the pixels.
                figure = "tag=Figure" + (",alt={" + text(description) + "}" if description else "")
                return (r"\par\tagstructbegin{" + figure + r"}\tagpdfparaOff\noindent\tagmcbegin{tag=Figure}"
                        + self.anchor + rf"\includegraphics[tagging-setup={{mode=off}},{size}]{{{asset}}}"
                        + r"\tagmcend\par\tagpdfparaOn\nopagebreak\tagstructbegin{tag=Caption}"
                        + r"{\small " + caption + r"\par}\tagstructend\tagstructend{}")
            graphic = r"\leavevmode" + self.anchor + rf"\includegraphics[{option},{size}]{{{asset}}}"
            if caption:
                return graphic + r"\par\nopagebreak " + caption
            return graphic

    def blocks(self, parent):
        for node in parent:
            if node.tag in {W+"sdt", W+"customXml", W+"ins", W+"moveTo"}:
                if node.tag == W+"sdt":
                    content = node.find(W+"sdtContent")
                    if content is not None:
                        yield from self.blocks(content)
                else:
                    yield from self.blocks(node)
            elif node.tag not in {W+"sectPr", W+"customXmlPr", W+"del", W+"moveFrom", W+"bookmarkStart", W+"bookmarkEnd"}:
                yield node

    def body(self):
        nodes = list(self.blocks(self.doc.element.body))
        captions = {}
        consumed = set()
        for i, n in enumerate(nodes):
            if n.tag != W+"p" or "Caption" not in self.properties(n)[1]:
                continue
            candidates = [j for j in (i-1, i+1) if 0 <= j < len(nodes) and
                          (nodes[j].tag == W+"tbl" or len(list(nodes[j].iter(W+"drawing"))) == 1)]
            # Conventionally table captions precede tables and figure captions follow drawings.
            preferred = [j for j in candidates if (nodes[j].tag == W+"tbl") == (j > i)]
            if len(preferred or candidates) == 1:
                j = (preferred or candidates)[0]
                if j not in captions:
                    captions[j] = n
                    consumed.add(i)
        result = []
        lists = []

        def close(depth):
            while len(lists) > depth:
                lists.pop()
                result.append(r"\end{itemize}")

        for i, node in enumerate(nodes):
            if i in consumed:
                continue
            cap = captions.get(i)
            caption = self.inline_paragraph(cap) if cap is not None else ""
            if node.tag == W+"tbl":
                close(0)
                name = words(cap) if cap is not None else val(node.find(W+"tblPr"), W+"tblCaption", "")
                with self.location("Table", name, caption):
                    result.append(render_table(node, self))
                continue
            if node.tag != W+"p":
                close(0)
                with self.location("Object", local(node)):
                    self.loss("Compare this unsupported Word object with the source and restore its content.")
                    if words(node):
                        result.append(self.anchor + text(words(node)) + r"\par")
                continue
            props, styles = self.properties(node)
            level = next((int(v) for p in props if (v := val(p,W+"outlineLvl")) is not None), 9)
            title = "Title" in styles
            if 0 <= level < 9:
                self.heading = words(node)
            with self.location("Paragraph", words(node)[:60], caption):
                item = self.numbering.item(props, self)
                content = self.inline_paragraph(node)
                if not content.strip():
                    continue
                if item is not None and (level >= 9 or title):
                    num, depth, label = item
                    if depth > len(lists) or depth >= 4:
                        self.review("Confirm the nesting depth of this list item in Acrobat.")
                    depth = min(depth, len(lists), 3)
                    close(depth + 1)
                    if len(lists) > depth and lists[depth] != num:
                        close(depth)
                    if len(lists) == depth:
                        result.append(r"\begin{itemize}")
                        lists.append(num)
                    result.append(r"\item[{" + text(label) + "}]" + self.anchor + content)
                    continue
                close(0)
                if level < 9 and not title:
                    commands = ["section", "subsection", "subsubsection", "paragraph", "subparagraph"]
                    if level >= len(commands):
                        self.review("Restore the original heading depth in the PDF tag tree.")
                    command = commands[min(level,len(commands)-1)]
                    if i+1 < len(nodes) and nodes[i+1].find(".//"+WP+"extent") is not None:
                        result.append(r"\Needspace{10\baselineskip}")
                    # Word can use list numbering on a paragraph with an outline
                    # level. Keep its numbering in the visible and tagged heading.
                    prefix = text(item[2]) + " " if item is not None and item[2] else ""
                    result.append("\\" + command + "[{" + prefix + text(words(node)) + "}]{" +
                                  prefix + content + "}" + self.anchor)
                elif title:
                    result.append(r"\par\tagpdfparaOff\noindent\tagstructbegin{tag=Title}\tagmcbegin{}"
                                  + r"{\LARGE\bfseries " + self.anchor + content
                                  + r"}\tagmcend\tagstructend\par\tagpdfparaOn")
                else:
                    if "Caption" in styles:
                        self.review("Confirm which object this caption describes and associate it in the tag tree.")
                    result.append(self.anchor + content + r"\par")
        close(0)
        for rel in self.doc.part.rels.values():
            if rel.reltype.rsplit("/",1)[-1] in {"footer","header","footnotes","endnotes"} and not rel.is_external:
                part = etree.fromstring(rel.target_part.blob, etree.XMLParser(resolve_entities=False, no_network=True))
                if words(part) or any(part.iter(W+"drawing")):
                    self.loss(f"Review {rel.target_part.partname} in Word and restore omitted text or graphics from this part.")
        if not result or not any(x.strip() for x in result):
            raise ValueError("The document has no renderable main-body content.")
        return "\n\n".join(result)

    def latex(self):
        values = {"LANG":self.language, "TITLE":text(self.title), "BODY":self.body()}
        template = files(__package__).joinpath("report.tex").read_text(encoding="utf-8")
        return re.sub(r"@@(LANG|TITLE|BODY)@@", lambda m: values[m[1]], template)
