"""Recover Word's logical table grid and emit tagged, page-breakable rows."""

from dataclasses import dataclass, field

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
CHART = "{http://schemas.openxmlformats.org/drawingml/2006/chart}chart"
FALSE = {"0", "false", "off"}
# Word supports up to 63 columns. Validate before expanding any untrusted span.
MAX_COLUMNS = 63


def _property(node, path, default=None):
    element = node.find("/".join(W + name for name in path.split("/")))
    return element.get(W + "val", "1") if element is not None else default


def _count(node, path, default):
    value = int(_property(node, path, default))
    if not 0 <= value <= MAX_COLUMNS:
        raise ValueError("invalid Word table grid width")
    return value


def _look(table, name, bit):
    look = table.find(W + "tblPr/" + W + "tblLook")
    if look is None:
        return False
    if W + name in look.attrib:
        return look.get(W + name).lower() not in FALSE
    return bool(int(look.get(W + "val", "0"), 16) & bit)


def _members(parent, tag):
    """Permit Word content controls and revisions around rows or cells."""
    for child in parent:
        if child.tag == tag:
            yield child
        elif child.tag in {W + n for n in ("sdt", "sdtContent", "ins", "moveTo", "customXml")}:
            yield from _members(child, tag)


@dataclass(eq=False)
class Cell:
    row: int
    col: int
    width: int
    nodes: list = field(default_factory=list)
    height: int = 1

    @property
    def key(self):
        return f"{self.row + 1}-{self.col + 1}"

    @property
    def name(self):
        return "".join(t.text or "" for n in self.nodes for t in n.iter(W + "t")).strip()


def _keep_continuation(origin, node, ctx):
    # Even runs with no w:t may hold a drawing, symbol or reference.
    if any(node.itertext()) or node.find(".//" + W + "r") is not None:
        origin.nodes.append(node)
        ctx.review(f"Check the additional content in merged cell {origin.key} against Word.")


def _grid(table, ctx):
    row_nodes = list(_members(table, W + "tr"))
    width = len(table.findall(W + "tblGrid/" + W + "gridCol"))
    grid = []
    previous_merges = {}
    for r, tr in enumerate(row_nodes):
        col = _count(tr, "trPr/gridBefore", 0)
        row = [None] * col
        horizontal = None
        vertical = {}
        for node in _members(tr, W + "tc"):
            span = _count(node, "tcPr/gridSpan", 1)
            if not span or col + span > MAX_COLUMNS:
                raise ValueError("unsupported grid span")
            vmerge = _property(node, "tcPr/vMerge")
            hmerge = _property(node, "tcPr/hMerge")
            # An empty merge element denotes a continuation. Track each vertical
            # run by its entire grid interval to avoid guessing across mismatches.
            if hmerge in ("1", "continue") and horizontal is not None:
                cell = horizontal
                cell.width += span
                _keep_continuation(cell, node, ctx)
            elif vmerge in ("1", "continue") and (col, span) in previous_merges:
                cell = previous_merges[col, span]
                cell.height += 1
                _keep_continuation(cell, node, ctx)
            else:
                cell = Cell(r, col, span, [node])
                if vmerge in ("1", "continue") or hmerge in ("1", "continue"):
                    ctx.review(f"Repair the unmatched merge at cell {cell.key} and confirm its headers.")
            horizontal = cell if hmerge is not None else None
            if vmerge is not None:
                vertical[col, span] = cell
            row.extend([cell] * span)
            col += span
        width = max(width, col + _count(tr, "trPr/gridAfter", 0))
        grid.append(row)
        previous_merges = vertical

    if not row_nodes or not 0 < width <= MAX_COLUMNS:
        raise ValueError("empty or oversized grid")
    for row in grid:
        row.extend([None] * (width - len(row)))
    return row_nodes, grid


def _headers(table, rows, ctx):
    count = 0
    for i, row in enumerate(rows):
        repeat = _property(row, "trPr/tblHeader", "0").lower() not in FALSE
        if repeat:
            if count == i:
                count += 1
            else:
                ctx.review(f"Inspect the header setting on row {i + 1}, outside the leading header rows.")
    if not count and _look(table, "firstRow", 0x20):
        count = 1
        ctx.review("Confirm that the styled first row contains column headers.")
    if not count:
        ctx.review("Identify column headers in Acrobat. Word did not identify a header row.")
    row_headers = _look(table, "firstColumn", 0x80)
    if row_headers:
        ctx.review("Confirm that the styled first column contains row headers.")
    return count, row_headers


def _applies(cell, grid, nhead, first_col):
    # Collect intersecting grid slots, then deduplicate merged headers in order.
    headers = {}
    for r in range(min(nhead, cell.row)):
        for header in grid[r][cell.col:cell.col + cell.width]:
            if header is not None and header is not cell:
                headers[header] = None
    if first_col and cell.col:
        for r in range(max(cell.row, nhead), cell.row + cell.height):
            header = grid[r][0]
            if header is not None and header.row >= nhead:
                headers[header] = None
    return list(headers)


def _paragraphs(node, ctx):
    paragraphs = []
    for p in node.iter(W + "p"):
        # Do not read textbox paragraphs a second time inside the enclosing run.
        if any(a.tag == W + "p" for a in p.iterancestors()):
            continue
        if any(a.tag in (W + "del", W + "moveFrom") for a in p.iterancestors()):
            continue
        properties, _ = ctx.properties(p)
        item = ctx.numbering.item(properties, ctx)
        content = ctx.inline_paragraph(p)
        if item is not None:
            ctx.loss("Restore list tags inside the table cell. List labels and paragraph text were preserved.")
            content = ctx.text(item[2]) + " " + content
        if any(prop is not None and _property(prop, "outlineLvl", "9") != "9" for prop in properties):
            ctx.loss("Restore the heading tag inside the table cell. Its text was preserved as a paragraph.")
        paragraphs.append(content)
    return paragraphs


def _linear(table, ctx, caption):
    """A grid we cannot represent still contributes all recoverable content."""
    return "\n".join([r"\par", ctx.anchor + caption + r"\par",
                      *(p + r"\par" for p in _paragraphs(table, ctx))])


def render_table(table, ctx):
    caption_bits = [ctx.caption]
    for name in ("tblCaption", "tblDescription"):
        raw = _property(table, "tblPr/" + name, "")
        value = ctx.text(raw)
        if value and value not in caption_bits:
            caption_bits.append(value)
    caption = r"\par ".join(filter(None, caption_bits))
    # A cell drawing has its own identity and must not inherit the table caption.
    old_caption = ctx.caption
    ctx.caption = ""
    try:
        return _render(table, ctx, caption)
    finally:
        ctx.caption = old_caption


def _render(table, ctx, caption):
    if table.find(".//" + CHART) is not None:
        ctx.loss("The enclosing table grid was omitted to preserve the embedded native chart and its data. Rebuild the enclosing table from Word.")
        return _linear(table, ctx, caption)
    if table.find(".//" + W + "tbl") is not None:
        ctx.loss("Nested table cells were flattened to paragraphs. Restore the inner table structure and header associations.")
    if table.find(".//" + W + "altChunk") is not None:
        ctx.loss("Restore the embedded subdocument omitted from this table.")
    if table.find(".//" + W + "hMerge") is not None and table.find(".//" + W + "vMerge") is not None:
        ctx.loss("Combined legacy horizontal and vertical merges were flattened. Rebuild their grid around the recovered cell contents.")
        return _linear(table, ctx, caption)
    try:
        row_nodes, grid = _grid(table, ctx)
        nhead, first_col = _headers(table, row_nodes, ctx)
    except ValueError:
        ctx.loss("The Word table grid was invalid or too wide to encode. Rebuild the table around the recovered cell contents.")
        return _linear(table, ctx, caption)

    width = len(grid[0])
    if any(None in row for row in grid):
        ctx.review("Word omits grid cells in one or more rows. Check the empty positions and their header associations.")

    header_rows = ",".join(str(i + 1) for i in range(nhead))
    col_ids = "1" if first_col else ""
    col_type = r">{\raggedright\arraybackslash}p{\dimexpr\linewidth/" + str(width)
    tail = r"-2\tabcolsep\relax}"
    lines = [r"\par\begingroup",
             r"\setlength{\tabcolsep}{\dimexpr\linewidth/" + str(max(90, width * 12)) + r"\relax}",
             rf"\tagpdfsetup{{table/header-rows={{{header_rows}}},table/header-columns={{{col_ids}}}}}"]
    if caption:
        # longtable opens the Table structure just before its array starts.
        # The kernel's normal longtable caption would occupy a logical grid row.
        lines.append(r"\AddToHookNext{cmd/LT@array/before}{\tagstructbegin{tag=Caption}\noindent "
                     + ctx.anchor + caption + r"\par\tagstructend\nopagebreak}")
    lines += [r"\begin{longtable}{*{" + str(width) + "}{" + col_type + tail + "}}", r"\hline"]
    anchor = "" if caption else ctx.anchor
    for r, row in enumerate(grid):
        rendered = []
        c = 0
        while c < width:
            cell = row[c]
            span = cell.width if cell is not None else 1
            value = ""
            if cell is not None and cell.row == r:
                associated = _applies(cell, grid, nhead, first_col)
                if cell.height > 1:
                    value += rf"\tagpdfsetup{{table/multirow={cell.height}}}"
                if r < nhead or (first_col and c == 0):
                    value += rf"\TagSaveHeader{{{cell.key}}}"
                if associated:
                    refs = " ".join(rf"\TagHeaderID{{{h.key}}}" for h in associated)
                    value += r"\TagApplyHeaders{" + refs + "}"
                if r < nhead and cell.width > 1:
                    lower = []
                    if r + 1 < nhead:
                        lower = list(dict.fromkeys(grid[r + 1][c:c + span]))
                    names = [f"“{h.name}”" for h in lower if h is not None and h is not cell and h.name]
                    columns = ", ".join(names) if names else f"columns {c + 1} through {c + span}"
                    ctx.review(f"Confirm that header “{cell.name or cell.key}” applies to {columns}.")
                if cell.width > 1 and any(h.width < cell.width for h in associated if h.row < nhead):
                    ctx.review(f"Confirm the column headers for merged cell {cell.key}.")
                value += anchor
                anchor = ""
                if r < nhead or (first_col and c == 0):
                    value += r"\bfseries{}"
                value += r"\par ".join(p for n in cell.nodes for p in _paragraphs(n, ctx))
            if span > 1:
                value = rf"\multicolumn{{{span}}}{{" + col_type + "*" + str(span) + tail + "}{" + value + "}"
            rendered.append(value)
            c += span
        lines.append(" & ".join(rendered) + r"\\")
        if r + 1 == nhead:
            lines.append(r"\hline")
    lines.extend((r"\hline", r"\end{longtable}\TagFinishTable\endgroup"))
    return "\n".join(lines)
