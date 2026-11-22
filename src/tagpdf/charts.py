"""Convert the literal and cached data in classic DrawingML chart parts."""

from bisect import bisect_right
from dataclasses import dataclass
from math import isfinite

C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def _text(element):
    if element is None:
        return ""
    paragraphs = element.findall(f".//{A}p")
    if paragraphs:
        return " ".join("".join(t.text or "" for t in p.iter(f"{A}t")) for p in paragraphs)
    return " ".join(t.text or "" for t in element.iter(f"{C}v"))


def _points(cache, ctx, field):
    result = {}
    for point in cache.findall(f"{C}pt"):
        try:
            index = int(point.get("idx", ""))
            if not 0 <= index < 2**32 or index in result:
                raise ValueError()
        except ValueError:
            ctx.loss(f"A {field} point has an invalid or repeated index. Recover it from the Word chart.")
            continue
        value = point.find(f"{C}v")
        if value is None:
            ctx.loss(f"The stored value for {field} point {index + 1} is missing. Recover it from Word.")
        result[index] = (value.text or "") if value is not None else ""
    return result


def _dimension(source, ctx, field, indices=()):
    """Resolve chart caches by point index. Sparse caches may represent blank cells."""
    if source is None:
        return {}
    for cache in source.iter():
        if cache.tag in {f"{C}{kind}" for kind in ("numLit", "numCache", "strLit", "strCache")}:
            code = cache.findtext(f"{C}formatCode", default="General")
            if code not in ("General", "0"):
                ctx.review(f"Confirm {field} units and formatting against the source format “{code}”. "
                           "The data table displays stored values.")
            points = _points(cache, ctx, field)
            count = cache.find(f"{C}ptCount")
            if count is not None and count.get("val", "0") != "0" and not points:
                ctx.loss(f"The {field} cache declares points but contains none. Recover them from Word.")
            return points
        if cache.tag == f"{C}multiLvlStrCache":
            # The lowest level comes first. A higher-level entry spans points up
            # to the next entry in its level (e.g. a year over four quarters).
            levels = [_points(level, ctx, field) for level in cache.findall(f"{C}lvl")]
            if not levels:
                break
            starts = [sorted(level) for level in levels]
            result = {}
            for index in sorted(set(indices).union(*levels)):
                names = []
                for level, keys in zip(reversed(levels[1:]), reversed(starts[1:])):
                    pos = bisect_right(keys, index)
                    if pos:
                        names.append(level[keys[pos - 1]])
                names.append(levels[0].get(index, ""))
                result[index] = " / ".join(filter(None, names))
            return result
    ctx.loss(f"The chart part has no usable cached {field} data. Recover {field} from the source workbook.")
    return {}


@dataclass
class Series:
    name: str
    categories: dict
    values: dict
    sizes: dict
    labels: dict
    plot: str


def _number(value):
    try:
        num = float(value)
    except (ValueError, TypeError):
        return None
    return num if isfinite(num) else None


def _plot(series, ymin, ymax, ctx):
    """Draw a monochrome vector plot. Point numbers key the plot to the table."""
    values = [(index, _number(raw)) for index, raw in sorted(series.values.items())]
    points = [(i, v) for i, v in values if v is not None]
    if not points:
        return ""
    # Normalize before subtracting to avoid overflow for opposite-sign extremes.
    scale = max(abs(ymin), abs(ymax), 1)
    low, high = ymin / scale, ymax / scale
    def y(value):
        return 14 + 42 * ((value / scale - low) / (high - low))

    last = max(series.categories.keys() | series.values.keys())
    step = 210 / (last + 1)
    baseline = y(0)
    lines = [r"\begin{picture}[off](250,64)",
             rf"\put(36,{baseline:.3f}){{\line(1,0){{210}}}}",
             r"\put(36,14){\line(0,1){42}}",
             rf"\put(33,14){{\makebox(0,0)[r]{{\scriptsize {ymin:.3g}}}}}",
             rf"\put(33,56){{\makebox(0,0)[r]{{\scriptsize {ymax:.3g}}}}}"]
    previous = None
    for i, v in points:
        x = 36 + (i + 0.5) * step
        height = y(v)
        if series.plot == "barChart":
            width = min(18, step * 0.55)
            lines.append(rf"\put({x - width / 2:.3f},{min(height,baseline):.3f})"
                         rf"{{\rule{{{width:.3f}pt}}{{{abs(height-baseline):.3f}pt}}}}")
        else:
            if previous and previous[0] == i - 1:
                _, px, py = previous
                lines.append(rf"\qbezier({px:.3f},{py:.3f})"
                             rf"({(x+px)/2:.3f},{(height+py)/2:.3f})({x:.3f},{height:.3f})")
            lines.append(rf"\put({x:.3f},{height:.3f}){{\circle*{{2}}}}")
            previous = i, x, height
        if last < 12 or i in (points[0][0], points[-1][0]):
            category = series.categories.get(i, "")
            tick = category if category and len(category) * 7 <= step else str(i + 1)
            lines.append(rf"\put({x:.3f},5){{\makebox(0,0){{\scriptsize {ctx.text(tick)}}}}}")
    return "\n".join([*lines, r"\end{picture}"])


def _point_labels(owner, ctx):
    labels = {}
    for label in owner.findall(f"{C}dLbls/{C}dLbl"):
        custom = label.find(f"{C}tx")
        if custom is None:
            continue
        index = label.find(f"{C}idx")
        value = _text(custom)
        try:
            i = int(index.get("val", "") if index is not None else "")
            if not 0 <= i < 2**32 or not value:
                raise ValueError()
        except ValueError:
            ctx.loss("A custom chart point label was unreadable. Recover its text and placement.")
            continue
        labels[i] = value
    return labels


def _series(group, node, number, ctx):
    tx = node.find(f"{C}tx")
    name = _text(tx)
    if tx is not None and not name:
        ctx.loss(f"Series {number} has an unresolved name. Restore the series label from the workbook.")
    name = name or f"Series {number}"
    def data(tag, label, indices=()):
        return _dimension(node.find(f"{C}{tag}"), ctx, f"{name} {label}", indices)

    xy = node.find(f"{C}yVal") is not None or node.find(f"{C}xVal") is not None
    values = data("yVal" if xy else "val", "values")
    categories = data("xVal" if xy else "cat", "categories", values)
    sizes = data("bubbleSize", "bubble sizes")
    labels = _point_labels(group, ctx) | _point_labels(node, ctx)
    if not values:
        ctx.loss(f"Series “{name}” has no readable values. Restore the series from the source chart.")
    if values and any(_number(value) is None for value in values.values()):
        ctx.review(f"Check blank or nonnumeric points in “{name}” against the source chart.")
    if categories and values and categories.keys() != values.keys():
        ctx.review(f"Check the blank category or value cells in “{name}” against the source chart.")
    return Series(name, categories, values, sizes, labels, group.tag.removeprefix(C))


def _data_table(s, ctx):
    widths = {
        3: [".12", ".40", ".31"],
        4: [".11", ".25", ".22", ".23"],
        5: [".09", ".20", ".17", ".17", ".17"],
    }[3 + bool(s.sizes) + bool(s.labels)]
    columns = "".join(f"p{{{w}\\linewidth}}" for w in widths)
    header = ["Point", "X" if s.plot in ("scatterChart", "bubbleChart") else "Category", s.name]
    if s.sizes:
        header.append("Bubble size")
    if s.labels:
        header.append("Label")
    rows = [r"\begingroup\tagpdfsetup{table/header-rows={1},table/header-columns={}}",
            r"\begin{longtable}[l]{@{}" + columns + "@{}}",
            " & ".join(r"\textbf{" + ctx.text(h) + "}" for h in header) + r" \\",
            r"\hline\endhead"]
    for i in sorted(s.categories.keys() | s.values.keys() | s.sizes.keys() | s.labels.keys()):
        cells = [str(i + 1), s.categories.get(i, ""), s.values.get(i, "")]
        if s.sizes:
            cells.append(s.sizes.get(i, ""))
        if s.labels:
            cells.append(s.labels.get(i, ""))
        rows.append(" & ".join(ctx.text(value) for value in cells) + r" \\")
    return "\n".join([*rows, r"\end{longtable}\endgroup"])


def render_chart(chart_root, ctx, description=""):
    """Render chart graphics and accessible data tables in document reading order."""
    chart = chart_root.find(f"{C}chart")
    area = chart.find(f"{C}plotArea") if chart is not None else None
    title = _text(chart.find(f"{C}title")) if chart is not None else ""
    if area is None:
        ctx.loss("The native chart has no supported plot area. Restore the chart and its data.")
        return ctx.anchor + ctx.caption + r"\par"
    groups = [child for child in area if isinstance(child.tag, str) and child.tag.startswith(C)
              and child.tag.endswith("Chart")]
    series = []
    for group in groups:
        for s in group.findall(f"{C}ser"):
            series.append(_series(group, s, len(series) + 1, ctx))
        grouping = group.find(f"{C}grouping")
        if grouping is not None and grouping.get("val") in ("stacked", "percentStacked"):
            ctx.loss("The stacked plot relationship was omitted. Rebuild the stacked plot from the series tables.")
        if group.tag not in (f"{C}barChart", f"{C}lineChart"):
            kind = group.tag.removeprefix(C)
            ctx.loss(f"The {kind} plot was omitted. Rebuild its graphic using the recovered series tables.")
    if not series:
        ctx.loss("No native chart series could be read. Restore the chart and its data from Word.")
    for chart_title in chart.iter(f"{C}title"):
        if chart_title.find(f"{C}tx") is not None and not _text(chart_title):
            ctx.loss("A chart or axis title has no readable text. Restore the title from Word.")
    all_series = list(chart_root.iter(f"{C}ser"))
    if len(all_series) > len(series):
        ctx.loss("Some native chart series occur outside a supported plot group. Recover those series from Word.")
    for tag, feature in [("trendline", "trend lines"), ("errBars", "error bars"),
                         ("logBase", "logarithmic scaling"), ("dispUnits", "axis display units"),
                         ("userShapes", "chart annotations"), ("extLst", "chart extension content")]:
        if chart_root.find(f".//{C}{tag}") is not None:
            ctx.loss(f"Conversion omitted {feature}. Check the source chart and restore its meaningful content.")
    if not description:
        ctx.review("Read the chart and its data table. Add a description stating the comparison or trend.")
    alt = description or f"{title or ctx.label}. Values and category labels follow in the data table."
    # A single Figure owns the source caption and graphic. Keep the data tables
    # as siblings so a reader that substitutes /Alt can still navigate cells.
    out = [r"\par\begingroup\setlength{\unitlength}{1pt}", r"\tagstructbegin{tag=Div}",
           r"\tagstructbegin{tag=Figure,alt={" + ctx.text(alt) + "}}"]
    heading = ctx.caption or ctx.text(title or ctx.label)
    out.extend([r"\tagstructbegin{tag=Caption}", r"\noindent " + ctx.anchor + heading + r"\par",
                r"\tagstructend"])
    if ctx.caption and title:
        out.append(r"\noindent\textbf{" + ctx.text(title) + r"}\par")
    axis_titles = [_text(t) for axis in area for t in axis.findall(f"{C}title")]
    if any(axis_titles):
        out.append(r"\noindent " + ctx.text(" / ".join(filter(None, axis_titles))) + r"\par")
    numbers = [v for s in series for raw in s.values.values() if (v := _number(raw)) is not None]
    ymin = min(0, min(numbers, default=0))
    ymax = max(0, max(numbers, default=0))
    if ymin == ymax:
        ymax = 1
    for s in series:
        if s.plot in ("barChart", "lineChart"):
            plot = _plot(s, ymin, ymax, ctx)
            if plot:
                out.extend([r"\noindent " + ctx.text(s.name) + r"\par\nobreak",
                            r"\tagpdfparaOff\noindent\tagmcbegin{tag=Figure}" + plot
                            + r"\tagmcend\par\tagpdfparaOn"])
            elif s.values:
                ctx.loss(f"No plotted values could be drawn for “{s.name}”. Rebuild its graphic from Word.")
    out.append(r"\tagstructend")
    for s in series:
        out.append(_data_table(s, ctx))
    return "\n".join([*out, r"\tagstructend\endgroup"])
