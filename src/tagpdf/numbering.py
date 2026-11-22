"""Resolve Word list instances, including numbering inherited from paragraph styles."""

from collections import defaultdict
from docx.oxml.ns import qn


def val(node, path, default=None):
    elem = node.find(path) if node is not None else None
    return elem.get(qn("w:val"), default) if elem is not None else default


def ordinal(value, fmt):
    if fmt in ("lowerLetter", "upperLetter") and 0 < value <= 18278:
        letters = ""
        while value:
            value, digit = divmod(value - 1, 26)
            letters = chr(65 + digit) + letters
        return letters.lower() if fmt == "lowerLetter" else letters
    if fmt in ("lowerRoman", "upperRoman") and 0 < value < 4000:
        result = ""
        for amount, letter in ((1000,"M"),(900,"CM"),(500,"D"),(400,"CD"),(100,"C"),
                               (90,"XC"),(50,"L"),(40,"XL"),(10,"X"),(9,"IX"),(5,"V"),(4,"IV"),(1,"I")):
            n, value = divmod(value, amount)
            result += letter * n
        return result.lower() if fmt == "lowerRoman" else result
    return str(value)


class Numbering:
    def __init__(self, doc):
        self.counts = defaultdict(dict)
        self.levels = {}
        rel = next((r for r in doc.part.rels.values() if r.reltype.endswith("/numbering")), None)
        if not rel:
            return
        root = rel.target_part.element
        definitions = {a.get(qn("w:abstractNumId")): a for a in root.findall(qn("w:abstractNum"))}
        for n in root.findall(qn("w:num")):
            num = n.get(qn("w:numId"))
            definition = definitions.get(val(n, qn("w:abstractNumId")))
            if definition is None:
                continue
            for level in definition.findall(qn("w:lvl")):
                self.levels[num, int(level.get(qn("w:ilvl")))] = level, None
            for override in n.findall(qn("w:lvlOverride")):
                key = num, int(override.get(qn("w:ilvl")))
                node = override.find(qn("w:lvl"))
                if node is None:
                    node = self.levels.get(key, (None, None))[0]
                self.levels[key] = node, val(override, qn("w:startOverride"))

    def item(self, properties, ctx):
        num = next((x for p in properties if (x := val(p, qn("w:numPr")+"/"+qn("w:numId"))) is not None), None)
        if num in (None, "0"):
            return None
        depth = int(next((x for p in properties if (x := val(p, qn("w:numPr")+"/"+qn("w:ilvl"))) is not None), "0"))
        node, start = self.levels.get((num, depth), (None, None))
        counts = self.counts[num]
        counts[depth] = counts.get(depth, int(start or val(node, qn("w:start"), "1")) - 1) + 1
        for d in list(counts):
            if d > depth:
                del counts[d]
        fmt = val(node, qn("w:numFmt"), "decimal")
        if node is None or fmt not in {"bullet", "decimal", "lowerLetter", "upperLetter", "lowerRoman", "upperRoman", "none"}:
            ctx.review("Compare this list's labels and numbering with Word.")
        label = val(node, qn("w:lvlText"), "%1.")
        for i in range(depth + 1):
            parent, _ = self.levels.get((num, i), (None, None))
            label = label.replace("%" + str(i+1), ordinal(counts.get(i, 1), val(parent, qn("w:numFmt"))))
        return num, depth, ("•" if fmt == "bullet" else "" if fmt == "none" else label)
