"""Account for active Word footers in the repaginated report."""

import re

from docx.oxml.ns import qn
from lxml import etree


PAGE_FIELD = re.compile(r"PAGE(?:\s+\\\*\s+MERGEFORMAT)?", re.IGNORECASE)
PAGE_LABEL = re.compile(r"(?:^|\t)\s*Page\s*(?:\|\s*)?$", re.IGNORECASE)
TEXT = qn("w:t")
TAB = qn("w:tab")
BREAK = qn("w:br")


def active_footers(document):
    """Resolve inherited first, even and default references in section order."""
    refs = {}
    active = {}
    blank_cover = False
    for i, section in enumerate(document.sections):
        for ref in section._sectPr.findall(qn("w:footerReference")):
            refs[ref.get(qn("w:type"))] = ref.get(qn("r:id"))
        kinds = ["default"]
        if section.different_first_page_header_footer:
            kinds.append("first")
            if i == 0:
                blank_cover = "first" not in refs
        if document.settings.odd_and_even_pages_header_footer:
            kinds.append("even")
        for kind in kinds:
            if kind in refs:
                active[refs[kind]] = None
    return list(active), blank_cover


def _run_content(node):
    for child in node:
        if child.tag == qn("w:r"):
            yield from (item for item in child if item.tag != qn("w:rPr"))
        elif child.tag != qn("w:pPr"):
            yield child


def _check_page_field(instruction):
    if not PAGE_FIELD.fullmatch(instruction.strip()):
        raise ValueError(f"unsupported field “{instruction.strip() or 'empty instruction'}”")


def footer_text(paragraph):
    """Return static text, excluding PAGE fields and their cached results."""
    chunks = [""]
    instruction = None
    reading_result = False
    for node in _run_content(paragraph):
        if node.tag == qn("w:fldChar"):
            kind = node.get(qn("w:fldCharType"))
            if kind == "begin" and instruction is None:
                instruction = ""
            elif kind == "separate" and instruction is not None and not reading_result:
                reading_result = True
            elif kind == "end" and instruction is not None:
                _check_page_field(instruction)
                chunks.append("")
                instruction, reading_result = None, False
            else:
                raise ValueError("nested or malformed field")
        elif node.tag == qn("w:instrText") and instruction is not None and not reading_result:
            instruction += node.text or ""
        elif node.tag == qn("w:fldSimple") and instruction is None:
            _check_page_field(node.get(qn("w:instr"), ""))
            if any(item.tag != TEXT for item in _run_content(node)):
                raise ValueError("unsupported content in a PAGE field result")
            chunks.append("")
        elif node.tag in {TEXT, TAB, BREAK} and (instruction is None or reading_result):
            if instruction is None:
                chunks[-1] += (node.text or "") if node.tag == TEXT else "\t" if node.tag == TAB else " "
        else:
            raise ValueError(f"unsupported element {etree.QName(node).localname}")
    if instruction is not None:
        raise ValueError("unterminated field")
    # The template supplies the page number. A conventional label immediately
    # before that field carries no additional information.
    static = [PAGE_LABEL.sub("", chunk) for chunk in chunks[:-1]] + chunks[-1:]
    return " ".join(" ".join(static).split())


def read_footers(document, body_text, loss):
    active, blank_cover = active_footers(document)
    unique = {}
    for rid in active:
        rel = document.part.rels.get(rid)
        if rel is None or rel.is_external or not rel.reltype.endswith("/footer"):
            loss(f"Restore the Word footer referenced by {rid}. Its part could not be read.")
            continue
        part = rel.target_part
        root = etree.fromstring(part.blob, etree.XMLParser(resolve_entities=False, no_network=True))
        for index, paragraph in enumerate(root, 1):
            try:
                if paragraph.tag != qn("w:p"):
                    raise ValueError(f"unsupported footer block {etree.QName(paragraph).localname}")
                value = footer_text(paragraph)
            except ValueError as exc:
                loss(f"Review {part.partname}, block {index}, in Word and restore omitted footer content ({exc}).")
                continue
            if value and value not in body_text:
                unique[value] = None
    return unique, blank_cover
