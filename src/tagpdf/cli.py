"""Command-line conversion, local validation, and exclusive publication of outputs."""

import argparse
from dataclasses import dataclass
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
from tempfile import TemporaryDirectory
from zipfile import BadZipFile

from docx.opc.exceptions import PackageNotFoundError
from lxml import etree

from .word import Report


def find_tool(name):
    override = os.environ.get("TAGPDF_" + name.upper())
    if override:
        return shutil.which(override)
    fallback = (Path("/Library/TeX/texbin/lualatex") if name == "lualatex"
                else Path.home() / "verapdf" / ("verapdf.bat" if os.name == "nt" else "verapdf"))
    return shutil.which(name) or shutil.which(str(fallback))


def compile_pdf(work):
    engine = find_tool("lualatex")
    if not engine:
        raise ValueError("LuaLaTeX was not found. Set TAGPDF_LUALATEX to the TeX Live executable.")
    env = os.environ.copy()
    env.update(openin_any="p", openout_any="p")
    for _ in range(3):
        run = subprocess.run([engine, "-no-shell-escape", "-interaction=nonstopmode", "-halt-on-error",
                              "-file-line-error", "report.tex"], cwd=work, env=env, capture_output=True,
                             timeout=180, text=True, errors="replace")
        if run.returncode:
            raise ValueError("LuaLaTeX failed to compile the report.\n" + "\n".join(run.stdout.splitlines()[-22:]))
    log = (work / "report.log").read_text(encoding="utf-8", errors="replace")
    return log


@dataclass
class Validation:
    status: str
    details: list[str]


def validate(work):
    validator = find_tool("verapdf")
    if not validator:
        return Validation("UNAVAILABLE", ["Set TAGPDF_VERAPDF to the local veraPDF CLI and validate the PDF."])
    try:
        run = subprocess.run([validator, "--format", "xml", "--flavour", "ua2", "report.pdf"],
                             cwd=work, capture_output=True, timeout=180)
        xml = etree.fromstring(run.stdout, etree.XMLParser(resolve_entities=False, no_network=True))
        reports = xml.findall(".//validationReport")
        if len(reports) != 1:
            raise ValueError("Expected one validation report from veraPDF.")
        report = reports[0]
        profile = report.get("profileName", "")
        verdict = report.get("isCompliant")
        if "PDF/UA-2" not in profile or verdict not in ("true", "false") or (verdict == "true" and run.returncode):
            raise ValueError(f"Unexpected validation result ({profile}, exit {run.returncode}).")
        details = [profile]
        version = xml.find(".//releaseDetails[@id='core']")
        if version is not None:
            details.append("veraPDF " + version.get("version", "unknown version"))
        summary = report.find("details")
        if summary is not None:
            details.append(f"{summary.get('passedChecks')} checks passed, {summary.get('failedChecks')} checks failed.")
        for rule in report.findall(".//rule[@status='failed']"):
            details.append(f"{rule.get('specification')} §{rule.get('clause')}, test {rule.get('testNumber')}"
                           f" ({rule.get('failedChecks')} failed) — {rule.findtext('description','').strip()}")
        return Validation("PASS" if verdict == "true" else "FAIL", details)
    except (OSError, subprocess.TimeoutExpired, etree.XMLSyntaxError, ValueError) as exc:
        return Validation("ERROR", [f"veraPDF could not return a verified PDF/UA-2 result. {exc}"])


def review_report(source, document, pages, validation):
    findings = list(dict.fromkeys(document.findings))
    lost = [f for f in findings if f.lost]
    lines = [f"Accessibility review for {source.with_suffix('.pdf').name}", "",
             "Conversion: " + ("INCOMPLETE — see Content loss." if lost else "No content loss detected."),
             f"Automated validation: PDF/UA-2 — {validation.status}", *validation.details]
    for heading, entries in [("Content loss", lost), ("Needs human review", [f for f in findings if not f.lost])]:
        lines.extend(["", heading])
        if not entries:
            lines.append("None detected.")
        for entry in entries:
            place = f"Page {pages[entry.marker]}" if entry.marker in pages else "Output page unknown"
            lines.append(f"- {place}, {entry.label} — {entry.action}")
    lines.extend(["", "Before publication", "Open the PDF in Acrobat Pro. Compare the content and tag reading order "
                  "with Word, resolve the items above, and run Acrobat’s accessibility checker. "
                  "The employee performing the Section 508 review decides whether to publish.", ""])
    return "\n".join(lines)


def publish(work, outputs):
    # Hard links install whole files exclusively. On volumes without hard links,
    # exclusive creation provides the same no-overwrite guarantee.
    created = []
    try:
        for staged, target in zip((work/"report.pdf", work/"review.txt"), outputs, strict=True):
            try:
                os.link(staged, target)
            except FileExistsError:
                raise
            except OSError:
                with target.open("xb") as dst:
                    created.append((target, os.fstat(dst.fileno())))
                    with staged.open("rb") as src:
                        shutil.copyfileobj(src, dst)
            else:
                created.append((target, staged.stat()))
    except BaseException:
        for target, stat in created:
            if target.exists() and os.path.samestat(target.stat(), stat):
                target.unlink()
        raise


def convert(source):
    if source.suffix.lower() != ".docx" or not source.is_file():
        raise ValueError("Choose an existing .docx file.")
    outputs = (source.with_suffix(".pdf"), source.with_suffix(".review.txt"))
    for path in outputs:
        if os.path.lexists(path):
            raise FileExistsError(f"Output already exists: {path}. Move it or use a renamed Word copy.")
    with TemporaryDirectory(prefix=".tagpdf-", dir=source.parent) as temp:
        work = Path(temp)
        document = Report(source, work)
        (work / "report.tex").write_text(document.latex(), encoding="utf-8")
        log = compile_pdf(work)
        for glyph in sorted(set(re.findall(r"Missing character: There is no (.*?) in font", log))):
            document.loss(f"Restore the missing glyph {glyph}. The output font could not display it.")
        if "Overfull \\hbox" in log or "Overfull \\vbox" in log:
            document.review("Inspect page margins for clipped or overlapping content flagged by LuaLaTeX.")
        pages = dict((int(a), int(b)) for a,b in re.findall(r"^(\d+) (\d+)$", (work/"report.pages").read_text(), re.M))
        validation = validate(work)
        report = review_report(source, document, pages, validation)
        (work / "review.txt").write_text(report, encoding="utf-8")
        publish(work, outputs)
    for path in outputs:
        print(path)
    print(f"veraPDF PDF/UA-2: {validation.status}. See the review report.")
    return 2 if validation.status != "PASS" or any(f.lost for f in document.findings) else 0


def main(argv=None):
    parser = argparse.ArgumentParser(prog="tagpdf", description="Convert a Word report to a tagged PDF for accessibility review.")
    parser.add_argument("source", type=Path, metavar="REPORT.docx")
    try:
        return convert(parser.parse_args(argv).source.absolute())
    except (OSError, ValueError, KeyError, BadZipFile, PackageNotFoundError,
            etree.XMLSyntaxError, subprocess.TimeoutExpired) as exc:
        print(f"tagpdf: {exc}", file=sys.stderr)
        return 1
