"""Verify generated instructor pages and downloads against the L1715 manuscripts."""

from html.parser import HTMLParser
from pathlib import Path
import sys
from collections import Counter
from zipfile import ZipFile

from lxml import etree

from build_instructor_version import (
    MANUSCRIPTS,
    BEGIN_RE,
    HEADING_TAGS,
    NS,
    PAGES,
    PRODUCTION_PREFIX,
    ROOT,
    SIM_RE,
    clean_download_nodes,
    body_children,
    download_sections,
    main_content,
    manuscript_sources,
    slug,
    text_of,
)


class Blocks(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.depth = 0
        self.current = []
        self.blocks = []

    def handle_starttag(self, tag, attrs):
        if self.depth:
            self.depth += 1
        elif "data-manuscript-block" in {name for name, _value in attrs}:
            self.depth = 1
            self.current = []

    def handle_data(self, data):
        if self.depth:
            self.current.append(data)

    def handle_endtag(self, tag):
        if self.depth:
            self.depth -= 1
            if not self.depth:
                self.blocks.append("".join(self.current))


def html_blocks(value):
    parser = Blocks()
    parser.feed(value)
    return parser.blocks


def document_blocks(path):
    with ZipFile(path) as archive:
        root = etree.fromstring(archive.read("word/document.xml"))
    body = root.find("w:body", NS)
    result = []
    for child in body:
        paragraphs = [child] if child.tag.endswith("}p") else child.xpath(".//w:p", namespaces=NS)
        for paragraph in paragraphs:
            value = text_of(paragraph)
            if value:
                result.append(value)
    return result


def node_blocks(nodes):
    result = []
    for child in clean_download_nodes(nodes):
        paragraphs = [child] if child.tag.endswith("}p") else child.xpath(".//w:p", namespaces=NS)
        for paragraph in paragraphs:
            value = text_of(paragraph)
            if value and not value.startswith("\\qqINSERT"):
                result.append(value)
    return result


def compare(label, expected, actual):
    if expected == actual:
        print(f"PASS: {label}")
        return True
    print(f"FAIL: {label}")
    for index in range(max(len(expected), len(actual))):
        source = expected[index] if index < len(expected) else "<missing>"
        output = actual[index] if index < len(actual) else "<missing>"
        if source != output:
            print(f"  Block {index + 1}\n    SOURCE: {source!r}\n    OUTPUT: {output!r}")
            break
    return False


def validate_page(source, page):
    expected = html_blocks("".join(main_content(source)))
    if source.name == "L1715_Debriefing Methods.docx":
        expected = [block for block in expected if block != "Debriefing Methods"]
    actual = html_blocks(page.read_text(encoding="utf-8"))
    content_matches = compare(
        f"{page.relative_to(ROOT)} matches {source.name}", expected, actual
    )
    # This audit deliberately bypasses main_content so a parser defect cannot
    # silently remove a source heading from both expected and actual content.
    _root, body = body_children(source)
    source_headings = []
    in_download = False
    for node in list(body):
        value = text_of(node).strip()
        if BEGIN_RE.search(value):
            in_download = True
            continue
        if in_download:
            if value.startswith("\\qqEND downloadable content"):
                in_download = False
            continue
        match = PRODUCTION_PREFIX.match(value)
        if match and match.group(1) in HEADING_TAGS:
            source_headings.append(PRODUCTION_PREFIX.sub("", value))
    actual_counts = Counter(actual)
    missing = []
    for heading in source_headings:
        if actual_counts[heading]:
            actual_counts[heading] -= 1
        else:
            missing.append(heading)
    if missing:
        print(f"FAIL: {page.relative_to(ROOT)} is missing source headings: {missing}")
        headings_match = False
    else:
        print(f"PASS: {page.relative_to(ROOT)} includes every source heading")
        headings_match = True
    return content_matches and headings_match


def validate_downloads(source, number):
    results = []
    used = set()
    output_dir = ROOT / "assets" / "downloads" / f"simulation-{number}"
    for section in download_sections(source):
        if section["button"] == "Information for Proctor":
            continue
        meaningful = [
            text_of(node).strip()
            for node in clean_download_nodes(section["nodes"])
            if text_of(node).strip()
        ]
        if len(meaningful) <= 1:
            continue
        expected = node_blocks(section["nodes"])
        base = slug(section["button"])
        filename = f"{base}.docx"
        suffix = 2
        while filename in used:
            filename = f"{base}-{suffix}.docx"
            suffix += 1
        used.add(filename)
        output = output_dir / filename
        if not output.exists():
            print(f"FAIL: missing {output.relative_to(ROOT)}")
            results.append(False)
            continue
        results.append(compare(str(output.relative_to(ROOT)), expected, document_blocks(output)))
    return results


def main():
    requested = {int(arg) for arg in sys.argv[1:]}
    results = [
        validate_page(MANUSCRIPTS / "L1715_Introduction.docx", PAGES / "introduction.html"),
        validate_page(MANUSCRIPTS / "L1715_Debriefing Methods.docx", PAGES / "debriefing-methods.html"),
    ]
    for source in manuscript_sources():
        match = SIM_RE.match(source.name)
        if not match:
            continue
        number = int(match.group(1))
        if requested and number not in requested:
            continue
        results.append(validate_page(source, PAGES / f"simulation-{number}.html"))
        results.extend(validate_downloads(source, number))
    return 0 if all(results) else 1


if __name__ == "__main__":
    raise SystemExit(main())
