"""Verify generated instructor pages and downloads against the L1715 manuscripts."""

from pathlib import Path
import re
import sys
from collections import Counter
from zipfile import ZipFile

from lxml import etree, html

from build_instructor_version import (
    MANUSCRIPTS,
    BEGIN_RE,
    HEADING_TAGS,
    NS,
    NOTE_END_RE,
    PAGES,
    PRODUCTION_PREFIX,
    ROOT,
    SIM_RE,
    clean_download_nodes,
    body_blocks,
    body_children,
    document_formatting,
    download_sections,
    effective_run_properties,
    manuscript_sources,
    prepare_reader_node,
    qn,
    slug,
    text_of,
)


def html_blocks(value):
    document = html.fromstring(value)
    return [
        "".join(element.itertext())
        for element in document.xpath("//*[@data-manuscript-block]")
    ]


def append_segment(segments, text, flags):
    if not text:
        return
    key = tuple(sorted(flags))
    if segments and segments[-1][1] == key:
        segments[-1] = (segments[-1][0] + text, key)
    else:
        segments.append((text, key))


def html_format_blocks(value):
    document = html.fromstring(value)
    results = []

    def child_flags(element, inherited):
        flags = set(inherited)
        if element.tag in {"strong", "b"}:
            flags.add("bold")
        if element.tag in {"em", "i"}:
            flags.add("italic")
        if element.tag == "u":
            flags.add("underline")
        if element.tag in {"s", "strike", "del"}:
            flags.add("strike")
        if element.tag == "sup":
            flags.add("superscript")
        if element.tag == "sub":
            flags.add("subscript")
        if element.tag == "a" and element.get("href"):
            flags.add("link")
        classes = set((element.get("class") or "").split())
        if "manuscript-caps" in classes:
            flags.add("caps")
        if "manuscript-small-caps" in classes:
            flags.add("smallCaps")
        return flags

    def collect(element, flags, segments):
        if element.text:
            append_segment(segments, element.text, flags)
        for child in element:
            collect(child, child_flags(child, flags), segments)
            if child.tail:
                append_segment(segments, child.tail, flags)

    for element in document.xpath("//*[@data-manuscript-block]"):
        segments = []
        collect(element, set(), segments)
        results.append(segments)
    return results


def source_format_segments(node, styles):
    segments = []

    def collect(parent, inherited):
        for child in parent:
            if child.tag == qn("w:r"):
                properties = effective_run_properties(child, styles)
                flags = set(inherited)
                for key, label in (
                    ("b", "bold"),
                    ("i", "italic"),
                    ("u", "underline"),
                    ("strike", "strike"),
                    ("caps", "caps"),
                    ("smallCaps", "smallCaps"),
                ):
                    if properties.get(key):
                        flags.add(label)
                vertical = properties.get("vertAlign")
                if vertical in {"superscript", "subscript"}:
                    flags.add(vertical)
                append_segment(
                    segments,
                    "".join(child.xpath(".//w:t/text()", namespaces=NS)),
                    flags,
                )
            elif child.tag == qn("w:hyperlink"):
                collect(child, set(inherited) | {"link"})
            elif child.tag != qn("w:del"):
                collect(child, inherited)

    collect(node, set())
    return segments


def source_format_blocks(path):
    _root, body = body_children(path)
    _relationships, styles = document_formatting(path)
    results = []
    in_download = False
    in_note = False
    for node in body_blocks(body):
        value = text_of(node).strip()
        value = re.sub(r"^(?:PHOTO HERE\s*)+", "", value).strip()
        if BEGIN_RE.search(value):
            in_download = True
            in_note = False
            continue
        if in_download:
            if value.startswith("\\qqEND downloadable content"):
                in_download = False
            continue
        if value.startswith("\\qqID:") or value.startswith("\\qqPSM:"):
            in_note = not NOTE_END_RE.search(value)
            continue
        if in_note:
            if NOTE_END_RE.search(value):
                in_note = False
                continue
            if not PRODUCTION_PREFIX.match(value):
                continue
            in_note = False
        if not value or value.startswith("Navigation menu/button") or value.startswith("\\qq"):
            continue
        if node.tag == qn("w:tbl"):
            for row in node.findall("w:tr", NS):
                for cell in row.findall("w:tc", NS):
                    cell_segments = []
                    for paragraph in cell.findall("w:p", NS):
                        clean = prepare_reader_node(paragraph)
                        for text, flags in source_format_segments(clean, styles):
                            append_segment(cell_segments, text, flags)
                    results.append(cell_segments)
            continue
        clean = prepare_reader_node(node)
        clean_text = text_of(clean)
        marker = PRODUCTION_PREFIX.match(value)
        if marker and marker.group(1) in {"cn", "ct"}:
            continue
        if path.name == "L1715_Debriefing Methods.docx" and clean_text == "Debriefing Methods":
            continue
        results.append(source_format_segments(clean, styles))
    return results


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


def compare_formatting(label, expected, actual):
    if expected == actual:
        counts = Counter(
            flag
            for block in expected
            for _text, flags in block
            for flag in flags
        )
        summary = ", ".join(f"{name}={count}" for name, count in sorted(counts.items()))
        print(f"PASS: {label} formatting" + (f" ({summary})" if summary else ""))
        return True
    print(f"FAIL: {label} formatting")
    for index in range(max(len(expected), len(actual))):
        source = expected[index] if index < len(expected) else [("<missing>", ())]
        output = actual[index] if index < len(actual) else [("<missing>", ())]
        if source != output:
            print(f"  Block {index + 1}\n    SOURCE: {source!r}\n    OUTPUT: {output!r}")
            break
    return False


def validate_page(source, page):
    source_formatted = source_format_blocks(source)
    expected = ["".join(text for text, _flags in block) for block in source_formatted]
    page_source = page.read_text(encoding="utf-8")
    actual = html_blocks(page_source)
    content_matches = compare(
        f"{page.relative_to(ROOT)} matches {source.name}", expected, actual
    )
    formatting_matches = compare_formatting(
        str(page.relative_to(ROOT)),
        source_formatted,
        html_format_blocks(page_source),
    )
    # This audit deliberately bypasses main_content so a parser defect cannot
    # silently remove a source heading from both expected and actual content.
    _root, body = body_children(source)
    source_headings = []
    in_download = False
    for node in body_blocks(body):
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
    return content_matches and formatting_matches and headings_match


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
