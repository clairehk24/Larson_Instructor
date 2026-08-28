"""Build instructor HTML pages and downloads from the available L1715 manuscripts."""

from __future__ import annotations

from copy import deepcopy
from html import escape, unescape
from pathlib import Path, PureWindowsPath
import re
from shutil import copy2
import sys
from zipfile import ZIP_DEFLATED, ZipFile

from docx import Document
from docx.oxml.ns import qn
from lxml import etree

from generate_simulation_downloads import finalize_document, stamp_footer, write_docx


ROOT = Path(__file__).resolve().parents[1]
MANUSCRIPTS = ROOT / "assets" / "Manuscripts"
PAGES = ROOT / "pages"
DOWNLOADS = ROOT / "assets" / "downloads"
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
PRODUCTION_PREFIX = re.compile(r"^<(cn|ct|a|b|c|d|lh|tt|title|txni|tx|fc)>")
HEADING_TAGS = {"a": "h2", "b": "h3", "c": "h4", "d": "h5", "lh": "h3", "tt": "h3", "title": "h2"}
LIST_SECTIONS = {
    "Learning Objectives": ("ol", "learning-objectives-list"),
    "CAATE 2020 Standards": ("ul", "standards-list"),
    "References": ("ol", "references-list"),
}
SOURCE_LIST_NUMBER_RE = re.compile(r"^(\d+\.\s*)")
SIM_RE = re.compile(r"L1715_Sim(\d{2})")
BEGIN_RE = re.compile(r"BEGIN downloadable content")
BUTTON_RE = re.compile(r"Button name:\s*(.*?)(?=<title>|\\?$)")
NOTE_END_RE = re.compile(r"xqq\\$", re.IGNORECASE)
TITLE_OVERRIDES = {
    5: "Special Test Roulette: Upper Extremity",
    9: "Coach Education",
}
SUPPORTING_RESOURCES = {
    2: (
        "L1715_Sim02_TS 02.01_Rowing-Basics-3.pdf",
        "L1715_Sim02_TS 02.02_Understanding Rowing.htm",
    ),
    16: (
        "L1715_Sim16_Instructor Slides.pptx",
        "L1715_Sim16_Notecards for HCP Evaluation 1.docx",
        "L1715_Sim16_Notecards for HCP Evaluation 2.docx",
    ),
    17: ("L1715_Sim17_Instructor Slides.pptx",),
    21: (
        "L1715_Sim21 HCP SCAT6 Rain Shoemaker.pdf",
        "L1715_Sim21 SP SCAT6 Rain Shoemaker.pdf",
    ),
    22: ("L1715_Sim22 SCAT6 Rain Shoemaker 96 Hours.pdf",),
}


def manuscript_sources():
    """Return the single primary instructor manuscript for each simulation."""
    sources = {}
    supporting_names = {
        filename
        for filenames in SUPPORTING_RESOURCES.values()
        for filename in filenames
    }
    for path in sorted(MANUSCRIPTS.glob("L1715_Sim*.docx")):
        match = SIM_RE.match(path.name)
        if not match or path.name in supporting_names:
            continue
        number = int(match.group(1))
        if number in sources:
            raise ValueError(
                f"Multiple primary manuscripts found for simulation {number}: "
                f"{sources[number].name}, {path.name}"
            )
        sources[number] = path
    return [sources[number] for number in sorted(sources)]


def text_of(node):
    return "".join(node.xpath(".//w:t/text()", namespaces=NS))


def remove_prefix(node, prefix):
    remaining = prefix
    for text_node in node.xpath(".//w:t", namespaces=NS):
        value = text_node.text or ""
        if not remaining:
            break
        count = min(len(value), len(remaining))
        if value[:count] != remaining[:count]:
            return
        text_node.text = value[count:]
        remaining = remaining[count:]


def body_children(path):
    with ZipFile(path) as archive:
        root = etree.fromstring(archive.read("word/document.xml"))
    return root, root.find("w:body", NS)


def next_nonempty_text(children, start):
    for node in children[start:]:
        value = text_of(node).strip()
        if value:
            return value
    return ""


def download_sections(path):
    _root, body = body_children(path)
    children = list(body)
    sections = []
    for index, node in enumerate(children):
        marker = text_of(node).strip()
        if not BEGIN_RE.search(marker):
            continue
        end = next(
            (candidate for candidate in range(index + 1, len(children))
             if text_of(children[candidate]).strip().startswith("\\qqEND downloadable content")
             or BEGIN_RE.search(text_of(children[candidate]))),
            len(children),
        )
        selected = [deepcopy(item) for item in children[index + 1:end]]
        if "<title>" in marker:
            inline = deepcopy(node)
            remove_prefix(inline, marker.split("<title>", 1)[0])
            selected.insert(0, inline)
        first = next_nonempty_text(selected, 0)
        title_match = re.match(r"<(?:title|b|a)>(.*)", first)
        button_match = BUTTON_RE.search(marker.rstrip("\\"))
        title = title_match.group(1).strip() if title_match else ""
        button = button_match.group(1).strip().rstrip("\\") if button_match else title
        if not title:
            title = button
        if title and selected:
            sections.append({"button": button or title, "title": title, "nodes": selected})
    return sections


def clean_download_nodes(nodes):
    cleaned = []
    in_note = False
    for node in nodes:
        value = text_of(node).strip()
        if value.startswith("\\qqID:") or value.startswith("\\qqPSM:"):
            in_note = not NOTE_END_RE.search(value)
            continue
        if in_note:
            if NOTE_END_RE.search(value):
                in_note = False
                continue
            # A new reader-facing block means the preceding production note
            # was a single unclosed instruction, not a multiline note.
            if not PRODUCTION_PREFIX.match(value):
                continue
            in_note = False
        if value.startswith("\\qqINSERT"):
            marker = re.match(r"^\\qqINSERT:?\s+(.*)$", value)
            image_name = PureWindowsPath(marker.group(1).strip()).stem if marker else ""
            if not image_name or not (ROOT / "assets" / "images" / f"{image_name}.png").exists():
                continue
        if value.startswith("\\qq") and not value.startswith("\\qqINSERT"):
            continue
        clone = deepcopy(node)
        clone_text = text_of(clone)
        match = PRODUCTION_PREFIX.match(clone_text)
        if match:
            remove_prefix(clone, match.group(0))
        if text_of(clone).strip() or clone.tag == qn("w:tbl"):
            cleaned.append(clone)
    return cleaned


def document_for_nodes(source, nodes):
    with ZipFile(source) as archive:
        root = etree.fromstring(archive.read("word/document.xml"))
    body = root.find("w:body", NS)
    section_properties = body.find("w:sectPr", NS)
    for node in list(body):
        body.remove(node)
    for node in clean_download_nodes(nodes):
        body.append(node)
    if section_properties is not None:
        body.append(deepcopy(section_properties))
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def slug(value):
    value = value.lower().replace("’", "").replace("'", "")
    value = re.sub(r"[^a-z0-9]+", "-", value).strip("-")
    return value[:90] or "instructor-download"


def write_downloads(source, sim_number):
    output_dir = DOWNLOADS / f"simulation-{sim_number}"
    output_dir.mkdir(parents=True, exist_ok=True)
    results = []
    used = set()
    for section in download_sections(source):
        if section["button"] == "Information for Proctor":
            # These markers contain production assembly directions, not reader-facing copy.
            continue
        nodes = clean_download_nodes(section["nodes"])
        meaningful = [text_of(node).strip() for node in nodes if text_of(node).strip()]
        if len(meaningful) <= 1:
            continue
        base = slug(section["button"])
        filename = f"{base}.docx"
        suffix = 2
        while filename in used:
            filename = f"{base}-{suffix}.docx"
            suffix += 1
        used.add(filename)
        output = output_dir / filename
        write_docx(source, output, document_for_nodes(source, section["nodes"]))
        finalize_document(output)
        results.append((section["button"], filename))
    supporting = [
        MANUSCRIPTS / filename
        for filename in SUPPORTING_RESOURCES.get(sim_number, ())
        if (MANUSCRIPTS / filename).exists()
    ]
    packaged_supporting = []
    for path in supporting:
        if path.suffix.lower() == ".docx":
            packaged = output_dir / path.name
            copy2(path, packaged)
            stamp_footer(packaged)
            packaged_supporting.append(packaged)
        else:
            packaged_supporting.append(path)
    if results or supporting:
        zip_path = output_dir / "all-instructor-downloads.zip"
        with ZipFile(zip_path, "w", ZIP_DEFLATED) as archive:
            for _label, filename in results:
                archive.write(output_dir / filename, filename)
            for path in packaged_supporting:
                archive.write(path, path.name)
    return results


def cell_html(cell):
    parts = []
    for paragraph in cell.findall("w:p", NS):
        value = text_of(paragraph)
        if value:
            parts.append(escape(PRODUCTION_PREFIX.sub("", value)))
    return "<br>".join(parts)


def table_html(node):
    rows = []
    for row_index, row in enumerate(node.findall("w:tr", NS)):
        cells = row.findall("w:tc", NS)
        if row_index == 0:
            markup = "".join(
                f'<th scope="col" data-manuscript-block>{cell_html(cell)}</th>'
                for cell in cells
            )
        else:
            markup = "".join(
                f"<td data-manuscript-block>{cell_html(cell)}</td>" for cell in cells
            )
        rows.append("<tr>" + markup + "</tr>")
    header = rows[0] if rows else ""
    body_rows = "".join(rows[1:])
    return (
        '<div class="table-scroll" role="region" aria-label="Scrollable data table" tabindex="0">'
        '<table class="manuscript-table"><thead>' + header + "</thead><tbody>"
        + body_rows + "</tbody></table></div>"
    )


def image_for_marker(value):
    marker = re.match(r"^\\qqINSERT:?\s+(.*)$", value)
    if not marker or not marker.group(1).strip():
        return None
    stem = PureWindowsPath(marker.group(1).strip()).stem
    candidate = ROOT / "assets" / "images" / f"{stem}.png"
    return candidate if candidate.exists() else None


def main_content(path):
    _root, body = body_children(path)
    output = []
    in_download = False
    in_note = False
    active_list = None

    def close_list():
        nonlocal active_list
        if active_list:
            output.append(f"</{active_list}>")
            active_list = None

    for node in list(body):
        value = text_of(node).strip()
        value = re.sub(r"^(?:PHOTO HERE\s*)+", "", value).strip()
        if BEGIN_RE.search(value):
            close_list()
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
            # Do not let a malformed or unclosed production note consume a
            # subsequent manuscript section.
            if not PRODUCTION_PREFIX.match(value):
                continue
            in_note = False
        if value.startswith("\\qqINSERT"):
            close_list()
            image = image_for_marker(value)
            if image:
                output.append(
                    '<figure class="manuscript-figure"><img src="../assets/images/'
                    f'{escape(image.name)}" alt=""></figure>'
                )
            continue
        if not value or value.startswith("Navigation menu/button") or value.startswith("\\qq"):
            continue
        if node.tag == qn("w:tbl"):
            close_list()
            output.append(table_html(node))
            continue
        marker = PRODUCTION_PREFIX.match(value)
        tag_name = marker.group(1) if marker else ""
        clean = PRODUCTION_PREFIX.sub("", value)
        if tag_name in ("cn", "ct"):
            continue
        html_tag = HEADING_TAGS.get(tag_name, "p")
        if html_tag.startswith("h"):
            close_list()
            output.append(f"<{html_tag} data-manuscript-block>{escape(clean)}</{html_tag}>")
            list_config = LIST_SECTIONS.get(clean)
            if list_config:
                active_list, list_class = list_config
                output.append(f'<{active_list} class="manuscript-list {list_class}">')
            continue
        if active_list and html_tag == "p":
            number_match = SOURCE_LIST_NUMBER_RE.match(clean)
            if number_match:
                prefix = number_match.group(1)
                rendered = (
                    f'<span class="source-list-number">{escape(prefix)}</span>'
                    f'{escape(clean[len(prefix):])}'
                )
            else:
                rendered = escape(clean)
            output.append(f"<li data-manuscript-block>{rendered}</li>")
            continue
        close_list()
        css = ' class="figure-caption"' if tag_name == "fc" else ""
        output.append(f"<{html_tag}{css} data-manuscript-block>{escape(clean)}</{html_tag}>")
    close_list()
    return output


def doc_title(path, fallback):
    for paragraph in Document(path).paragraphs:
        value = paragraph.text.strip()
        if value.startswith("<ct>"):
            title = re.sub(r"^Simulation \d+:?\s*", "", value.removeprefix("<ct>"))
            return title or fallback
    return fallback


def page_shell(title, kicker, content):
    return f'''<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>{escape(title)}</title><link rel="stylesheet" href="page.css"></head>
<body data-manuscript><header class="lesson-header"><p class="kicker">{escape(kicker)}</p><h1>{escape(title)}</h1></header>
<main class="page">{content}</main></body></html>
'''


def build_manuscript_page(source, number, fallback_title):
    title = doc_title(source, fallback_title)
    blocks = main_content(source)
    cards = []
    current = []
    for block in blocks:
        if block.startswith("<h2") and current:
            cards.append('<section class="content-card">' + "".join(current) + "</section>")
            current = []
        current.append(block)
    if current:
        cards.append('<section class="content-card">' + "".join(current) + "</section>")
    downloads = write_downloads(source, number)
    supporting = [
        filename
        for filename in SUPPORTING_RESOURCES.get(number, ())
        if (MANUSCRIPTS / filename).exists()
    ]
    if downloads or supporting:
        links = ['<a class="download-card featured" href="../assets/downloads/simulation-{0}/all-instructor-downloads.zip" download><span class="file-icon" aria-hidden="true">ZIP</span><span><strong>All Instructor Downloads</strong><small>ZIP archive</small></span><span class="download-arrow" aria-hidden="true">↓</span></a>'.format(number)]
        for label, filename in downloads:
            links.append(f'<a class="download-card" href="../assets/downloads/simulation-{number}/{escape(filename)}" download><span class="file-icon" aria-hidden="true">DOCX</span><span><strong>{escape(label)}</strong><small>Word document</small></span><span class="download-arrow" aria-hidden="true">↓</span></a>')
        for filename in supporting:
            extension = Path(filename).suffix.removeprefix(".").upper()
            label = re.sub(r"^L1715_Sim\d{2}[_ ]*", "", Path(filename).stem).replace("_", " ")
            href = (
                f"../assets/downloads/simulation-{number}/{filename}"
                if extension == "DOCX"
                else f"../assets/Manuscripts/{filename}"
            )
            links.append(
                f'<a class="download-card" href="{escape(href)}" download>'
                f'<span class="file-icon" aria-hidden="true">{escape(extension)}</span><span><strong>{escape(label)}</strong>'
                f'<small>Supporting instructor resource</small></span><span class="download-arrow" aria-hidden="true">&darr;</span></a>'
            )
        cards.append('<section class="content-card"><h2>Instructor Downloads</h2><div class="download-grid activity-grid">' + "".join(links) + "</div></section>")
    if number == 9:
        cards.insert(0, '<section class="content-card tint"><h2>Instructor Manuscript Status</h2><p>This is a partial instructor manuscript.</p></section>')
    page = PAGES / f"simulation-{number}.html"
    page.write_text(page_shell(title, f"Simulation {number}", "".join(cards)), encoding="utf-8")
    return title, len(downloads)


def build_introduction():
    source = MANUSCRIPTS / "L1715_Introduction.docx"
    stamp_footer(DOWNLOADS / "copyright-page-placeholder.docx")
    page_path = PAGES / "introduction.html"
    try:
        blocks = main_content(source)
    except PermissionError:
        # OneDrive or Word can temporarily lock a hydrated manuscript. Keep the
        # last verified page usable and attach the navigation enhancement.
        rendered = page_path.read_text(encoding="utf-8")
        if "introduction-nav.js" not in rendered:
            rendered = rendered.replace(
                "</body>", '<script src="introduction-nav.js"></script></body>'
            )
            page_path.write_text(rendered, encoding="utf-8")
        print(f"LOCKED: retained {page_path.relative_to(ROOT)} and added section navigation")
        return
    headings = []
    anchored_blocks = []
    for block in blocks:
        match = re.match(r'(<h2)([^>]*>)(.*?)(</h2>)$', block)
        if not match:
            anchored_blocks.append(block)
            continue
        label = unescape(re.sub(r"<[^>]+>", "", match.group(3)))
        anchor = f"introduction-{slug(label)}"
        headings.append((label, anchor))
        anchored_blocks.append(
            f'{match.group(1)} id="{anchor}"{match.group(2)}{match.group(3)}{match.group(4)}'
        )

    jump_links = "".join(
        f'<a class="section-jump" href="#{anchor}">{escape(label)}</a>'
        for label, anchor in headings
    )
    jump_links += '<a class="section-jump" href="#introduction-copyright">Copyright</a>'
    navigation = (
        '<details class="content-card introduction-navigation" id="introduction-navigation" open>'
        '<summary>Section navigation</summary>'
        f'<div class="section-jump-grid">{jump_links}</div></details>'
    )

    cards, current = [], []
    for block in anchored_blocks:
        if block.startswith("<h2") and current:
            cards.append('<section class="content-card">' + "".join(current) + "</section>")
            current = []
        current.append(block)
    if current:
        cards.append('<section class="content-card">' + "".join(current) + "</section>")
    for index in range(1, len(cards)):
        cards[index] = cards[index].replace(
            "</section>",
            '<a class="section-return" href="#introduction-navigation">↑ Return to section navigation</a></section>',
            1,
        )
    copyright_card = (
        '<section class="content-card" id="introduction-copyright"><h2>Copyright</h2>'
        '<div class="download-grid activity-grid"><a class="download-card featured" '
        'href="../assets/downloads/copyright-page-placeholder.docx" download>'
        '<span class="file-icon" aria-hidden="true">DOCX</span><span>'
        '<strong>Download Copyright Page</strong><small>Placeholder Word document</small>'
        '</span><span class="download-arrow" aria-hidden="true">↓</span></a></div>'
        '<a class="section-return" href="#introduction-navigation">'
        '↑ Return to section navigation</a></section>'
    )
    rendered = page_shell(
        "Introduction", "Instructor guide", navigation + "".join(cards) + copyright_card
    )
    rendered = rendered.replace("</body>", '<script src="introduction-nav.js"></script></body>')
    page_path.write_text(rendered, encoding="utf-8")


def build_debriefing():
    source = MANUSCRIPTS / "L1715_Debriefing Methods.docx"
    blocks = [
        block
        for block in main_content(source)
        if block != "<p data-manuscript-block>Debriefing Methods</p>"
    ]
    content = '<section class="content-card">' + "".join(blocks) + "</section>"
    (PAGES / "debriefing-methods.html").write_text(page_shell("Debriefing Methods", "Instructor resource", content), encoding="utf-8")


def resource_page(filename, title, label):
    content = f'<section class="content-card"><div class="download-grid"><a class="download-card featured" href="../assets/Manuscripts/{escape(filename)}" download><span class="file-icon" aria-hidden="true">{escape(label)}</span><span><strong>{escape(title)}</strong><small>{escape(filename)}</small></span><span class="download-arrow" aria-hidden="true">↓</span></a></div></section>'
    return page_shell(title, "Instructor resource", content)


def build_resource_pages():
    resources = (
        ("simulation-checklist.html", "L1715_Simulation Checklist.xlsx", "Simulation Checklist", "XLSX"),
        ("simulation-finder.html", "L1715_Simulation Finder.xlsx", "Simulation Finder", "XLSX"),
    )
    for page, source, title, label in resources:
        (PAGES / page).write_text(resource_page(source, title, label), encoding="utf-8")


def navigation_titles():
    source = (ROOT / "data" / "navigation.js").read_text(encoding="utf-8")
    return {
        int(number): title
        for number, title in re.findall(
            r'id: "simulation-(\d+)", title: "Simulation \d+: ([^"]+)"', source
        )
    }


def build_pending_pages(available):
    titles = navigation_titles()
    for number in range(1, 33):
        if number in available:
            continue
        title = titles.get(number, f"Simulation {number}")
        resources = SUPPORTING_RESOURCES.get(number, ())
        links = []
        for filename in resources:
            extension = Path(filename).suffix.removeprefix(".").upper()
            links.append(
                f'<a class="download-card" href="../assets/Manuscripts/{escape(filename)}" download>'
                f'<span class="file-icon" aria-hidden="true">{escape(extension)}</span><span><strong>{escape(filename)}</strong>'
                f'<small>Supporting instructor resource</small></span><span class="download-arrow" aria-hidden="true">↓</span></a>'
            )
        resource_section = ""
        if links:
            resource_section = '<section class="content-card"><h2>Available Instructor Resources</h2><div class="download-grid">' + "".join(links) + "</div></section>"
        notice = '<section class="content-card tint"><h2>Instructor Manuscript Status</h2><p>The instructor manuscript for this simulation is not yet available.</p></section>'
        (PAGES / f"simulation-{number}.html").write_text(
            page_shell(title, f"Simulation {number}", notice + resource_section), encoding="utf-8"
        )


def main():
    requested = {int(arg) for arg in sys.argv[1:]}
    titles = navigation_titles()
    build_introduction()
    build_debriefing()
    build_resource_pages()
    built = []
    for source in manuscript_sources():
        match = SIM_RE.match(source.name)
        if not match:
            continue
        number = int(match.group(1))
        if requested and number not in requested:
            continue
        fallback = TITLE_OVERRIDES.get(number, titles.get(number, f"Simulation {number}"))
        title, count = build_manuscript_page(source, number, fallback)
        built.append(number)
        print(f"BUILT: simulation-{number} ({title}; {count} downloads)")
    if not requested:
        build_pending_pages(set(built))
    print("AVAILABLE:", ", ".join(map(str, built)))


if __name__ == "__main__":
    main()
