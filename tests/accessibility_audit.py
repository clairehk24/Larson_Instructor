"""Static WCAG 2.0 Level AA checks for the instructor web interface."""

from pathlib import Path
import re
import sys

from lxml import html


ROOT = Path(__file__).resolve().parents[1]
HTML_FILES = [ROOT / "index.html", *sorted((ROOT / "pages").glob("*.html"))]
FOCUSABLE_XPATH = (
    ".//a[@href] | .//button | .//input[not(@type='hidden')] | .//select | "
    ".//textarea | .//summary | .//iframe | .//*[@tabindex]"
)


def normalized(value):
    return " ".join((value or "").split())


def visible_text(node):
    parts = []

    def collect(current):
        if current.get("aria-hidden") == "true" or current.get("hidden") is not None:
            return
        if current.text:
            parts.append(current.text)
        for child in current:
            collect(child)
            if child.tail:
                parts.append(child.tail)

    collect(node)
    return normalized(" ".join(parts))


def accessible_name(node, document):
    if normalized(node.get("aria-label")):
        return normalized(node.get("aria-label"))
    labelled_by = normalized(node.get("aria-labelledby"))
    if labelled_by:
        labels = []
        for identifier in labelled_by.split():
            match = document.xpath(f"//*[@id={identifier!r}]")
            if match:
                labels.append(visible_text(match[0]))
        if normalized(" ".join(labels)):
            return normalized(" ".join(labels))
    identifier = node.get("id")
    if identifier:
        labels = document.xpath(f"//label[@for={identifier!r}]")
        if labels and visible_text(labels[0]):
            return visible_text(labels[0])
    return visible_text(node)


def audit_html(path):
    failures = []
    source = path.read_text(encoding="utf-8")
    document = html.fromstring(source)
    relative = path.relative_to(ROOT)

    if normalized(document.get("lang")) == "":
        failures.append("missing the document language")
    titles = document.xpath("//head/title")
    if len(titles) != 1 or not visible_text(titles[0]):
        failures.append("needs one nonempty page title")
    viewports = document.xpath("//meta[translate(@name,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')='viewport']")
    if len(viewports) != 1:
        failures.append("needs one viewport declaration")
    elif re.search(r"user-scalable\s*=\s*no|maximum-scale\s*=\s*1(?:\.0+)?(?:\D|$)", viewports[0].get("content", ""), re.I):
        failures.append("viewport prevents text zoom")

    mains = document.xpath("//main")
    if len(mains) != 1:
        failures.append(f"has {len(mains)} main landmarks; expected 1")
    h1s = document.xpath("//h1")
    if len(h1s) != 1 or not visible_text(h1s[0]):
        failures.append(f"has {len(h1s)} usable h1 headings; expected 1")
    headings = document.xpath("//h1|//h2|//h3|//h4|//h5|//h6")
    previous_level = 0
    for heading in headings:
        level = int(heading.tag[1])
        if not visible_text(heading):
            failures.append("contains an empty heading")
        if previous_level and level > previous_level + 1:
            failures.append(
                f"heading level jumps from h{previous_level} to h{level}: {visible_text(heading)!r}"
            )
        previous_level = level

    identifiers = [value for value in document.xpath("//*[@id]/@id") if value]
    duplicates = sorted({value for value in identifiers if identifiers.count(value) > 1})
    if duplicates:
        failures.append(f"contains duplicate IDs: {', '.join(duplicates)}")

    for image in document.xpath("//img"):
        if image.get("alt") is None:
            failures.append(f"image is missing alt text: {image.get('src', '<unknown>')}")
    for frame in document.xpath("//iframe"):
        if not normalized(frame.get("title")):
            failures.append("iframe is missing a descriptive title")
    for control in document.xpath("//button|//input[not(@type='hidden')]|//select|//textarea"):
        if not accessible_name(control, document):
            failures.append(f"{control.tag} control has no accessible name")
    for link in document.xpath("//a[@href]"):
        if not normalized(link.get("href")):
            failures.append("link has an empty href")
        if not accessible_name(link, document):
            failures.append(f"link has no accessible name: {link.get('href')!r}")

    for value in document.xpath("//*[@tabindex]/@tabindex"):
        try:
            if int(value) > 0:
                failures.append(f"uses positive tabindex={value}")
        except ValueError:
            failures.append(f"uses invalid tabindex={value!r}")

    for table in document.xpath("//table"):
        headers = table.xpath(".//th")
        if not headers:
            failures.append("data table has no header cells")
        for header in headers:
            if header.get("scope") not in {"col", "row", "colgroup", "rowgroup"}:
                failures.append(f"table header lacks scope: {visible_text(header)!r}")
        wrappers = table.xpath("ancestor::*[contains(concat(' ', normalize-space(@class), ' '), ' table-scroll ')][1]")
        if not wrappers:
            failures.append("table is missing its horizontal-scroll container")
        else:
            wrapper = wrappers[0]
            if wrapper.get("tabindex") != "0":
                failures.append("scrollable table is not keyboard focusable")
            if wrapper.get("role") != "region" or not accessible_name(wrapper, document):
                failures.append("scrollable table region has no accessible name")

    for hidden in document.xpath("//*[@aria-hidden='true']"):
        if hidden.xpath(FOCUSABLE_XPATH):
            failures.append("aria-hidden content contains a focusable element")

    return [(str(relative), message) for message in failures]


def channel(value):
    value /= 255
    return value / 12.92 if value <= 0.04045 else ((value + 0.055) / 1.055) ** 2.4


def luminance(color):
    red, green, blue = (int(color[index:index + 2], 16) for index in (1, 3, 5))
    return 0.2126 * channel(red) + 0.7152 * channel(green) + 0.0722 * channel(blue)


def contrast(foreground, background):
    lighter, darker = sorted((luminance(foreground), luminance(background)), reverse=True)
    return (lighter + 0.05) / (darker + 0.05)


def css_variable(path, name):
    source = path.read_text(encoding="utf-8")
    match = re.search(rf"{re.escape(name)}\s*:\s*(#[0-9a-fA-F]{{6}})", source)
    if not match:
        raise ValueError(f"{path.relative_to(ROOT)} is missing {name}")
    return match.group(1)


def audit_contrast():
    failures = []
    for path in (ROOT / "styles.css", ROOT / "pages" / "page.css"):
        pairs = (
            ("--ink", "#ffffff", 4.5),
            ("--muted", "#ffffff", 4.5),
            ("--accent-dark", "#ffffff", 4.5),
            ("--accent-dark", css_variable(path, "--accent-soft"), 4.5),
        )
        for variable, background, minimum in pairs:
            foreground = css_variable(path, variable)
            ratio = contrast(foreground, background)
            if ratio < minimum:
                failures.append(
                    (str(path.relative_to(ROOT)), f"{variable} contrast is {ratio:.2f}:1; needs {minimum}:1")
                )
    return failures


def main():
    failures = []
    for path in HTML_FILES:
        failures.extend(audit_html(path))
    failures.extend(audit_contrast())
    if failures:
        print(f"FAIL: {len(failures)} accessibility issue(s)")
        for path, message in failures:
            print(f"  {path}: {message}")
        return 1
    print(f"PASS: {len(HTML_FILES)} HTML files passed static WCAG 2.0 AA checks")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
