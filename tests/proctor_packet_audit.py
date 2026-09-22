"""Verify generated Information for Proctor packets and ZIP packaging."""

from pathlib import Path
import sys
from zipfile import ZipFile

from lxml import etree


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from build_instructor_version import PROCTOR_PACKETS  # noqa: E402


NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def document_root(path):
    with ZipFile(path) as archive:
        return etree.fromstring(archive.read("word/document.xml"))


def document_text(path):
    root = document_root(path)
    return " ".join(root.xpath("//w:t/text()", namespaces=NS)).split()


def document_blocks(path):
    root = document_root(path)
    return [
        " ".join(node.xpath(".//w:t/text()", namespaces=NS)).strip()
        for node in root.xpath("//w:body/*[not(self::w:sectPr)]", namespaces=NS)
        if " ".join(node.xpath(".//w:t/text()", namespaces=NS)).strip()
    ]


def page_break_count(path):
    root = document_root(path)
    return len(root.xpath("//w:br[@w:type='page']", namespaces=NS))


def main():
    failures = []
    packet_count = 0
    for simulation, packets in sorted(PROCTOR_PACKETS.items()):
        output_dir = ROOT / "assets" / "downloads" / f"simulation-{simulation}"
        page_source = (ROOT / "pages" / f"simulation-{simulation}.html").read_text(
            encoding="utf-8"
        )
        zip_path = output_dir / "all-instructor-downloads.zip"
        if not zip_path.exists():
            failures.append(f"simulation-{simulation}: missing all-instructor-downloads.zip")
            continue
        with ZipFile(zip_path) as archive:
            zip_names = set(archive.namelist())

        for label, filename, components in packets:
            packet_count += 1
            packet_path = output_dir / filename
            if not packet_path.exists():
                failures.append(f"simulation-{simulation}: missing {filename}")
                continue
            if filename not in zip_names:
                failures.append(f"simulation-{simulation}: ZIP is missing {filename}")
            if f'href="../assets/downloads/simulation-{simulation}/{filename}"' not in page_source:
                failures.append(f"simulation-{simulation}: page is missing the {filename} link")

            try:
                packet_words = document_text(packet_path)
            except Exception as error:
                failures.append(f"simulation-{simulation}: cannot read {filename}: {error}")
                continue
            if packet_words[: len(label.split())] != label.split():
                failures.append(f"simulation-{simulation}: {filename} has the wrong title")

            packet_text = " ".join(packet_words)
            packet_blocks = document_blocks(packet_path)
            first_component_blocks = document_blocks(output_dir / components[0])
            if len(packet_blocks) < 2 or packet_blocks[1] != first_component_blocks[0]:
                failures.append(
                    f"simulation-{simulation}: {filename} starts its content on a cover page"
                )

            expected_page_breaks = len(components) - 1 + sum(
                page_break_count(output_dir / component) for component in components
            )
            if page_break_count(packet_path) != expected_page_breaks:
                failures.append(
                    f"simulation-{simulation}: {filename} has an unexpected page break"
                )
            for component in components:
                component_path = output_dir / component
                if not component_path.exists():
                    failures.append(
                        f"simulation-{simulation}: missing packet component {component}"
                    )
                    continue
                component_text = " ".join(document_text(component_path))
                if component_text not in packet_text:
                    failures.append(
                        f"simulation-{simulation}: {filename} does not contain {component}"
                    )

    if packet_count != 38:
        failures.append(f"expected 38 proctor packets; found {packet_count}")
    if set(PROCTOR_PACKETS) & {11, 14, 16, 17}:
        failures.append("simulations 11, 14, 16, and 17 must not have generated packets")

    if failures:
        print(f"FAIL: {len(failures)} proctor packet issue(s)")
        for failure in failures:
            print(f"  {failure}")
        return 1
    print("PASS: 38 proctor packets are complete, linked, and included in 28 ZIP files")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
