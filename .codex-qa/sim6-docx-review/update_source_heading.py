from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
import os

from lxml import etree


source = Path("assets/Manuscripts/L1715_Sim06_Dyspnea.docx")
temporary = source.with_suffix(".docx.tmp")
document_part = "word/document.xml"
word_namespace = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

with ZipFile(source, "r") as archive:
    document = etree.fromstring(archive.read(document_part))
    matches = [
        paragraph
        for paragraph in document.xpath(".//w:p", namespaces=word_namespace)
        if "".join(paragraph.xpath(".//w:t/text()", namespaces=word_namespace)).startswith(
            "<b>HCP Prebrief"
        )
    ]
    if len(matches) != 1:
        raise RuntimeError(f"Expected one HCP Prebrief production tag, found {len(matches)}")
    text_nodes = matches[0].xpath(".//w:t", namespaces=word_namespace)
    for node in text_nodes:
        if node.text == "b":
            node.text = "a"
            break
    else:
        raise RuntimeError("HCP Prebrief production level was not stored as an isolated run")
    updated_xml = etree.tostring(
        document, xml_declaration=True, encoding="UTF-8", standalone=True
    )

    with ZipFile(temporary, "w", ZIP_DEFLATED) as output:
        for item in archive.infolist():
            payload = updated_xml if item.filename == document_part else archive.read(item.filename)
            output.writestr(item, payload)

os.replace(temporary, source)
