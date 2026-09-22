from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
import sys


root = Path.cwd()
sys.path.insert(0, str(root / "scripts"))

from build_instructor_version import (  # noqa: E402
    build_manuscript_page,
    doc_title,
    download_manifest,
    download_specs,
)
from generate_simulation_downloads import finalize_document, write_docx  # noqa: E402
from build_instructor_version import document_for_nodes  # noqa: E402


source = root / "assets" / "Manuscripts" / "L1715_Sim06_Dyspnea.docx"
output_dir = root / "assets" / "downloads" / "simulation-6"
target_name = "instructions-for-sp-anaphylaxis.docx"

build_manuscript_page(
    source,
    6,
    doc_title(source, "Dyspnea"),
    write_download_assets=False,
)

matching = [spec for spec in download_specs(source) if spec[2] == target_name]
if len(matching) != 1:
    raise RuntimeError(f"Expected one {target_name} download specification, found {len(matching)}")
section, _label, filename = matching[0]
output = output_dir / filename
write_docx(source, output, document_for_nodes(source, section["nodes"]))
finalize_document(output)

zip_path = output_dir / "all-instructor-downloads.zip"
with ZipFile(zip_path, "w", ZIP_DEFLATED) as archive:
    for _label, download_name in download_manifest(source):
        archive.write(output_dir / download_name, download_name)
