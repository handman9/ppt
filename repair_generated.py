from __future__ import annotations

import sys
import zipfile
from pathlib import Path

from lxml import etree as ET


NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def repair(src: Path, dst: Path) -> None:
    with zipfile.ZipFile(src, "r") as zin:
        relroot = ET.fromstring(zin.read("ppt/_rels/presentation.xml.rels"))
        master_ids = [
            rel.get("Id")
            for rel in relroot
            if rel.get("Type", "").endswith("/slideMaster")
        ]
        if not master_ids:
            raise ValueError("slideMaster relationship not found")
        master_id = master_ids[0]

        with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                data = zin.read(info.filename)
                if info.filename == "ppt/presentation.xml":
                    root = ET.fromstring(data)
                    for item in root.xpath(".//p:sldMasterId", namespaces=NS):
                        item.set(f"{{{NS['r']}}}id", master_id)
                    data = ET.tostring(
                        root,
                        xml_declaration=True,
                        encoding="UTF-8",
                        standalone=True,
                    )
                zout.writestr(info, data)


def main() -> None:
    if len(sys.argv) != 3:
        raise SystemExit("usage: repair_generated.py source.pptx fixed.pptx")
    repair(Path(sys.argv[1]), Path(sys.argv[2]))


if __name__ == "__main__":
    main()
