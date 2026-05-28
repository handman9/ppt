from __future__ import annotations

import cgi
import html
import io
import os
import re
import sys
import time
import zipfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import quote, unquote

from lxml import etree as ET


ROOT = Path(__file__).resolve().parent
OUTPUT_DIR = ROOT / "outputs"
OUTPUT_DIR.mkdir(exist_ok=True)

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

R_NS = NS["r"]
REL_NS = NS["rel"]


HTML = """<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>예배 자막 PPTX 생성기</title>
  <style>
    :root {
      color-scheme: dark;
      font-family: "Malgun Gothic", "Apple SD Gothic Neo", system-ui, sans-serif;
      background: #101113;
      color: #f6f5f0;
    }
    * { box-sizing: border-box; }
    body { margin: 0; min-height: 100vh; background: #101113; }
    main { width: min(980px, calc(100vw - 36px)); margin: 0 auto; padding: 34px 0 42px; }
    header { display: flex; align-items: end; justify-content: space-between; gap: 20px; margin-bottom: 22px; }
    h1 { margin: 0; font-size: 26px; line-height: 1.25; font-weight: 760; letter-spacing: 0; }
    .panel { border: 1px solid #303238; background: #17191d; border-radius: 8px; padding: 22px; }
    label { display: block; font-size: 14px; color: #d8d5cd; margin: 0 0 8px; }
    .field { margin-bottom: 18px; }
    input[type=file], textarea, select {
      width: 100%;
      border: 1px solid #3a3d44;
      background: #0f1012;
      color: #f6f5f0;
      border-radius: 6px;
      padding: 12px 13px;
      font: inherit;
    }
    textarea { min-height: 320px; resize: vertical; line-height: 1.65; }
    .row { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }
    .hint { color: #a9a59b; font-size: 13px; line-height: 1.55; margin-top: 8px; }
    button {
      border: 0;
      background: #f2c14e;
      color: #161616;
      border-radius: 6px;
      padding: 13px 18px;
      font-size: 15px;
      font-weight: 760;
      cursor: pointer;
    }
    button:hover { background: #ffd36b; }
    .error { border-color: #6e3333; background: #251819; color: #ffd4d4; margin-bottom: 16px; }
    footer { color: #8f8b83; font-size: 12px; margin-top: 16px; }
    @media (max-width: 720px) { .row { grid-template-columns: 1fr; } header { display: block; } }
  </style>
</head>
<body>
  <main>
    <header>
      <h1>예배 자막 PPTX 생성기</h1>
    </header>
    {message}
    <form class="panel" method="post" action="/generate" enctype="multipart/form-data">
      <div class="field">
        <label for="template">샘플 PPTX</label>
        <input id="template" name="template" type="file" accept=".pptx" required>
        <div class="hint">자막 디자인으로 사용할 PPTX입니다. 첫 번째로 텍스트가 들어있는 슬라이드의 가장 큰 텍스트 박스를 자막 위치로 사용합니다.</div>
      </div>
      <div class="row">
        <div class="field">
          <label for="mode">빈 줄 처리</label>
          <select id="mode" name="blank_mode">
            <option value="skip">빈 줄은 건너뛰기</option>
            <option value="slide">빈 줄도 빈 자막 슬라이드로 만들기</option>
          </select>
        </div>
        <div class="field">
          <label for="filename">파일 이름</label>
          <input id="filename" name="filename" type="text" value="예배자막.pptx">
        </div>
      </div>
      <div class="field">
        <label for="lyrics">가사</label>
        <textarea id="lyrics" name="lyrics" placeholder="가사를 한 줄씩 붙여 넣으세요.&#10;한 줄이 슬라이드 한 장이 됩니다." required></textarea>
      </div>
      <button type="submit">PPTX 생성</button>
    </form>
    <footer>생성 방식: 템플릿의 배경, 이미지, 폰트, 위치를 유지하고 자막 텍스트만 줄마다 교체합니다.</footer>
  </main>
</body>
</html>"""


def clean_filename(name: str) -> str:
    name = (name or "예배자막.pptx").strip().replace("/", "_").replace("\\", "_")
    if not name.lower().endswith(".pptx"):
        name += ".pptx"
    return name


def slide_number(name: str) -> int:
    m = re.search(r"slide(\d+)\.xml$", name)
    return int(m.group(1)) if m else 0


def relationship_id(n: int) -> str:
    return f"rId{n}"


def get_rel_type(target_type: str) -> str:
    return f"http://schemas.openxmlformats.org/officeDocument/2006/relationships/{target_type}"


def find_template_slide(zf: zipfile.ZipFile) -> tuple[str, bytes | None]:
    slide_names = sorted(
        [n for n in zf.namelist() if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)],
        key=slide_number,
    )
    for name in slide_names:
        root = ET.fromstring(zf.read(name))
        if root.xpath(".//a:t[normalize-space(.)!='']", namespaces=NS):
            rel_name = f"ppt/slides/_rels/{Path(name).name}.rels"
            return name, zf.read(rel_name) if rel_name in zf.namelist() else None
    if not slide_names:
        raise ValueError("템플릿 PPTX에 슬라이드가 없습니다.")
    name = slide_names[0]
    rel_name = f"ppt/slides/_rels/{Path(name).name}.rels"
    return name, zf.read(rel_name) if rel_name in zf.namelist() else None


def shape_area(shape: ET._Element) -> int:
    xfrm = shape.find(".//a:xfrm", namespaces=NS)
    if xfrm is None:
        return 0
    ext = xfrm.find("a:ext", namespaces=NS)
    if ext is None:
        return 0
    return int(ext.get("cx", "0")) * int(ext.get("cy", "0"))


def set_caption_text(slide_xml: bytes, text: str) -> bytes:
    root = ET.fromstring(slide_xml)
    shapes = root.xpath(".//p:sp[p:txBody]", namespaces=NS)
    text_shapes = [
        shp for shp in shapes
        if shp.xpath(".//a:t[normalize-space(.)!='']", namespaces=NS)
    ]
    candidates = text_shapes or shapes
    if not candidates:
        raise ValueError("템플릿 슬라이드에서 텍스트 박스를 찾지 못했습니다.")

    target = max(candidates, key=shape_area)
    text_nodes = target.xpath(".//a:t", namespaces=NS)
    if not text_nodes:
        raise ValueError("템플릿 텍스트 박스에 편집 가능한 텍스트 노드가 없습니다.")
    text_nodes[0].text = text
    for node in text_nodes[1:]:
        node.text = ""

    c_nv_pr = target.find(".//p:cNvPr", namespaces=NS)
    if c_nv_pr is not None:
        c_nv_pr.set("name", "Caption")

    return ET.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def presentation_non_slide_relation_ids(data: bytes) -> set[str]:
    root = ET.fromstring(data)
    return {
        rel.get("Id", "")
        for rel in root
        if rel.get("Type") != get_rel_type("slide")
    }


def make_slide_relation_ids(existing_ids: set[str], slide_count: int) -> list[str]:
    used_numbers = [
        int(rid[3:])
        for rid in existing_ids
        if rid.startswith("rId") and rid[3:].isdigit()
    ]
    start = max(used_numbers + [0]) + 1
    return [relationship_id(start + i) for i in range(slide_count)]


def update_presentation_xml(data: bytes, slide_rel_ids: list[str]) -> bytes:
    root = ET.fromstring(data)
    sld_id_lst = root.find("p:sldIdLst", namespaces=NS)
    if sld_id_lst is None:
        sld_id_lst = ET.SubElement(root, f"{{{NS['p']}}}sldIdLst")
    for child in list(sld_id_lst):
        sld_id_lst.remove(child)
    for i, rel_id in enumerate(slide_rel_ids):
        item = ET.SubElement(sld_id_lst, f"{{{NS['p']}}}sldId")
        item.set("id", str(256 + i))
        item.set(f"{{{R_NS}}}id", rel_id)
    return ET.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def update_presentation_rels(data: bytes, slide_rel_ids: list[str]) -> bytes:
    root = ET.fromstring(data)
    for rel in list(root):
        if rel.get("Type") == get_rel_type("slide"):
            root.remove(rel)

    for i, rel_id in enumerate(slide_rel_ids):
        rel = ET.Element(f"{{{REL_NS}}}Relationship")
        rel.set("Id", rel_id)
        rel.set("Type", get_rel_type("slide"))
        rel.set("Target", f"slides/slide{i + 1}.xml")
        root.insert(i, rel)

    return ET.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def update_content_types(data: bytes, slide_count: int) -> bytes:
    root = ET.fromstring(data)
    slide_ct = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
    for item in list(root):
        if item.tag == f"{{{NS['ct']}}}Override" and item.get("ContentType") == slide_ct:
            root.remove(item)
    for i in range(slide_count):
        item = ET.Element(f"{{{NS['ct']}}}Override")
        item.set("PartName", f"/ppt/slides/slide{i + 1}.xml")
        item.set("ContentType", slide_ct)
        root.append(item)
    return ET.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def update_app_xml(data: bytes, slide_count: int) -> bytes:
    try:
        root = ET.fromstring(data)
    except ET.XMLSyntaxError:
        return data
    app_ns = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
    slides = root.find(f"{{{app_ns}}}Slides")
    if slides is not None:
        slides.text = str(slide_count)
        return ET.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    return data


def generate_pptx(template_bytes: bytes, lyric_lines: list[str]) -> bytes:
    if not lyric_lines:
        raise ValueError("생성할 가사 줄이 없습니다.")

    input_io = io.BytesIO(template_bytes)
    output_io = io.BytesIO()

    with zipfile.ZipFile(input_io, "r") as zin:
        template_slide_name, template_rels = find_template_slide(zin)
        template_slide_xml = zin.read(template_slide_name)
        slide_rel_ids = make_slide_relation_ids(
            presentation_non_slide_relation_ids(zin.read("ppt/_rels/presentation.xml.rels")),
            len(lyric_lines),
        )
        skip = {
            n for n in zin.namelist()
            if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)
            or re.fullmatch(r"ppt/slides/_rels/slide\d+\.xml\.rels", n)
            or re.fullmatch(r"ppt/notesSlides/notesSlide\d+\.xml", n)
            or re.fullmatch(r"ppt/notesSlides/_rels/notesSlide\d+\.xml\.rels", n)
        }

        with zipfile.ZipFile(output_io, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                if info.filename in skip:
                    continue
                data = zin.read(info.filename)
                if info.filename == "ppt/presentation.xml":
                    data = update_presentation_xml(data, slide_rel_ids)
                elif info.filename == "ppt/_rels/presentation.xml.rels":
                    data = update_presentation_rels(data, slide_rel_ids)
                elif info.filename == "[Content_Types].xml":
                    data = update_content_types(data, len(lyric_lines))
                elif info.filename == "docProps/app.xml":
                    data = update_app_xml(data, len(lyric_lines))
                zout.writestr(info, data)

            for idx, line in enumerate(lyric_lines, start=1):
                zout.writestr(
                    f"ppt/slides/slide{idx}.xml",
                    set_caption_text(template_slide_xml, line),
                )
                if template_rels is not None:
                    zout.writestr(f"ppt/slides/_rels/slide{idx}.xml.rels", template_rels)

    return output_io.getvalue()


def render_page(message: str = "") -> bytes:
    return HTML.replace("{message}", message).encode("utf-8")


class Handler(BaseHTTPRequestHandler):
    def do_GET(self) -> None:
        if self.path.startswith("/outputs/"):
            return self.serve_output()
        self.respond(200, render_page(), "text/html; charset=utf-8")

    def do_POST(self) -> None:
        if self.path != "/generate":
            self.respond(404, b"Not found", "text/plain; charset=utf-8")
            return
        try:
            form = cgi.FieldStorage(
                fp=self.rfile,
                headers=self.headers,
                environ={
                    "REQUEST_METHOD": "POST",
                    "CONTENT_TYPE": self.headers.get("Content-Type", ""),
                },
            )
            template = form["template"]
            if not getattr(template, "file", None):
                raise ValueError("샘플 PPTX 파일을 선택하세요.")
            template_bytes = template.file.read()
            lyrics = form.getfirst("lyrics", "")
            blank_mode = form.getfirst("blank_mode", "skip")
            lines = lyrics.replace("\r\n", "\n").replace("\r", "\n").split("\n")
            if blank_mode == "skip":
                lines = [line.strip() for line in lines if line.strip()]
            else:
                lines = [line.strip() for line in lines]
            output = generate_pptx(template_bytes, lines)
            requested_name = clean_filename(form.getfirst("filename", "예배자막.pptx"))
            stem = Path(requested_name).stem
            safe_name = clean_filename(f"{stem}_{time.strftime('%Y%m%d_%H%M%S')}.pptx")
            out_path = OUTPUT_DIR / safe_name
            out_path.write_bytes(output)
            self.send_response(303)
            self.send_header("Location", f"/outputs/{quote(safe_name)}")
            self.end_headers()
        except Exception as exc:
            msg = f'<div class="panel error">생성 중 오류가 났습니다: {html.escape(str(exc))}</div>'
            self.respond(400, render_page(msg), "text/html; charset=utf-8")

    def serve_output(self) -> None:
        name = self.path.split("/outputs/", 1)[1].split("?", 1)[0]
        name = os.path.basename(unquote(name))
        path = OUTPUT_DIR / name
        if not path.exists():
            self.respond(404, b"File not found", "text/plain; charset=utf-8")
            return
        data = path.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation")
        self.send_header("Content-Disposition", f'attachment; filename="{quote(name)}"')
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def respond(self, status: int, body: bytes, content_type: str) -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, fmt: str, *args: object) -> None:
        sys.stderr.write("%s - %s\n" % (self.address_string(), fmt % args))


def main() -> None:
    port = int(os.environ.get("PORT", "8765"))
    host = "0.0.0.0" if os.environ.get("RENDER") else "127.0.0.1"
    server = ThreadingHTTPServer((host, port), Handler)
    print(f"예배 자막 PPTX 생성기: http://{host}:{port}")
    server.serve_forever()


if __name__ == "__main__":
    main()
