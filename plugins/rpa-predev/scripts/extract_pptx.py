"""Extract text, tables, notes and images from a 현업 업무 설명 PPT.

Usage:
    python extract_pptx.py 업무설명.pptx [--images out_dir] [--json]

Prints a Markdown dump (## Slide N sections). With --images, every embedded
picture is saved as slideNN_imgK.<ext> so screenshots can be inspected.
Text inside screenshots is NOT extracted — look at the images (or a preview
made with preview_pptx.py) to read menu names, buttons and fields.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


def walk(shapes):
    for shp in shapes:
        if shp.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from walk(shp.shapes)
        else:
            yield shp


def slide_info(slide, idx, img_dir: Path | None):
    texts, tables, images = [], [], []
    shapes = sorted(walk(slide.shapes), key=lambda s: ((s.top or 0), (s.left or 0)))
    for shp in shapes:
        if shp.has_text_frame:
            t = "\n".join(p.text for p in shp.text_frame.paragraphs if p.text.strip())
            if t.strip():
                texts.append(t.strip())
        if getattr(shp, "has_table", False) and shp.has_table:
            tables.append([[c.text.strip() for c in row.cells] for row in shp.table.rows])
        if shp.shape_type == MSO_SHAPE_TYPE.PICTURE:
            name = f"slide{idx:02d}_img{len(images) + 1}.{shp.image.ext}"
            if img_dir:
                (img_dir / name).write_bytes(shp.image.blob)
            images.append(name)
    notes = ""
    if slide.has_notes_slide:
        notes = slide.notes_slide.notes_text_frame.text.strip()
    return {"slide": idx, "texts": texts, "tables": tables, "notes": notes, "images": images}


def to_markdown(info: dict) -> str:
    out = [f"## Slide {info['slide']}"]
    out += info["texts"]
    for t in info["tables"]:
        if not t:
            continue
        out.append("| " + " | ".join(t[0]) + " |")
        out.append("|" + "---|" * len(t[0]))
        out += ["| " + " | ".join(c.replace("\n", " ") for c in r) + " |" for r in t[1:]]
    if info["notes"]:
        out.append(f"> 노트: {info['notes']}")
    if info["images"]:
        out.append(f"[이미지 {len(info['images'])}개: {', '.join(info['images'])}]")
    return "\n".join(out)


def main() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pptx", type=Path)
    ap.add_argument("--images", type=Path, help="embedded images output directory")
    ap.add_argument("--json", action="store_true", help="print JSON instead of Markdown")
    a = ap.parse_args()
    if a.images:
        a.images.mkdir(parents=True, exist_ok=True)
    prs = Presentation(str(a.pptx))
    infos = [slide_info(s, i, a.images) for i, s in enumerate(prs.slides, start=1)]
    if a.json:
        print(json.dumps(infos, ensure_ascii=False, indent=2))
    else:
        print(f"# {a.pptx.name} — {len(infos)} slides\n")
        print("\n\n".join(to_markdown(i) for i in infos))


if __name__ == "__main__":
    main()
