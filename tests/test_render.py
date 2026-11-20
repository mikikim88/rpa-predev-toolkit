"""Smoke tests: every renderer runs on the bundled examples and produces valid files.

Run:  python -m pytest tests -q      (or: python tests/test_render.py)
"""
from __future__ import annotations

import json
import sys
import tempfile
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1] / "plugins" / "rpa-predev"
sys.path.insert(0, str(ROOT / "scripts"))

import build_assessment  # noqa: E402
import read_intake  # noqa: E402
import render_pdd  # noqa: E402
import render_sdd  # noqa: E402
from openpyxl import Workbook, load_workbook  # noqa: E402
from pptx import Presentation  # noqa: E402
from scoring import grade_of, screen, total_score  # noqa: E402

EX = ROOT / "examples"


def load(name):
    return json.loads((EX / name).read_text(encoding="utf-8"))


def test_embedded_template_and_drm_passthrough(tmp_path: Path):
    import io

    from _template import template_bytes
    from drm import readable_copy
    assert zipfile.is_zipfile(io.BytesIO(template_bytes()))
    p = tmp_path / "plain.pptx"
    p.write_bytes(template_bytes())
    assert readable_copy(p) == p.resolve()  # normal files are used as-is


def test_scoring_rules():
    full = {k: {"score": 4} for k in
            ["standardization", "data_structure", "system_stability", "volume", "judgment"]}
    assert total_score(full) == 20 and grade_of(20) == "상"
    assert grade_of(19) == "중" and grade_of(11) == "하"
    assert total_score({"standardization": 5}) is None
    assert screen({"human_judgment": "필수 (사람 판단)"})[0] == "부적합 가능"
    assert screen({"monthly_volume": 100})[0] == "정보 부족"
    assert screen({"monthly_volume": 10, "minutes_per_case": 3})[0] == "검토 필요"
    assert screen({"monthly_volume": 300, "minutes_per_case": 5})[0] == "적합 후보"


def test_assessment_workbook(tmp_path: Path):
    out = build_assessment.build(load("assessment.json"), tmp_path / "a.xlsx")
    wb = load_workbook(out)
    assert wb.sheetnames == ["0_안내", "1_과제목록", "2_타당성평가", "3_ROI", "4_우선순위", "5_확인필요",
                             "6_현업작성양식"]
    assert wb["1_과제목록"]["N4"].value == "적합 후보"
    assert str(wb["2_타당성평가"]["H4"].value).startswith("=IF(COUNT(")


def test_intake_roundtrip(tmp_path: Path):
    wb = Workbook()
    build_assessment.sheet_intake(wb, standalone=True)
    ws = wb.active
    for col, v in enumerate(["경비 정산 검증", "재무팀", "담당", "설명", "월", 200, 3], start=1):
        ws.cell(row=6, column=col, value=v)
    p = tmp_path / "intake.xlsx"
    wb.save(p)
    rows = read_intake.read(p)
    assert len(rows) == 1 and rows[0]["task_name"] == "경비 정산 검증"
    assert rows[0]["monthly_volume"] == 200


def test_pdd_deck(tmp_path: Path):
    out = render_pdd.build(load("pdd.json"), tmp_path / "pdd.pptx")
    assert zipfile.is_zipfile(out)
    assert len(Presentation(str(out)).slides) >= 12


def test_pdd_handles_empty_spec(tmp_path: Path):
    out = render_pdd.build({"meta": {"task_name": "빈 업무"}}, tmp_path / "empty.pptx")
    text = " ".join(sh.text_frame.text for s in Presentation(str(out)).slides
                    for sh in s.shapes if sh.has_text_frame)
    assert "확인 필요" in text


def test_sdd_deck(tmp_path: Path):
    out = render_sdd.build(load("sdd.json"), tmp_path / "sdd.pptx")
    assert len(Presentation(str(out)).slides) >= 12
    text = " ".join(sh.text_frame.text for s in Presentation(str(out)).slides
                    for sh in s.shapes if sh.has_text_frame)
    assert "(Credential Store에 보관)" in " ".join(
        c.text for s in Presentation(str(out)).slides for sh in s.shapes if sh.has_table
        for row in sh.table.rows for c in row.cells)
    assert text


if __name__ == "__main__":
    fns = [v for k, v in dict(globals()).items() if k.startswith("test_")]
    for fn in fns:
        with tempfile.TemporaryDirectory() as d:
            if fn.__code__.co_argcount:
                fn(Path(d))
            else:
                fn()
        print(f"PASS {fn.__name__}")
