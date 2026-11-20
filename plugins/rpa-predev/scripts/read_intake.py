"""Read filled 현업 조사 양식 (6_현업작성양식 sheet) into assessment.json candidates.

Usage:
    python read_intake.py 조사서1.xlsx [조사서2.xlsx ...] -o assessment.json [--append]

--append : keep existing candidates in -o file and add new ones (IDs continue).
Scores / 자동화율 are left empty — they are filled in the feasibility step.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_assessment import INTAKE_COLS  # noqa: E402

FIELD_MAP = {
    "업무명": "task_name", "부서": "department", "담당자": "owner", "업무 설명": "description",
    "수행 빈도": "frequency", "월 처리건수": "monthly_volume", "건당 처리시간(분)": "minutes_per_case",
    "사용 시스템": "systems", "입력 데이터 형태": "input_type", "판단 개입": "human_judgment",
    "예외 상황": "exceptions_note", "비고": "notes",
}


def find_sheet(wb):
    for ws in wb.worksheets:
        for r in range(1, 6):
            if ws.cell(row=r, column=1).value == "업무명":
                return ws, r
    raise SystemExit("업무명 헤더가 있는 시트를 찾지 못했습니다 (6_현업작성양식).")


def read(path: Path) -> list[dict]:
    wb = load_workbook(path, data_only=True)
    ws, hr = find_sheet(wb)
    heads = [ws.cell(row=hr, column=c).value for c in range(1, ws.max_column + 1)]
    out = []
    for r in range(hr + 1, ws.max_row + 1):
        row = {h: ws.cell(row=r, column=i + 1).value for i, h in enumerate(heads) if h in FIELD_MAP}
        name = row.get("업무명")
        if not name or str(name).startswith("(예시)"):
            continue
        if name == INTAKE_COLS[0][1]:
            continue  # the template's field-description row
        cand = {FIELD_MAP[k]: v for k, v in row.items() if v not in (None, "")}
        cand["source"] = f"{path.name} {r}행"
        out.append(cand)
    return out


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("files", nargs="+", type=Path)
    ap.add_argument("-o", "--out", type=Path, default=Path("assessment.json"))
    ap.add_argument("--append", action="store_true")
    a = ap.parse_args()
    data = {"meta": {}, "candidates": []}
    if a.append and a.out.exists():
        data = json.loads(a.out.read_text(encoding="utf-8-sig"))
    existing = {c.get("task_name") for c in data["candidates"]}
    n = len(data["candidates"])
    added = 0
    for p in a.files:
        for c in read(p):
            if c.get("task_name") in existing:
                print(f"[SKIP] 중복 업무명: {c['task_name']}")
                continue
            n += 1
            added += 1
            data["candidates"].append({"id": f"T-{n:03d}", **c})
    a.out.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"[OK] {a.out}  (+{added}, total {len(data['candidates'])})")


if __name__ == "__main__":
    main()
