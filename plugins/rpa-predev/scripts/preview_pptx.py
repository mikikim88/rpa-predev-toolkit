"""Render a .pptx to PNG images for visual review (input analysis or output QA).

Usage:
    python preview_pptx.py deck.pptx [-o preview_dir]

Tries, in order:
  1. Microsoft PowerPoint (Windows, via PowerShell COM)
  2. LibreOffice (soffice) -> PDF, then pdftoppm -> PNG if available
Prints the produced files. Exit code 2 if no renderer is available.
"""
from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
from pathlib import Path


def via_powerpoint(src: Path, out: Path) -> bool:
    if os.name != "nt":
        return False
    ps = f"""
$ErrorActionPreference = 'Stop'
$pp = New-Object -ComObject PowerPoint.Application
try {{
  $p = $pp.Presentations.Open('{src}', $true, $false, $false)
  $i = 1
  foreach ($s in $p.Slides) {{ $s.Export((Join-Path '{out}' ('slide{{0:D2}}.png' -f $i)), 'PNG', 1600, 900); $i++ }}
  $p.Close()
}} finally {{ $pp.Quit() }}
"""
    try:
        r = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", ps],
                           capture_output=True, text=True, timeout=300)
    except (OSError, subprocess.TimeoutExpired):
        return False
    return r.returncode == 0 and any(out.glob("slide*.png"))


def via_libreoffice(src: Path, out: Path) -> bool:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        for cand in [r"C:\Program Files\LibreOffice\program\soffice.exe",
                     "/Applications/LibreOffice.app/Contents/MacOS/soffice"]:
            if Path(cand).exists():
                soffice = cand
    if not soffice:
        return False
    subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", str(out), str(src)],
                   capture_output=True, timeout=300)
    pdf = out / (src.stem + ".pdf")
    if not pdf.exists():
        return False
    if shutil.which("pdftoppm"):
        subprocess.run(["pdftoppm", "-png", "-r", "110", str(pdf), str(out / "slide")], timeout=300)
    return True


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pptx", type=Path)
    ap.add_argument("-o", "--out", type=Path)
    a = ap.parse_args()
    src = a.pptx.resolve()
    out = (a.out or src.with_name(src.stem + "_preview")).resolve()
    out.mkdir(parents=True, exist_ok=True)
    for f in out.glob("slide*.png"):
        f.unlink()
    if via_powerpoint(src, out) or via_libreoffice(src, out):
        files = sorted(p.name for p in out.iterdir())
        print(f"[OK] {out}\n" + "\n".join(files))
        return
    print("[WARN] PowerPoint 또는 LibreOffice가 없어 미리보기를 만들 수 없습니다. "
          "extract_pptx.py로 텍스트만 확인하세요.", file=sys.stderr)
    sys.exit(2)


if __name__ == "__main__":
    main()
