"""Make DRM-encrypted Office inputs readable.

On corporate PCs a document-security agent may encrypt .pptx/.xlsx files on disk.
python-pptx / openpyxl then fail ("Package not found", "File is not a zip file"),
while PowerPoint / Excel still open them. readable_copy() returns the original
path when it is a normal Office (ZIP) file; otherwise, on Windows, it asks
PowerPoint or Excel to save a copy into the temp folder and returns that path.
"""
from __future__ import annotations

import os
import subprocess
import tempfile
import zipfile
from pathlib import Path

APPS = {
    ".pptx": ("PowerPoint.Application", "$d = $app.Presentations.Open($src, $true, $false, $false)"),
    ".pptm": ("PowerPoint.Application", "$d = $app.Presentations.Open($src, $true, $false, $false)"),
    ".xlsx": ("Excel.Application", "$app.DisplayAlerts = $false; $d = $app.Workbooks.Open($src, 0, $true)"),
    ".xlsm": ("Excel.Application", "$app.DisplayAlerts = $false; $d = $app.Workbooks.Open($src, 0, $true)"),
}


class UnreadableOfficeFile(RuntimeError):
    pass


def _ps_quote(s: str) -> str:
    return "'" + s.replace("'", "''") + "'"


def readable_copy(path: str | os.PathLike) -> Path:
    src = Path(path).resolve()
    if zipfile.is_zipfile(src):
        return src
    ext = src.suffix.lower()
    if os.name != "nt" or ext not in APPS:
        raise UnreadableOfficeFile(
            f"{src.name}: Office(ZIP) 형식이 아닙니다. 문서보안(DRM)으로 암호화된 파일이면 "
            "Office에서 열어 '다른 이름으로 저장'한 복사본을 사용하세요.")
    out_dir = Path(tempfile.mkdtemp(prefix="rpa_predev_"))
    dst = out_dir / src.name
    progid, open_cmd = APPS[ext]
    ps = (f"$ErrorActionPreference='Stop'; $src={_ps_quote(str(src))}; $dst={_ps_quote(str(dst))}; "
          f"$app = New-Object -ComObject {progid}; "
          f"try {{ {open_cmd}; $d.SaveCopyAs($dst); $d.Close($false) }} finally {{ $app.Quit() }}")
    if progid.startswith("PowerPoint"):
        ps = ps.replace("$d.Close($false)", "$d.Close()")
    try:
        r = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", ps],
                           capture_output=True, text=True, timeout=300)
    except (OSError, subprocess.TimeoutExpired) as e:
        raise UnreadableOfficeFile(f"{src.name}: Office 자동화 실행 실패 ({e})") from e
    if r.returncode != 0 or not zipfile.is_zipfile(dst):
        raise UnreadableOfficeFile(
            f"{src.name}: 암호화된 파일을 Office로 변환하지 못했습니다. {r.stderr.strip()[:300]}")
    return dst
