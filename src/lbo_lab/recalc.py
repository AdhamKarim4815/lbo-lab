"""Headless LibreOffice helpers: recalculate a workbook, or render it to PDF."""

from __future__ import annotations

import shutil
import subprocess
import tempfile
from pathlib import Path


def find_soffice() -> str | None:
    for name in ("soffice", "libreoffice"):
        path = shutil.which(name)
        if path:
            return path
    return None


_RECALC_XCU = """<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse">
<value>0</value></prop></item>
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="ODFRecalcMode" oor:op="fuse">
<value>0</value></prop></item>
</oor:items>
"""


def _convert(src: Path, fmt: str, outdir: Path, timeout: int) -> Path:
    soffice = find_soffice()
    if soffice is None:
        raise RuntimeError("LibreOffice (soffice) is not installed")
    outdir.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as profile:
        # A throwaway profile that forces "always recalculate" on load, so cells written without
        # cached values (as xlsxwriter does) are computed instead of showing zero.
        (Path(profile) / "user").mkdir()
        (Path(profile) / "user" / "registrymodifications.xcu").write_text(_RECALC_XCU, encoding="utf-8")
        cmd = [soffice, f"-env:UserInstallation=file://{profile}", "--headless",
               "--convert-to", fmt, "--outdir", str(outdir), str(src)]
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout, check=False)
    out = outdir / (src.stem + "." + fmt.split(":")[0])
    if proc.returncode != 0 or not out.exists():
        raise RuntimeError(f"LibreOffice conversion failed: {proc.stderr or proc.stdout}")
    return out


def recalculate(src: Path, outdir: Path, timeout: int = 180) -> Path:
    """Open ``src`` in headless LibreOffice, recalculate, and save a copy as xlsx in ``outdir``."""
    src, outdir = Path(src).resolve(), Path(outdir).resolve()
    if src.parent == outdir:
        raise ValueError("outdir must differ from the source directory")
    return _convert(src, "xlsx", outdir, timeout)


def to_pdf(src: Path, outdir: Path, timeout: int = 180) -> Path:
    return _convert(Path(src).resolve(), "pdf", Path(outdir).resolve(), timeout)
