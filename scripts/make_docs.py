"""Regenerate examples/halcyon_lbo.xlsx and the PNGs in docs/ (needs LibreOffice and pdftoppm)."""

from __future__ import annotations

import shutil
import subprocess
import tempfile
from pathlib import Path

from PIL import Image, ImageChops

from lbo_lab import charts, excel, recalc, sensitivity
from lbo_lab.assumptions import Assumptions
from lbo_lab.engine import run_model

ROOT = Path(__file__).resolve().parents[1]


def summary_png(xlsx: Path, out: Path) -> None:
    """Render the workbook to PDF with LibreOffice and rasterize the Summary page."""
    with tempfile.TemporaryDirectory() as tmp:
        pdf = recalc.to_pdf(xlsx, Path(tmp))
        pages = int(subprocess.run(["pdfinfo", str(pdf)], capture_output=True, text=True, check=True)
                    .stdout.split("Pages:")[1].split()[0])
        for page in range(1, pages + 1):
            text = subprocess.run(["pdftotext", "-f", str(page), "-l", str(page), str(pdf), "-"],
                                  capture_output=True, text=True, check=True).stdout
            if "LBO returns summary" in text:
                break
        else:
            raise RuntimeError("Summary page not found in PDF")
        stem = Path(tmp) / "page"
        subprocess.run(["pdftoppm", "-png", "-r", "150", "-f", str(page), "-l", str(page), "-singlefile",
                        str(pdf), str(stem)], check=True)
        img = Image.open(f"{stem}.png").convert("RGB")
        bbox = ImageChops.difference(img, Image.new("RGB", img.size, (255, 255, 255))).getbbox()
        pad = 18
        box = (max(bbox[0] - pad, 0), max(bbox[1] - pad, 0), min(bbox[2] + pad, img.width),
               min(bbox[3] + pad, img.height))
        img.crop(box).save(out, optimize=True)


def main() -> None:
    a = Assumptions()
    (ROOT / "examples").mkdir(exist_ok=True)
    (ROOT / "docs").mkdir(exist_ok=True)
    xlsx = ROOT / "examples" / "halcyon_lbo.xlsx"
    excel.build_model_file(xlsx, a)
    res = run_model(a)
    charts.make_all(res, sensitivity.entry_vs_exit(a), ROOT / "docs")
    summary_png(xlsx, ROOT / "docs" / "returns_summary.png")
    shutil.rmtree(ROOT / "docs" / ".cache", ignore_errors=True)
    print("done")


if __name__ == "__main__":
    main()
