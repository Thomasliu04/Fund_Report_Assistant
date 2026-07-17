"""将 Excel 工作表区域导出为 PNG（优先用本机 Microsoft Excel）。"""

from __future__ import annotations

import io
import subprocess
from pathlib import Path


def export_sheet_range_via_excel(
    xlsx_path: str | Path,
    output_png: str | Path,
    *,
    sheet_name: str = "总规模排名",
) -> Path:
    """
    用 Excel for Mac 打开工作簿，将该表导出为 PDF，再转 PNG。
    PDF 写在 output 同目录，避免 macOS 对 /var/folders 临时目录弹授权框。
    """
    xlsx_path = Path(xlsx_path).resolve()
    output_png = Path(output_png).resolve()
    output_png.parent.mkdir(parents=True, exist_ok=True)
    pdf_path = output_png.with_suffix(".pdf")

    script = f'''
tell application "Microsoft Excel"
    activate
    set wb to open workbook workbook file name POSIX file "{xlsx_path}"
    delay 1.5
    set pdfPOSIX to POSIX file "{pdf_path}"
    tell active workbook
        save workbook as filename pdfPOSIX file format PDF file format with overwrite
    end tell
    close active workbook saving no
end tell
'''
    result = subprocess.run(
        ["osascript", "-e", script],
        capture_output=True,
        text=True,
        timeout=120,
    )
    if result.returncode != 0 or not pdf_path.exists():
        raise RuntimeError(
            "Excel 导出 PDF 失败。\n"
            f"stdout: {result.stdout}\nstderr: {result.stderr}\n"
            "请在「系统设置 → 隐私与安全性 → 自动化」中允许终端控制 Excel；\n"
            "若弹出 Grant File Access，请选择项目下的 output 目录。"
        )

    try:
        import fitz  # pymupdf
    except ImportError:
        subprocess.check_call(["python3", "-m", "pip", "install", "-q", "pymupdf"])
        import fitz

    from PIL import Image

    doc = fitz.open(pdf_path)
    page = doc[0]
    pix = page.get_pixmap(matrix=fitz.Matrix(2.5, 2.5), alpha=False)
    im = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
    w, h = im.size
    pixels = im.load()

    def is_white(x, y):
        p = pixels[x, y]
        return p[0] > 250 and p[1] > 250 and p[2] > 250

    left, top, right, bottom = 0, 0, w - 1, h - 1
    while left < w and all(is_white(left, y) for y in range(0, h, 4)):
        left += 1
    while right > left and all(is_white(right, y) for y in range(0, h, 4)):
        right -= 1
    while top < h and all(is_white(x, top) for x in range(0, w, 4)):
        top += 1
    while bottom > top and all(is_white(x, bottom) for x in range(0, w, 4)):
        bottom -= 1
    pad = 8
    im = im.crop((max(0, left - pad), max(0, top - pad), min(w, right + pad), min(h, bottom + pad)))
    im.save(output_png, format="PNG")
    doc.close()
    return output_png
