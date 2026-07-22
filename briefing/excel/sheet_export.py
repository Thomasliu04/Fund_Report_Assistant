"""把表图 Excel 各 sheet 导出为高清 PNG（经 Microsoft Excel 渲染条件格式）。

后台静默打开 Excel（不 activate），Copy Picture → 剪贴板 PDF（矢量）
→ pymupdf 高倍栅格化 → 白底 PNG。

说明：人工 Paste Special 进 PPT 得到的是 EMF 矢量；Mac 剪贴板无 EMF，
但有 PDF。将 PDF 按 2.5x 栅格化后清晰度接近手工粘贴，远优于直接取 PNGf。
失败返回 {}，由调用方回退到 PIL。
"""

from __future__ import annotations

import subprocess
import tempfile
from pathlib import Path

from PIL import Image

# Copy Picture 剪贴板 PDF 的逻辑尺寸约等于屏幕 PNG；2.5x ≈ 200dpi 量级
_PDF_RENDER_SCALE = 2.5


def _flatten_white(png_path: Path) -> Path:
    """Excel 复制出的图常把白底做成透明，贴 PPT 前铺白底。"""
    im = Image.open(png_path)
    if im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info):
        rgba = im.convert("RGBA")
        bg = Image.new("RGBA", rgba.size, (255, 255, 255, 255))
        bg.alpha_composite(rgba)
        out = bg.convert("RGB")
    else:
        out = im.convert("RGB")
    out.save(png_path, format="PNG")
    return png_path


def _pdf_to_png(pdf_path: Path, png_path: Path, *, scale: float = _PDF_RENDER_SCALE) -> bool:
    try:
        import fitz
    except ImportError:
        return False
    try:
        doc = fitz.open(pdf_path)
        if doc.page_count < 1:
            return False
        page = doc[0]
        pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
        pix.save(str(png_path))
        doc.close()
        return png_path.exists() and png_path.stat().st_size > 100
    except Exception:
        return False


def excel_available() -> bool:
    return Path("/Applications/Microsoft Excel.app").exists()


def export_sheets_via_excel(
    xlsx_path: str | Path,
    sheet_names: list[str],
    output_dir: str | Path,
    *,
    timeout_sec: int = 240,
    render_scale: float = _PDF_RENDER_SCALE,
) -> dict[str, Path]:
    """
    静默打开 xlsx，按 sheet 导出高清 PNG。
    返回 {sheet_name: png_path}；失败返回 {}。
    """
    xlsx_path = Path(xlsx_path).resolve()
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    if not excel_available() or not xlsx_path.exists() or not sheet_names:
        return {}

    names = [n for n in sheet_names if n and n != "使用说明"]
    if not names:
        return {}

    targets: list[tuple[str, Path]] = []
    for i, name in enumerate(names):
        targets.append((name, output_dir / f"sheet_{i:02d}.png"))

    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        runner = Path(tmp) / "export.applescript"
        # 每个 sheet：优先写 PDF，失败再写 PNGf
        lines = [
            'tell application "Microsoft Excel"',
            "  set wasVisible to visible",
            "  set visible to false",
            "  set display alerts to false",
            f'  open POSIX file "{xlsx_path}"',
            "  delay 1.5",
            "  set theWB to active workbook",
        ]
        pdf_paths: list[Path] = []
        for i, (name, dest) in enumerate(targets):
            esc = name.replace("\\", "\\\\").replace('"', '\\"')
            pdf_path = tmp_dir / f"sheet_{i:02d}.pdf"
            pdf_paths.append(pdf_path)
            dest_s = str(dest.resolve())
            pdf_s = str(pdf_path.resolve())
            lines += [
                "  try",
                f'    set sh to sheet "{esc}" of theWB',
                "    set pa to print area of sh",
                '    if pa is missing value or pa is "" then',
                "      set rng to used range of sh",
                "    else",
                "      set rng to range pa of sh",
                "    end if",
                "    copy picture rng appearance screen format picture",
                "    delay 0.35",
                "    set wrotePdf to false",
                "    try",
                "      set pdfData to the clipboard as «class PDF »",
                f'      set pdfRef to open for access POSIX file "{pdf_s}" with write permission',
                "      set eof of pdfRef to 0",
                "      write pdfData to pdfRef",
                "      close access pdfRef",
                "      set wrotePdf to true",
                "    end try",
                "    if wrotePdf is false then",
                "      set pngData to the clipboard as «class PNGf»",
                f'      set outPath to POSIX file "{dest_s}"',
                "      set fRef to open for access outPath with write permission",
                "      set eof of fRef to 0",
                "      write pngData to fRef",
                "      close access fRef",
                "    end if",
                "  end try",
            ]
        lines += [
            "  try",
            "    close theWB saving no",
            "  end try",
            "  try",
            "    if (count of workbooks) is 0 then",
            "      set visible to wasVisible",
            "      if wasVisible is false then quit",
            "    else",
            "      set visible to wasVisible",
            "    end if",
            "  end try",
            "end tell",
        ]
        runner.write_text("\n".join(lines), encoding="utf-8")
        try:
            r = subprocess.run(
                ["osascript", str(runner)],
                capture_output=True,
                text=True,
                timeout=timeout_sec,
            )
        except (OSError, subprocess.TimeoutExpired):
            _force_close_excel_workbook(xlsx_path)
            return {}
        if r.returncode != 0:
            _force_close_excel_workbook(xlsx_path)
            return {}

        # PDF → 高清 PNG
        for (name, dest), pdf_path in zip(targets, pdf_paths):
            if pdf_path.exists() and pdf_path.stat().st_size > 100:
                if not _pdf_to_png(pdf_path, dest, scale=render_scale):
                    # 栅格化失败则尝试已有 PNG（若 AppleScript 写过）
                    pass

    out: dict[str, Path] = {}
    for name, dest in targets:
        if dest.exists() and dest.stat().st_size > 100:
            try:
                _flatten_white(dest)
                out[name] = dest
            except OSError:
                continue
    # 一张都没导出成功 → 视为失败
    if len(out) < max(1, len(names) // 2):
        return {}
    return out


def _force_close_excel_workbook(xlsx_path: Path) -> None:
    """尽量关掉本次打开的工作簿，避免残留窗口。"""
    name = xlsx_path.name.replace('"', "")
    script = f'''
tell application "Microsoft Excel"
  try
    repeat with wb in workbooks
      try
        if name of wb contains "{name}" then close wb saving no
      end try
    end repeat
  end try
end tell
'''
    try:
        subprocess.run(["osascript", "-e", script], capture_output=True, timeout=30)
    except (OSError, subprocess.TimeoutExpired):
        pass
