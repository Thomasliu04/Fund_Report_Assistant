"""将表格数据渲染为接近简报风格的 PNG（用于替换 PPT 中的 EMF 贴图）。"""

from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


# 简报表常见配色（深蓝表头 + 浅底 + 银华行高亮）
HEADER_BG = (211, 40, 32)
HEADER_FG = (255, 255, 255)
ROW_BG_ALT = (232, 240, 248)
ROW_BG = (255, 255, 255)
GRID = (180, 198, 216)
TEXT = (40, 40, 40)
FOCUS_BG = (255, 243, 205)
FOCUS_FG = (153, 0, 0)


def _font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    candidates = [
        "/System/Library/Fonts/PingFang.ttc",
        "/System/Library/Fonts/STHeiti Light.ttc",
        "/Library/Fonts/Arial Unicode.ttf",
        "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
        "/System/Library/Fonts/Hiragino Sans GB.ttc",
    ]
    for path in candidates:
        if Path(path).exists():
            try:
                return ImageFont.truetype(path, size, index=0)
            except OSError:
                continue
    return ImageFont.load_default()


def render_table_image(
    headers: list[str],
    rows: list[list[str]],
    output_path: str | Path,
    *,
    focus_row_indices: list[int] | None = None,
    col_widths: list[int] | None = None,
    scale: int = 2,
) -> Path:
    """渲染表格为 PNG。scale=2 便于 PPT 中清晰显示。"""
    focus_row_indices = focus_row_indices or []
    font = _font(11 * scale)
    font_h = _font(11 * scale, bold=True)

    n_cols = len(headers)
    if col_widths is None:
        # 按内容估算
        col_widths = []
        for c in range(n_cols):
            samples = [headers[c]] + [r[c] for r in rows if c < len(r)]
            w = max(len(s) for s in samples) * 7 * scale + 16 * scale
            col_widths.append(max(w, 48 * scale))

    row_h = 22 * scale
    header_h = 26 * scale
    pad = 2 * scale
    width = sum(col_widths) + pad * 2
    height = header_h + row_h * len(rows) + pad * 2

    img = Image.new("RGB", (width, height), (255, 255, 255))
    draw = ImageDraw.Draw(img)

    # header
    x = pad
    y = pad
    for c, h in enumerate(headers):
        w = col_widths[c]
        draw.rectangle([x, y, x + w, y + header_h], fill=HEADER_BG, outline=GRID)
        draw.text((x + 6 * scale, y + 5 * scale), h, fill=HEADER_FG, font=font_h)
        x += w

    # body
    for ri, row in enumerate(rows):
        y = pad + header_h + ri * row_h
        x = pad
        is_focus = ri in focus_row_indices
        bg = FOCUS_BG if is_focus else (ROW_BG_ALT if ri % 2 else ROW_BG)
        fg = FOCUS_FG if is_focus else TEXT
        for c in range(n_cols):
            w = col_widths[c]
            cell = row[c] if c < len(row) else ""
            draw.rectangle([x, y, x + w, y + row_h], fill=bg, outline=GRID)
            # 首列居中，公司名左对齐，其余右对齐
            if c == 0:
                tw = draw.textlength(cell, font=font)
                tx = x + (w - tw) / 2
            elif c == 1 and any("\u4e00" <= ch <= "\u9fff" for ch in cell):
                tx = x + 6 * scale
            else:
                tw = draw.textlength(cell, font=font)
                tx = x + w - tw - 6 * scale
            draw.text((tx, y + 4 * scale), cell, fill=fg, font=font)
            x += w

    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    img.save(out, format="PNG")
    return out
