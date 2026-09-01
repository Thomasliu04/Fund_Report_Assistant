"""通用带数据条/色阶的表格 PNG 渲染（对齐简报 Excel 条件格式观感）。"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Literal

from PIL import Image, ImageDraw, ImageFont

from briefing.excel.table_layout import (
    font_pt_for_columns,
    measure_column_widths,
    normalize_widths,
    target_content_width,
    wrap_header,
)

BarMode = Literal["none", "bidirectional", "positive"]
ScaleMode = Literal["none", "rank", "growth"]  # rank: 小红大绿；growth: 小绿大红
BarColor = Literal["red", "blue"]


@dataclass
class ColumnSpec:
    key: str
    header: str
    width: int = 70  # 逻辑宽度（再乘 scale）；autofit 时会被覆盖
    align: Literal["left", "center", "right"] = "center"
    fmt: Literal["text", "int", "pct"] = "text"
    bar: BarMode = "none"
    color_scale: ScaleMode = "none"
    bar_color: BarColor = "red"


def _font(size: int, *, bold: bool = False):
    """加载中文字体；bold 时优先 SemiBold/Medium/Bold 字重。"""
    candidates: list[tuple[str, int]] = []
    if bold:
        candidates.extend(
            [
                ("/System/Library/Fonts/STHeiti Medium.ttc", 0),
                ("/System/Library/Fonts/Supplemental/Songti.ttc", 1),
                ("/Library/Fonts/Arial Bold.ttf", 0),
                ("/System/Library/Fonts/PingFang.ttc", 8),
                ("/System/Library/Fonts/PingFang.ttc", 7),
            ]
        )
    candidates.extend(
        [
            ("/System/Library/Fonts/PingFang.ttc", 0),
            ("/System/Library/Fonts/STHeiti Light.ttc", 0),
            ("/Library/Fonts/Arial Unicode.ttf", 0),
        ]
    )
    for path, index in candidates:
        if Path(path).exists():
            try:
                return ImageFont.truetype(path, size, index=index)
            except OSError:
                continue
    return ImageFont.load_default()


def _lerp(a, b, t):
    return tuple(int(a[i] + (b[i] - a[i]) * t) for i in range(3))


def _color_scale(value: float, vmin: float, vmax: float) -> tuple[int, int, int]:
    """排名色阶：小值红 → 大值绿。"""
    if vmax <= vmin:
        return (255, 235, 132)
    t = max(0.0, min(1.0, (value - vmin) / (vmax - vmin)))
    if t < 0.5:
        return _lerp((248, 105, 107), (255, 235, 132), t * 2)
    return _lerp((255, 235, 132), (99, 190, 123), (t - 0.5) * 2)


def _growth_color_scale(value: float, vmin: float, vmax: float) -> tuple[int, int, int]:
    """增速色阶：小值绿 → 大值红（越高越红）。"""
    if vmax <= vmin:
        return (255, 235, 132)
    t = max(0.0, min(1.0, (value - vmin) / (vmax - vmin)))
    if t < 0.5:
        return _lerp((99, 190, 123), (255, 235, 132), t * 2)
    return _lerp((255, 235, 132), (248, 105, 107), (t - 0.5) * 2)


def _bar_fill(color: str, frac: float, positive: bool = True) -> tuple[int, int, int]:
    """渐变感：正红/蓝由浅到深；负绿。"""
    frac = max(0.15, min(1.0, frac))
    if not positive:
        base = (80, 170, 90)
    elif color == "blue":
        base = (99, 142, 198)  # #638EC6
    else:
        base = (255, 85, 90)  # #FF555A
    return _lerp((255, 255, 255), base, 0.35 + 0.65 * frac)


def _format_cell(val: Any, fmt: str) -> str:
    if val is None:
        return ""
    if fmt == "int":
        try:
            return str(int(round(float(val))))
        except (TypeError, ValueError):
            return str(val)
    if fmt == "pct":
        try:
            v = float(val)
            if abs(v) <= 2:
                v = v * 100
            return f"{int(round(v))}%"
        except (TypeError, ValueError):
            return str(val)
    return str(val)


def render_styled_table(
    columns: list[ColumnSpec],
    rows: list[dict],
    output_path: str | Path,
    *,
    focus_company: str = "示例基金",
    focus_key: str = "company",
    scale: int = 2,
    autofit: bool = True,
    target_width: int | None = None,
) -> Path:
    """
    按列规格渲染表格。

    autofit: 按内容测量列宽并归一到 target_width（提升 PPT 可读性）
    target_width: 内容区目标像素宽（不含外边距）；默认按列数估算
    """
    n_cols = len(columns)
    font_pt = font_pt_for_columns(n_cols)
    font = _font(font_pt * scale)
    font_h = _font(font_pt * scale, bold=True)
    pad = 2 * scale

    if autofit and columns:
        measured = measure_column_widths(columns, rows, font, font_h, scale=scale)
        tw = target_width if target_width is not None else target_content_width(n_cols, scale=scale)
        col_w = normalize_widths(measured, tw, scale=scale)
    else:
        col_w = [c.width * scale for c in columns]

    # 双行表头时加高
    header_lines = [wrap_header(c.header, col_w[i], font_h, scale=scale) for i, c in enumerate(columns)]
    max_hdr_lines = max((len(h) for h in header_lines), default=1)
    row_h = max(16, int(round(17 + (font_pt - 11) * 0.8))) * scale
    header_h = (26 if max_hdr_lines == 1 else 34) * scale

    n = len(rows)
    width = sum(col_w) + pad * 2
    height = header_h + row_h * max(n, 1) + pad * 2
    img = Image.new("RGB", (width, height), (255, 255, 255))
    draw = ImageDraw.Draw(img)
    header_bg = (242, 220, 219)
    grid = (180, 198, 216)

    # header
    x = pad
    for c, col in enumerate(columns):
        w = col_w[c]
        draw.rectangle([x, pad, x + w, pad + header_h], fill=header_bg, outline=grid)
        lines = header_lines[c]
        line_gap = 12 * scale if len(lines) > 1 else 0
        total_text_h = len(lines) * (font_pt * scale) + max(0, len(lines) - 1) * 2
        ty = pad + max(2 * scale, (header_h - total_text_h) // 2)
        for li, line in enumerate(lines):
            tw = draw.textlength(line, font=font_h)
            tx = x + max(2 * scale, (w - tw) / 2)
            draw.text((tx, ty + li * (font_pt * scale + 2)), line, fill=(51, 51, 51), font=font_h)
        x += w

    # precompute scales / bar dens
    bar_span: dict[str, tuple[float, float]] = {}
    bar_max: dict[str, float] = {}
    scale_range: dict[str, tuple[float, float]] = {}
    for col in columns:
        vals = []
        for r in rows:
            try:
                vals.append(float(r.get(col.key, 0) or 0))
            except (TypeError, ValueError):
                pass
        if col.bar == "bidirectional" and vals:
            vmin = min(min(vals), 0.0)
            vmax = max(max(vals), 0.0)
            if vmax - vmin < 1e-12:
                vmax = 1.0
            bar_span[col.key] = (vmin, vmax)
        elif col.bar == "positive" and vals:
            bar_max[col.key] = max(abs(v) for v in vals) or 1.0
        if col.color_scale in {"rank", "growth"} and vals:
            scale_range[col.key] = (min(vals), max(vals))

    focus_token = (focus_company or "").replace("基金", "").strip()

    for ri, row in enumerate(rows):
        y = pad + header_h + ri * row_h
        company = str(row.get(focus_key, ""))
        is_focus = bool(focus_token) and focus_token != "__none__" and (
            focus_token in company.replace("基金", "") or focus_company in company
        )
        base_bg = (252, 228, 214) if is_focus else ((255, 255, 255) if ri % 2 == 0 else (245, 248, 252))
        x = pad
        for c, col in enumerate(columns):
            w = col_w[c]
            raw = row.get(col.key, "")
            text = _format_cell(raw, col.fmt)
            bg = base_bg
            if col.color_scale in scale_range and col.key in scale_range:
                try:
                    fn = _growth_color_scale if col.color_scale == "growth" else _color_scale
                    bg = fn(float(raw), *scale_range[col.key])
                except (TypeError, ValueError):
                    pass
            draw.rectangle([x, y, x + w, y + row_h], fill=bg, outline=grid)

            bar_top, bar_bot = y + 3 * scale, y + row_h - 3 * scale
            if col.bar == "bidirectional" and col.key in bar_span:
                try:
                    val = float(raw)
                except (TypeError, ValueError):
                    val = 0.0
                vmin, vmax = bar_span[col.key]
                span = vmax - vmin
                inset = 3 * scale
                usable = w - 2 * inset
                zero_frac = (0.0 - vmin) / span
                mid = x + inset + usable * zero_frac
                left_room = usable * zero_frac
                right_room = usable * (1.0 - zero_frac)
                if val >= 0 and vmax > 0:
                    frac = val / vmax
                    bar_w = max(int(right_room * frac), 2 if abs(val) > 1e-9 else 0)
                    if bar_w > 0:
                        draw.rectangle(
                            [mid, bar_top, mid + bar_w, bar_bot],
                            fill=_bar_fill(col.bar_color, frac, True),
                        )
                elif val < 0 and vmin < 0:
                    frac = abs(val) / abs(vmin)
                    bar_w = max(int(left_room * frac), 2)
                    if bar_w > 0:
                        draw.rectangle(
                            [mid - bar_w, bar_top, mid, bar_bot],
                            fill=_bar_fill(col.bar_color, frac, False),
                        )
                draw.line(
                    [mid, y + scale, mid, y + row_h - scale],
                    fill=(80, 80, 80),
                    width=max(1, scale // 2),
                )
            elif col.bar == "positive" and col.key in bar_max:
                try:
                    val = float(raw)
                except (TypeError, ValueError):
                    val = 0.0
                denom = bar_max[col.key]
                frac = max(0.0, min(1.0, abs(val) / denom))
                bar_w = int((w - 6 * scale) * frac)
                if bar_w > 0:
                    draw.rectangle(
                        [x + 3 * scale, bar_top, x + 3 * scale + bar_w, bar_bot],
                        fill=_bar_fill(col.bar_color, frac, val >= 0),
                    )

            tw = draw.textlength(text, font=font)
            # 防溢出：过长则略缩进仍尽量完整显示
            if col.align == "left":
                tx = x + 3 * scale
            elif col.align == "right":
                tx = x + w - tw - 3 * scale
            else:
                tx = x + (w - tw) / 2
            if tw > w - 4 * scale:
                tx = x + 2 * scale
            fg = (120, 0, 0) if is_focus else (30, 30, 30)
            text_y = y + max(2 * scale, (row_h - font_pt * scale) // 2)
            if col.bar != "none":
                for dx, dy in ((-1, 0), (1, 0), (0, -1), (0, 1)):
                    draw.text((tx + dx, text_y + dy), text, fill=(255, 255, 255), font=font)
            if is_focus:
                for dx, dy in ((0, 0), (1, 0), (0, 1), (1, 1)):
                    draw.text((tx + dx, text_y + dy), text, fill=fg, font=font)
            else:
                draw.text((tx, text_y), text, fill=fg, font=font)
            x += w

    if not rows:
        draw.text((pad + 10, pad + header_h + 10), "暂无数据", fill=(120, 120, 120), font=font)

    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    img.save(out, format="PNG")
    return out
