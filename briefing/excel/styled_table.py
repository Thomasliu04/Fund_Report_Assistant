"""通用带数据条/色阶的表格 PNG 渲染（对齐简报 Excel 条件格式观感）。"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Literal

from PIL import Image, ImageDraw, ImageFont

BarMode = Literal["none", "bidirectional", "positive"]
ScaleMode = Literal["none", "rank"]  # rank: 小值红大值绿


@dataclass
class ColumnSpec:
    key: str
    header: str
    width: int = 70  # 逻辑宽度（再乘 scale）
    align: Literal["left", "center", "right"] = "center"
    fmt: Literal["text", "int", "pct"] = "text"
    bar: BarMode = "none"
    color_scale: ScaleMode = "none"


def _font(size: int):
    for path in (
        "/System/Library/Fonts/PingFang.ttc",
        "/System/Library/Fonts/STHeiti Light.ttc",
        "/Library/Fonts/Arial Unicode.ttf",
    ):
        if Path(path).exists():
            try:
                return ImageFont.truetype(path, size, index=0)
            except OSError:
                continue
    return ImageFont.load_default()


def _lerp(a, b, t):
    return tuple(int(a[i] + (b[i] - a[i]) * t) for i in range(3))


def _color_scale(value: float, vmin: float, vmax: float) -> tuple[int, int, int]:
    if vmax <= vmin:
        return (255, 235, 132)
    t = max(0.0, min(1.0, (value - vmin) / (vmax - vmin)))
    if t < 0.5:
        return _lerp((248, 105, 107), (255, 235, 132), t * 2)
    return _lerp((255, 235, 132), (99, 190, 123), (t - 0.5) * 2)


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
            # 已是百分比数值（如 9）或小数（0.09）
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
    focus_company: str = "银华基金",
    focus_key: str = "company",
    scale: int = 2,
) -> Path:
    """按列规格渲染表格。"""
    font = _font(9 * scale)
    font_h = _font(9 * scale)
    col_w = [c.width * scale for c in columns]
    row_h = 17 * scale
    header_h = 26 * scale
    pad = 2 * scale

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
        draw.text((x + 3 * scale, pad + 5 * scale), col.header, fill=(51, 51, 51), font=font_h)
        x += w

    # precompute scales / bar dens
    # bidirectional: (vmin<=0, vmax>=0)；positive: abs max
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
        if col.color_scale == "rank" and vals:
            scale_range[col.key] = (min(vals), max(vals))

    for ri, row in enumerate(rows):
        y = pad + header_h + ri * row_h
        company = str(row.get(focus_key, ""))
        is_focus = focus_company in company
        base_bg = (252, 228, 214) if is_focus else ((255, 255, 255) if ri % 2 == 0 else (245, 248, 252))
        x = pad
        for c, col in enumerate(columns):
            w = col_w[c]
            raw = row.get(col.key, "")
            text = _format_cell(raw, col.fmt)
            bg = base_bg
            if col.color_scale == "rank" and col.key in scale_range:
                try:
                    bg = _color_scale(float(raw), *scale_range[col.key])
                except (TypeError, ValueError):
                    pass
            draw.rectangle([x, y, x + w, y + row_h], fill=bg, outline=grid)

            # data bars
            bar_top, bar_bot = y + 3 * scale, y + row_h - 3 * scale
            if col.bar == "bidirectional" and col.key in bar_span:
                try:
                    val = float(raw)
                except (TypeError, ValueError):
                    val = 0.0
                vmin, vmax = bar_span[col.key]
                span = vmax - vmin
                # 零轴按 min~max 比例偏移：正数跨度大 → 轴靠左；负数跨度大 → 轴靠右
                inset = 3 * scale
                usable = w - 2 * inset
                zero_frac = (0.0 - vmin) / span
                mid = x + inset + usable * zero_frac
                left_room = usable * zero_frac
                right_room = usable * (1.0 - zero_frac)
                if val >= 0 and vmax > 0:
                    bar_w = max(int(right_room * (val / vmax)), 2 if abs(val) > 1e-9 else 0)
                    if bar_w > 0:
                        draw.rectangle([mid, bar_top, mid + bar_w, bar_bot], fill=(232, 80, 80))
                elif val < 0 and vmin < 0:
                    bar_w = max(int(left_room * (abs(val) / abs(vmin))), 2)
                    if bar_w > 0:
                        draw.rectangle([mid - bar_w, bar_top, mid, bar_bot], fill=(80, 170, 90))
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
                    color = (232, 80, 80) if val >= 0 else (80, 170, 90)
                    draw.rectangle(
                        [x + 3 * scale, bar_top, x + 3 * scale + bar_w, bar_bot],
                        fill=color,
                    )

            tw = draw.textlength(text, font=font)
            if col.align == "left":
                tx = x + 3 * scale
            elif col.align == "right":
                tx = x + w - tw - 3 * scale
            else:
                tx = x + (w - tw) / 2
            fg = (156, 0, 6) if is_focus else (30, 30, 30)
            # 有数据条时加浅色描边，避免压在条子上难读
            if col.bar != "none":
                for dx, dy in ((-1, 0), (1, 0), (0, -1), (0, 1)):
                    draw.text((tx + dx, y + 3 * scale + dy), text, fill=(255, 255, 255), font=font)
            draw.text((tx, y + 3 * scale), text, fill=fg, font=font)
            x += w

    if not rows:
        draw.text((pad + 10, pad + header_h + 10), "暂无数据", fill=(120, 120, 120), font=font)

    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    img.save(out, format="PNG")
    return out
