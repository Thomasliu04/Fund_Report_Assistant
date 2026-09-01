"""不依赖 Excel GUI 时：按同一套规则把数据条/色阶画成 PNG（预览兜底）。"""

from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


HEADERS = [
    "排名", "排名变化", "基金公司", "总规模", "上半年增量",
    "上半年增速%", "货币", "非货", "货币排名", "非货排名",
]


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
    """排名色阶：小值红 → 中黄 → 大值绿。"""
    if vmax <= vmin:
        return (255, 235, 132)
    t = (value - vmin) / (vmax - vmin)
    if t < 0.5:
        return _lerp((248, 105, 107), (255, 235, 132), t * 2)
    return _lerp((255, 235, 132), (99, 190, 123), (t - 0.5) * 2)


def render_ranking_with_databars(
    rows: list[dict],
    output_path: str | Path,
    *,
    focus_company: str = "示例基金",
    scale: int = 2,
) -> Path:
    """绘制接近 Excel 数据条/色阶观感的宽表 PNG。"""
    font = _font(10 * scale)
    font_h = _font(10 * scale)
    col_w = [int(x * scale) for x in [40, 52, 110, 70, 90, 70, 70, 90, 70, 70]]
    row_h = 18 * scale
    header_h = 28 * scale
    pad = 2 * scale

    n = len(rows)
    width = sum(col_w) + pad * 2
    height = header_h + row_h * n + pad * 2
    img = Image.new("RGB", (width, height), (255, 255, 255))
    draw = ImageDraw.Draw(img)

    header_bg = (242, 220, 219)
    grid = (180, 198, 216)

    # header
    x = pad
    for c, h in enumerate(HEADERS):
        w = col_w[c]
        draw.rectangle([x, pad, x + w, pad + header_h], fill=header_bg, outline=grid)
        draw.text((x + 4 * scale, pad + 6 * scale), h, fill=(51, 51, 51), font=font_h)
        x += w

    incs = [r["increment"] for r in rows]
    nons = [r["non_money"] for r in rows]
    money_ranks = [r["money_rank"] for r in rows]
    nm_ranks = [r["non_money_rank"] for r in rows]
    max_inc = max(abs(v) for v in incs) or 1
    max_non = max(nons) or 1
    mr_min, mr_max = min(money_ranks), max(money_ranks)
    nr_min, nr_max = min(nm_ranks), max(nm_ranks)

    bar_blue = (99, 142, 198)
    bar_blue_end = (210, 225, 240)

    for ri, row in enumerate(rows):
        y = pad + header_h + ri * row_h
        is_focus = focus_company in row["company"]
        base_bg = (252, 228, 214) if is_focus else ((255, 255, 255) if ri % 2 == 0 else (245, 248, 252))
        values = [
            str(row["rank"]),
            str(row["rank_change"]),
            row["company"],
            str(round(row["aum"])),
            str(round(row["increment"])),
            str(round(row["growth_pct"] * 100)),
            str(round(row["money"])),
            str(round(row["non_money"])),
            str(row["money_rank"]),
            str(row["non_money_rank"]),
        ]
        x = pad
        for c, text in enumerate(values):
            w = col_w[c]
            bg = base_bg
            # 色阶列
            if c == 8:
                bg = _color_scale(row["money_rank"], mr_min, mr_max)
            elif c == 9:
                bg = _color_scale(row["non_money_rank"], nr_min, nr_max)

            draw.rectangle([x, y, x + w, y + row_h], fill=bg, outline=grid)

            # 数据条：增量（双向轴，负左正右）；非货（单向，自左向右）
            if c == 4:
                val = row["increment"]
                denom = max_inc
                frac = max(0.0, min(1.0, abs(val) / denom))
                half = (w - 8 * scale) // 2
                mid = x + w // 2
                bar_w = max(int(half * frac), 2 if abs(val) > 0 else 0)
                bar_top = y + 3 * scale
                bar_bot = y + row_h - 3 * scale
                if bar_w > 0:
                    if val >= 0:
                        draw.rectangle(
                            [mid, bar_top, mid + bar_w, bar_bot],
                            fill=(232, 80, 80),
                        )
                    else:
                        draw.rectangle(
                            [mid - bar_w, bar_top, mid, bar_bot],
                            fill=(80, 170, 90),
                        )
                # 中轴（零点）
                draw.line([mid, y + 1 * scale, mid, y + row_h - 1 * scale], fill=(80, 80, 80), width=max(1, scale // 2))
            elif c == 7:
                val = row["non_money"]
                denom = max_non
                frac = max(0.0, min(1.0, abs(val) / denom))
                bar_w = int((w - 8 * scale) * frac)
                if bar_w > 0:
                    draw.rectangle(
                        [x + 4 * scale, y + 3 * scale, x + 4 * scale + bar_w, y + row_h - 3 * scale],
                        fill=(232, 80, 80) if val >= 0 else (80, 170, 90),
                    )

            # 文字
            tw = draw.textlength(text, font=font)
            if c == 2:
                tx = x + 4 * scale
            else:
                tx = x + (w - tw) / 2
            fg = (156, 0, 6) if is_focus else (30, 30, 30)
            draw.text((tx, y + 3 * scale), text, fill=fg, font=font)
            x += w

    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    img.save(out, format="PNG")
    return out
