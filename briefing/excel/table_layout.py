"""表格列宽 / 字号自适应，提升 PPT 贴图可读性。"""

from __future__ import annotations

from typing import Any, Sequence


def font_pt_for_columns(n_cols: int, *, base: int = 11) -> int:
    """列越多字号略降，但 Excel/PPT 保底可读。"""
    if n_cols <= 14:
        return base
    if n_cols <= 18:
        return max(10, base - 1)
    return max(9, base - 2)


def target_content_width(n_cols: int, *, scale: int = 2) -> int:
    """
    目标表内容宽度（像素，已含 scale）。
    宽表约对齐 16:9 幻灯片主图区；列少时不会过宽。
    """
    per = 118 if n_cols <= 11 else (100 if n_cols <= 15 else (88 if n_cols <= 19 else 78))
    logical = n_cols * per
    lo, hi = 520 * scale, 1180 * scale
    return int(max(lo, min(hi, logical * scale)))


def format_header_wrap(header: str) -> str:
    """
    案例风格表头换行：
    - 主动权益排名 → 主动权益\\n排名
    - 25年排名变化 → 25年\\n排名变化
    已含换行则原样返回。
    """
    s = str(header).strip() if header is not None else ""
    if not s or "\n" in s:
        return s
    if s.endswith("排名变化") and len(s) > 4:
        return f"{s[:-4]}\n排名变化"
    if s.endswith("排名") and s != "排名":
        return f"{s[:-2]}\n排名"
    return s


def _header_lines(header: str) -> list[str]:
    wrapped = format_header_wrap(header)
    lines = [ln for ln in wrapped.split("\n") if ln != ""]
    return lines or [""]


def _text_px(text: str, font) -> float:
    try:
        if hasattr(font, "getlength"):
            return float(font.getlength(text))
        bbox = font.getbbox(text)
        return float(bbox[2] - bbox[0])
    except Exception:
        return float(max(len(text), 1) * 7)


def _cell_text(val: Any, fmt: str) -> str:
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


def measure_column_widths(
    columns: Sequence[Any],
    rows: list[dict],
    font,
    font_h,
    *,
    scale: int = 2,
) -> list[int]:
    """按表头 + 单元格内容测量列宽（像素）。"""
    pad = 8 * scale
    widths: list[int] = []
    for col in columns:
        lines = _header_lines(col.header)
        flat = "".join(lines)
        need = max(_text_px(ln, font_h) for ln in lines) + pad

        sample_n = 0
        for row in rows:
            text = _cell_text(row.get(col.key, ""), getattr(col, "fmt", "text"))
            if text:
                need = max(need, _text_px(text, font) + pad)
                sample_n += 1
                if sample_n >= 40:
                    break

        if getattr(col, "bar", "none") != "none":
            need = max(need, 56 * scale)
        if getattr(col, "key", "") in {"company", "label"} or "公司" in flat or flat in {
            "类型",
            "管理人",
        }:
            need = max(need, 72 * scale)

        min_w = 36 * scale
        max_w = 150 * scale
        # 仅「排名」或短前缀（货币/非货）可收窄；主动权益等须容纳首行全宽
        if flat == "排名":
            max_w = 52 * scale
            min_w = 32 * scale
        elif flat.endswith("排名") and "变化" not in flat and "增量" not in flat:
            first = lines[0] if lines else flat
            if len(first) <= 2:
                max_w = 60 * scale
                min_w = 36 * scale
            else:
                min_w = max(min_w, int(_text_px(first, font_h) + pad))
                max_w = 110 * scale
        widths.append(int(max(min_w, min(max_w, need))))
    return widths


def normalize_widths(
    widths: Sequence[int],
    target: int,
    *,
    scale: int = 2,
    min_w: int | None = None,
) -> list[int]:
    """把列宽总和收到目标宽度附近，保持相对比例。"""
    min_w = (34 * scale) if min_w is None else min_w
    w = [max(min_w, int(x)) for x in widths]
    total = sum(w) or 1
    if abs(total - target) / target < 0.04:
        return w

    factor = target / total
    scaled = [max(min_w, int(round(x * factor))) for x in w]
    drift = target - sum(scaled)
    i = 0
    n = len(scaled)
    while drift != 0 and n:
        j = i % n
        if drift > 0:
            scaled[j] += 1
            drift -= 1
        elif scaled[j] > min_w:
            scaled[j] -= 1
            drift += 1
        i += 1
        if i > n * 8:
            break
    return scaled


def wrap_header(header: str, col_w: int, font, *, scale: int = 2) -> list[str]:
    """表头换行：优先尊重案例式 \\n；否则按列宽拆成最多两行。"""
    lines = _header_lines(header)
    if len(lines) >= 2:
        return lines[:2]
    h = lines[0] if lines else ""
    if not h:
        return [""]
    if _text_px(h, font) + 6 * scale <= col_w:
        return [h]
    # 优先在「排名」前拆
    if h.endswith("排名") and len(h) > 2:
        return [h[:-2], "排名"]
    if h.endswith("排名变化") and len(h) > 4:
        return [h[:-4], "排名变化"]
    best = max(1, len(h) // 2)
    for cut in range(2, len(h) - 1):
        a, b = h[:cut], h[cut:]
        if max(_text_px(a, font), _text_px(b, font)) + 6 * scale <= col_w:
            best = cut
            break
    return [h[:best], h[best:]]


def _disp_len(text: str) -> float:
    """显示宽度：中文约 2，ASCII 约 1（Excel 列宽近似）。"""
    n = 0.0
    for ch in str(text):
        n += 2.0 if ord(ch) > 127 else 1.0
    return n


def excel_char_width(col: Any, rows: list[dict], *, n_cols: int) -> float:
    """
    估算 openpyxl 列宽（字符单位），兼顾表头/数字可读性。
    多行表头按「最长一行」计宽，避免主动权益\\n排名被压成单行裁切。
    """
    lines = _header_lines(getattr(col, "header", ""))
    flat = "".join(lines)
    fmt = getattr(col, "fmt", "text")
    key = getattr(col, "key", "")

    w = max(_disp_len(ln) for ln in lines) + 2.4

    sample = 0
    for row in rows:
        text = _cell_text(row.get(key, ""), fmt)
        if text:
            w = max(w, _disp_len(text) + 2)
            sample += 1
            if sample >= 35:
                break

    if getattr(col, "bar", "none") != "none":
        w = max(w, 9)
    if key in {"company", "label"} or "公司" in flat or flat in {"类型", "管理人"}:
        w = max(w, 8)

    if flat == "排名":
        w = min(max(w, 4.5), 6.0)
    elif flat.endswith("排名") and "变化" not in flat:
        first = lines[0] if lines else flat
        # 主动权益 / 被动权益 等：首行必须完整可见（对齐案例约 9–13）
        if len(first) >= 3:
            w = max(w, _disp_len(first) + 2.6, 9.0)
        else:
            w = max(w, 6.5)

    max_w = 16.0 if n_cols <= 12 else (14.0 if n_cols <= 16 else (12.5 if n_cols <= 19 else 11.0))
    return round(max(4.2, min(max_w, w)), 1)


def fit_excel_widths(
    widths: list[float],
    *,
    n_cols: int,
    budget: float | None = None,
    min_widths: list[float] | None = None,
) -> list[float]:
    """若总宽超出预算，按比例压缩；min_widths 用于保护换行表头列。"""
    if budget is None:
        # 列多时放宽预算，避免把「主动权益」列压到裁切
        budget = 110.0 if n_cols <= 12 else (130.0 if n_cols <= 16 else 150.0)
    floors = min_widths or [4.5] * len(widths)
    total = sum(widths) or 1.0
    if total <= budget:
        return [round(max(floors[i], w), 1) for i, w in enumerate(widths)]
    factor = budget / total
    return [round(max(floors[i], w * factor), 1) for i, w in enumerate(widths)]
