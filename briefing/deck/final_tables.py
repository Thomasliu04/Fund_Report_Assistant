"""以底稿「终表」为基准：按固定 10 页结构读入并导出带条件格式的表图工作簿。

终表 sheet 名对齐 25Q4 底稿；H1 旧名作为别名兼容。
第 6 页为左右双表（非货ETF / 权益ETF），导出为两个 sheet。
"""

from __future__ import annotations

import re
import warnings
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Literal

import pandas as pd

from briefing.excel.deck_workbook import build_deck_workbook
from briefing.excel.styled_table import ColumnSpec

SheetKind = Literal[
    "category",
    "total",
    "non_money",
    "increment",
    "business",
    "etf",
]


@dataclass(frozen=True)
class FinalTableSpec:
    """一张导出表对应底稿中的一块矩形区域。"""

    out_name: str
    sources: tuple[str, ...]
    kind: SheetKind
    # 1-based inclusive; None = 从第 1 列读到表头连续非空为止
    col_start: int = 1
    col_end: int | None = None
    # 行业表：读到首个空行即停（避免 top30 / 示例分块）
    stop_at_blank_row: bool = False
    # 行业表只要合计~FOF，不要后面的杂项行
    max_data_rows: int | None = None


# 行业表展示顺序（对齐案例；跳过 REITs/另类）
_CATEGORY_ORDER = (
    "合计",
    "非货",
    "货币",
    "固收",
    "被动权益",
    "主动权益",
    "固收+",
    "FOF",
)

# 与 25Q4 简报数据.pptx / 底稿终表对齐的 10 页（第 6 页拆成两张）
FINAL_TABLE_SPECS: tuple[FinalTableSpec, ...] = (
    FinalTableSpec(
        "01_行业变化",
        ("1整体情况", "行业变化"),
        "category",
        stop_at_blank_row=True,
        max_data_rows=None,
    ),
    FinalTableSpec("02_总规模", ("2管理人总规模", "总规模带格式", "总规模"), "total", max_data_rows=30),
    FinalTableSpec("03_非货", ("3管理人非货", "非货"), "non_money", max_data_rows=30),
    FinalTableSpec("04_非货增量", ("4非货增量", "非货增量排名"), "increment", max_data_rows=30),
    FinalTableSpec("05_主动权益", ("5主动权益", "主动权益排名"), "business", max_data_rows=30),
    FinalTableSpec(
        "06_非货ETF含联接",
        ("6权益ETF（含联接）", "ETF含联接"),
        "etf",
        col_start=1,
        # None：读连续表头至空列/右表起点（Q3 左表可为 13 列，不再写死 10）
        col_end=None,
        max_data_rows=30,
    ),
    FinalTableSpec(
        "06_权益ETF含联接",
        ("6权益ETF（含联接）",),
        "etf",
        # 规范化后右表常从 ≥20 列起；未规范化时动态定位
        col_start=0,
        col_end=None,
        max_data_rows=30,
    ),
    FinalTableSpec("07_货币", ("7货币", "货币"), "business", max_data_rows=30),
    FinalTableSpec("08_固收", ("8固收", "固收排名"), "business", max_data_rows=30),
    FinalTableSpec("09_固收+", ("9固收+", "固收+排名"), "business", max_data_rows=30),
    FinalTableSpec("10_FOF", ("10 FOF", "10FOF", "FOF"), "business", max_data_rows=30),
)


def _norm_header(v: Any) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return str(v).replace("\n", "").replace(" ", "").strip()


def _display_header(v: Any) -> str:
    """表头展示：日期数字避免变成 20250930.0。"""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    if isinstance(v, float) and v == int(v):
        iv = int(v)
        if 19000101 <= iv <= 21001231:
            return str(iv)
    if isinstance(v, int) and 19000101 <= v <= 21001231:
        return str(v)
    return str(v).strip()


def _is_chart_helper(h: str) -> bool:
    return h in {"x轴", "Y轴", "显示值", "100"} or h.startswith("x轴")


def _to_number(v: Any) -> Any:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "").replace("%", "")
    if not s or s.lower() in {"nan", "none", "-"}:
        return None
    try:
        return float(s)
    except ValueError:
        return str(v).strip()


def _header_width(header: str) -> int:
    n = max(len(header), 2)
    return min(120, max(40, n * 8))


def _column_spec(header: str, col_idx: int, kind: SheetKind) -> ColumnSpec:
    """按列名推断格式 / 数据条 / 色阶（对齐 Q4 底稿条件格式）。"""
    h = _norm_header(header)
    key = f"c{col_idx}"
    width = _header_width(h)

    # 公司名居中（对齐底稿）
    if h in {"公司", "基金公司", "管理人", "类型"}:
        return ColumnSpec(key, header, width=max(width, 90), align="center", fmt="text")

    # 主排名列（1..N）：无条件格式
    if col_idx == 0 and (h == "排名" or h.endswith("排名")) and "变化" not in h and "增量" not in h:
        return ColumnSpec(key, header, width=40, align="center", fmt="int")

    # 排名变化 → 数据条
    if "排名变化" in h:
        return ColumnSpec(key, header, width=width, align="right", fmt="int", bar="bidirectional")

    # 各类排名（含增量排名）→ 色阶（小红大绿）
    if "排名" in h:
        return ColumnSpec(key, header, width=width, align="center", fmt="int", color_scale="rank")

    # 增速/增幅
    if "增速" in h or "增幅" in h:
        # 行业表：越高越红；其他表一般无色阶
        if kind == "category":
            return ColumnSpec(key, header, width=width, align="center", fmt="pct", color_scale="growth")
        return ColumnSpec(key, header, width=width, align="center", fmt="pct")

    # 行业表：新发 / 净值 / 持营 不加数据条
    if kind == "category" and any(x in h for x in ("新发", "持营", "净值")):
        return ColumnSpec(key, header, width=width, align="center", fmt="int")

    # 增量 → 数据条（红）
    if "增量" in h:
        return ColumnSpec(key, header, width=width, align="right", fmt="int", bar="bidirectional")

    # 业务表：净值 / 持营 → 数据条
    if kind != "category" and any(x in h for x in ("净值变化", "净值影响", "持营")):
        return ColumnSpec(key, header, width=width, align="right", fmt="int", bar="bidirectional")

    # 非货表：分类型规模列 → 蓝色数据条（案例中间「主动权益规模」等）
    # H1 底稿常省略「规模」后缀，仅写「主动权益 / 被动权益 / 固收+ …」
    if kind == "non_money":
        if h in {"非货总计", "非货规模"}:
            return ColumnSpec(
                key, header, width=width, align="right", fmt="int", bar="positive", bar_color="blue"
            )
        breakdown = {
            "主动权益规模",
            "被动权益规模",
            "固收+规模",
            "固收规模",
            "FOF规模",
            "主动权益",
            "被动权益",
            "固收+",
            "固收",
            "FOF",
        }
        if h in breakdown or (h.endswith("规模") and h not in {"非货规模", "非货总计"}):
            return ColumnSpec(
                key, header, width=width, align="right", fmt="int", bar="positive", bar_color="blue"
            )

    # 规模栏 → 蓝色数据条（总规模 / 业务表「规模」/ 货币·非货规模等）
    h_aum = h.rstrip("↓")
    if h_aum in {
        "规模",
        "总规模",
        "非货总计",
        "非货规模",
        "货币规模",
        "上年末规模",
    } or (kind == "total" and h_aum in {"货币", "非货"}):
        return ColumnSpec(
            key, header, width=width, align="right", fmt="int", bar="positive", bar_color="blue"
        )

    # 日期列或规模/新发等数值
    if re.fullmatch(r"\d{8}.*", h) or any(
        x in h for x in ("规模", "新发", "期末", "上季", "年初", "上年末", "24年末", "产品分类", "总计")
    ):
        return ColumnSpec(key, header, width=width, align="center", fmt="int")

    return ColumnSpec(key, header, width=width, align="center", fmt="text")


def _resolve_sheet(xl: pd.ExcelFile, sources: tuple[str, ...]) -> str:
    names = set(xl.sheet_names)
    for s in sources:
        if s in names:
            return s
    raise ValueError(f"未找到终表 sheet，候选={sources}，实际={xl.sheet_names}")


def _find_etf_right_block(raw: pd.DataFrame) -> tuple[int, int] | None:
    """定位权益ETF右表列范围（1-based inclusive）。H1 约 15–24，Q4 约 20–29。"""
    for c in range(raw.shape[1]):
        h = _norm_header(raw.iloc[0, c])
        if "权益ETF" in h and "排名" in h:
            start = c + 1
            # 连续非空表头，遇空列或辅助列停止
            end = start
            for j in range(c, raw.shape[1]):
                hj = _norm_header(raw.iloc[0, j])
                if not hj or _is_chart_helper(hj):
                    break
                end = j + 1
            if end >= start:
                return start, end
    return None


def _company_is_focus(name: Any, focus_company: str, focus_short: str) -> bool:
    if name is None:
        return False
    n = str(name).replace("基金", "").strip()
    short = (focus_short or focus_company or "").replace("基金", "").strip()
    full = (focus_company or "").strip()
    if short and short in n:
        return True
    if full and full in str(name):
        return True
    return False


def _read_block(
    path: Path,
    sheet: str,
    *,
    col_start: int,
    col_end: int | None,
    stop_at_blank_row: bool,
    max_data_rows: int | None,
    kind: SheetKind,
    focus_company: str = "",
    focus_short: str = "",
) -> tuple[list[ColumnSpec], list[dict]]:
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        raw = pd.read_excel(path, sheet_name=sheet, header=None, engine="openpyxl")

    # 权益ETF 右表：按表头动态定位（H1/Q4 列起点不同）
    if kind == "etf" and col_start == 0:
        found = _find_etf_right_block(raw)
        if not found:
            return [], []
        col_start, col_end = found

    start = col_start - 1
    if start < 0 or start >= raw.shape[1]:
        return [], []

    if col_end is not None:
        end = min(col_end, raw.shape[1])
        if end <= start:
            return [], []
        headers_raw = [raw.iloc[0, c] for c in range(start, end)]
    else:
        headers_raw = []
        c = start
        # H1 行业表：首格常为空，第 2 格起为日期 → 补「类型」表头
        if kind == "category" and not _norm_header(raw.iloc[0, c]):
            nxt = _norm_header(raw.iloc[0, c + 1]) if c + 1 < raw.shape[1] else ""
            if re.match(r"^\d{8}", nxt):
                headers_raw.append("类型")
                c = start  # 数据仍从空首列读类型名；表头列表先占「类型」
                # 真实表头从 start+1 起追加；end = start + len(headers)
                for j in range(start + 1, raw.shape[1]):
                    h = _norm_header(raw.iloc[0, j])
                    if not h or _is_chart_helper(h):
                        break
                    headers_raw.append(raw.iloc[0, j])
                end = start + len(headers_raw)
            else:
                return [], []
        else:
            while c < raw.shape[1]:
                h = _norm_header(raw.iloc[0, c])
                if not h or _is_chart_helper(h):
                    break
                # ETF 左表：勿跨过右表起点（无空列间隔时）
                if (
                    kind == "etf"
                    and col_start == 1
                    and "权益ETF" in h
                    and "排名" in h
                ):
                    break
                headers_raw.append(raw.iloc[0, c])
                c += 1
            end = start + len(headers_raw)

    if not headers_raw or all(_norm_header(h) == "" for h in headers_raw):
        return [], []

    display_headers = [_display_header(h) for h in headers_raw]
    columns = [_column_spec(h, i, kind) for i, h in enumerate(display_headers)]
    focus_keys = {_norm_header(h) for h in ("公司", "基金公司", "管理人", "类型")}
    company_idx = next(
        (i for i, h in enumerate(display_headers) if _norm_header(h) in focus_keys),
        0,
    )

    def _row_from_vals(vals: list[Any]) -> dict[str, Any]:
        row: dict[str, Any] = {}
        for i, col in enumerate(columns):
            row[col.key] = _to_number(vals[i])
        co_key = columns[company_idx].key
        raw_co = row.get(co_key, "")
        if raw_co is None:
            raw_co = ""
        co = str(raw_co).strip()
        if co.endswith("基金") and len(co) > 2:
            co = co[:-2]
        row["company"] = co
        if _norm_header(display_headers[company_idx]) in {"公司", "基金公司", "管理人"}:
            row[co_key] = co
        return row

    rows: list[dict] = []
    focus_extra: dict | None = None
    for r in range(1, len(raw)):
        vals = [raw.iloc[r, c] for c in range(start, end)]
        if len(vals) < len(columns):
            vals.extend([None] * (len(columns) - len(vals)))
        if stop_at_blank_row and all(
            v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() == ""
            for v in vals
        ):
            break
        first = vals[0]
        if first is None or (isinstance(first, float) and pd.isna(first)):
            if stop_at_blank_row and rows:
                break
            continue
        first_s = str(first).strip()
        if first_s in {"排名", "类型"} or first_s.startswith("非货top"):
            continue
        if kind == "category":
            if first_s not in _CATEGORY_ORDER:
                continue
        else:
            try:
                float(first)
            except (TypeError, ValueError):
                continue

        row = _row_from_vals(vals)
        at_cap = max_data_rows is not None and len(rows) >= max_data_rows
        if at_cap:
            # TopN 截断后仍保留关注公司一行（如固收第 32 的示例）
            if (
                focus_extra is None
                and kind != "category"
                and _company_is_focus(row.get("company"), focus_company, focus_short)
            ):
                focus_extra = row
            continue
        rows.append(row)

    if focus_extra is not None and not any(
        _company_is_focus(r.get("company"), focus_company, focus_short) for r in rows
    ):
        rows.append(focus_extra)

    if kind == "category" and rows:
        order = {name: i for i, name in enumerate(_CATEGORY_ORDER)}
        rows.sort(key=lambda r: order.get(str(r.get("company", "")), 99))

    return columns, rows


def load_final_tables(
    xlsx_path: str | Path,
    *,
    focus_company: str = "示例基金",
    focus_short: str = "示例",
) -> list[tuple[str, list[ColumnSpec], list[dict]]]:
    """从底稿 xlsx 读取全部终表块。关注公司若在 TopN 外仍保留一行。"""
    path = Path(xlsx_path)
    if not path.exists():
        raise FileNotFoundError(path)

    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        xl = pd.ExcelFile(path, engine="openpyxl")

    sheets: list[tuple[str, list[ColumnSpec], list[dict]]] = []
    for spec in FINAL_TABLE_SPECS:
        try:
            sheet_name = _resolve_sheet(xl, spec.sources)
        except ValueError:
            # 第 6 页右表仅 Q4 有；缺则跳过
            if spec.out_name == "06_权益ETF含联接":
                continue
            raise
        cols, rows = _read_block(
            path,
            sheet_name,
            col_start=spec.col_start,
            col_end=spec.col_end,
            stop_at_blank_row=spec.stop_at_blank_row,
            max_data_rows=spec.max_data_rows,
            kind=spec.kind,
            focus_company=focus_company,
            focus_short=focus_short,
        )
        if not cols:
            # 可选块（如白表缺权益ETF右表）直接跳过
            continue
        sheets.append((spec.out_name, cols, rows))
    if not sheets:
        raise ValueError(f"未能从 {path} 读出任何终表")
    return sheets


def export_tables_from_final_xlsx(
    xlsx_path: str | Path,
    output_path: str | Path,
    *,
    period_label: str = "",
    focus_company: str = "示例",
    normalize: bool = True,
) -> Path:
    """终表 → 带数据条/色阶的表图工作簿。默认先规范化为案例口径。"""
    path = Path(xlsx_path)
    if normalize:
        from briefing.importers.normalize_draft import normalize_draft_xlsx

        norm_path = Path(output_path).with_name(Path(output_path).stem + "_normalized_draft.xlsx")
        path = normalize_draft_xlsx(
            path,
            norm_path,
            focus_company=focus_company,
            focus_short=str(focus_company).replace("基金", "") or "示例",
        )

    short = str(focus_company).replace("基金", "") or "示例"
    sheets = load_final_tables(path, focus_company=focus_company, focus_short=short)
    # 推断期别
    if not period_label:
        stem = Path(xlsx_path).stem
        for token in (
            "26H1",
            "26H2",
            "26Q4",
            "25Q4",
            "25Q3",
            "25Q2",
            "25Q1",
            "25H2",
            "25H1",
            "24H2",
            "24H1",
        ):
            if token.lower() in stem.lower():
                period_label = token
                break
        period_label = period_label or "本期"

    return build_deck_workbook(
        sheets,
        output_path,
        period_label=period_label,
        focus_company=focus_company,
    )
