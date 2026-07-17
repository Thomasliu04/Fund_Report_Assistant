"""构建带 Excel 条件格式（数据条/色阶）的总规模排名表。"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


HEADERS = [
    "排名",
    "排名变化",
    "基金公司",
    "总规模",
    "上半年增量",
    "上半年增速%",
    "货币",
    "非货",
    "货币排名",
    "非货排名",
]


def build_total_ranking_workbook(rows: list[dict], output_path: str | Path, focus_company: str = "银华基金") -> Path:
    """
    rows 每项字段::
        rank, rank_change, company, aum, increment, growth_pct,
        money, non_money, money_rank, non_money_rank
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "总规模排名"

    header_fill = PatternFill("solid", fgColor="F2DCDB")
    header_font = Font(bold=True, color="333333", name="微软雅黑", size=10)
    thin = Border(
        left=Side(style="thin", color="B4C6DC"),
        right=Side(style="thin", color="B4C6DC"),
        top=Side(style="thin", color="B4C6DC"),
        bottom=Side(style="thin", color="B4C6DC"),
    )
    focus_fill = PatternFill("solid", fgColor="FCE4D6")

    for c, h in enumerate(HEADERS, 1):
        cell = ws.cell(1, c, h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = thin

    for i, row in enumerate(rows, 2):
        values = [
            row["rank"],
            row["rank_change"],
            row["company"],
            round(row["aum"]),
            round(row["increment"]),
            round(row["growth_pct"] * 100),
            round(row["money"]),
            round(row["non_money"]),
            row["money_rank"],
            row["non_money_rank"],
        ]
        is_focus = focus_company in str(row["company"])
        for c, v in enumerate(values, 1):
            cell = ws.cell(i, c, v)
            cell.font = Font(name="微软雅黑", size=9, color="9C0006" if is_focus else "000000")
            cell.border = thin
            cell.alignment = Alignment(horizontal="center" if c != 3 else "left")
            if is_focus:
                cell.fill = focus_fill

    n = len(rows) + 1  # last data row

    # 数据条：上半年增量 (E)、非货 (H) — 对应示例中的「趋势感」
    ws.conditional_formatting.add(
        f"E2:E{n}",
        DataBarRule(
            start_type="min",
            end_type="max",
            color="638EC6",
            showValue=True,
            minLength=None,
            maxLength=None,
        ),
    )
    ws.conditional_formatting.add(
        f"H2:H{n}",
        DataBarRule(
            start_type="min",
            end_type="max",
            color="638EC6",
            showValue=True,
            minLength=None,
            maxLength=None,
        ),
    )

    # 色阶：货币排名 / 非货排名（数值越小越好 → 红到绿）
    for col in ("I", "J"):
        ws.conditional_formatting.add(
            f"{col}2:{col}{n}",
            ColorScaleRule(
                start_type="min",
                start_color="F8696B",
                mid_type="percentile",
                mid_value=50,
                mid_color="FFEB84",
                end_type="max",
                end_color="63BE7B",
            ),
        )

    widths = [6, 8, 14, 10, 12, 10, 10, 10, 10, 10]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 30
    for r in range(2, n + 1):
        ws.row_dimensions[r].height = 16

    # 定义打印/导出区域 — 横向一页放下全部列
    ws.print_area = f"A1:J{n}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = 0.25
    ws.page_margins.right = 0.25
    ws.page_margins.top = 0.25
    ws.page_margins.bottom = 0.25
    ws.sheet_view.showGridLines = False

    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    # 注意：openpyxl 无法稳定写出「负左正右」的 x14 数据条扩展；
    # 强行补 XML 会导致 Excel 报“内容有问题”。双向轴效果由 PPT 预览渲染负责。
    return out
