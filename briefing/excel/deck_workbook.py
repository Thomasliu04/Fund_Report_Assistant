"""全册表图 Excel：各页一张 sheet，带数据条/色阶/银华高亮，便于人工调美观。"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from briefing.excel.databar_patch import patch_workbook_bidirectional_databars
from briefing.excel.styled_table import ColumnSpec

# 与 PPT 一致：正红
_BAR_POS = "E85050"


def _cell_value(val: Any, fmt: str):
    if val is None:
        return None
    if fmt == "int":
        try:
            return int(round(float(val)))
        except (TypeError, ValueError):
            return val
    if fmt == "pct":
        try:
            v = float(val)
            # 底稿多为 0.15；若已是 15 这类百分数则先还原为小数
            if abs(v) > 2:
                v = v / 100.0
            return round(v, 6)
        except (TypeError, ValueError):
            return val
    return val


def _add_sheet(
    wb: Workbook,
    title: str,
    columns: list[ColumnSpec],
    rows: list[dict],
    *,
    focus_company: str = "银华基金",
    focus_key: str = "company",
) -> list[str]:
    """写入 sheet，返回需要双向轴补丁的区域列表（如 E2:E31）。"""
    ws = wb.create_sheet(title[:31])
    header_fill = PatternFill("solid", fgColor="F2DCDB")
    header_font = Font(bold=True, color="333333", name="微软雅黑", size=10)
    thin = Border(
        left=Side(style="thin", color="B4C6DC"),
        right=Side(style="thin", color="B4C6DC"),
        top=Side(style="thin", color="B4C6DC"),
        bottom=Side(style="thin", color="B4C6DC"),
    )
    focus_fill = PatternFill("solid", fgColor="FCE4D6")
    body_font = Font(name="微软雅黑", size=9)
    focus_font = Font(name="微软雅黑", size=9, color="9C0006")

    for c, col in enumerate(columns, 1):
        cell = ws.cell(1, c, col.header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")
        cell.border = thin

    for r_i, row in enumerate(rows, 2):
        is_focus = focus_company != "__none__" and focus_company in str(row.get(focus_key, ""))
        for c, col in enumerate(columns, 1):
            raw = row.get(col.key)
            val = _cell_value(raw, col.fmt)
            cell = ws.cell(r_i, c, val)
            cell.font = focus_font if is_focus else body_font
            cell.border = thin
            align = {"left": "left", "right": "right"}.get(col.align, "center")
            cell.alignment = Alignment(horizontal=align, vertical="center")
            if col.fmt == "pct" and isinstance(val, (int, float)):
                cell.number_format = '0%'
            if is_focus:
                cell.fill = focus_fill

    n = len(rows) + 1
    bi_ranges: list[str] = []
    if n >= 2:
        for c, col in enumerate(columns, 1):
            letter = get_column_letter(c)
            rng = f"{letter}2:{letter}{n}"
            if col.bar != "none":
                ws.conditional_formatting.add(
                    rng,
                    DataBarRule(
                        start_type="min",
                        end_type="max",
                        color=_BAR_POS,
                        showValue=True,
                        minLength=None,
                        maxLength=None,
                    ),
                )
                # 双向列（含正负）打补丁；单向正值列也打补丁无害，轴居中即可
                if col.bar == "bidirectional":
                    bi_ranges.append(rng)
            if col.color_scale == "rank":
                ws.conditional_formatting.add(
                    rng,
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

    for c, col in enumerate(columns, 1):
        ws.column_dimensions[get_column_letter(c)].width = max(6, round(col.width / 7, 1))
    ws.row_dimensions[1].height = 32
    for r in range(2, n + 1):
        ws.row_dimensions[r].height = 16

    if n >= 1:
        last = get_column_letter(len(columns))
        ws.print_area = f"A1:{last}{max(n, 1)}"
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
    return bi_ranges


def _readme_sheet(wb: Workbook, period_label: str) -> None:
    ws = wb.active
    ws.title = "使用说明"
    lines = [
        f"{period_label} 公募行业数据简报 · 表图工作簿",
        "",
        "【分工建议】",
        "· PPT：叙事文字（观点句需人工敲定）+ 贴入表图",
        "· 本 Excel：各页排名表的可编辑源文件（数据条 / 色阶 / 银华高亮）",
        "",
        "【怎么用】",
        "1. 在对应 sheet 里改条件格式、列宽、颜色、字体等美观项",
        "2. 选中表格区域 → 复制 → 选择性粘贴到 PPT，或另存区域为图片再替换",
        "3. 数字一般不要手改；若底稿更新请重新跑生成流程",
        "",
        "【数据条】",
        "正值向右红色，负值向左绿色（与 PPT 表图一致）；轴在零点。",
        "",
        "【Sheet 对应】",
        "01_行业变化 → PPT 第1页表",
        "02_总规模 → 第2页",
        "03_非货 → 第3页",
        "04_非货增量 → 第4页",
        "05_主动权益 → 第5页",
        "06_非货ETF含联接 / 06_权益ETF含联接 → 第6页双表",
        "07_货币 → 第7页",
        "08_固收 → 第8页",
        "09_固收+ → 第9页",
        "10_FOF → 第10页",
        "",
        "【数据来源】",
        "本工作簿以底稿「终表」为准（含他表 VLOOKUP 过来的排名列），按列做数据条/色阶。",
    ]
    ws["A1"] = lines[0]
    ws["A1"].font = Font(name="微软雅黑", size=14, bold=True, color="D32820")
    for i, line in enumerate(lines[1:], 2):
        ws.cell(i, 1, line).font = Font(name="微软雅黑", size=10)
    ws.column_dimensions["A"].width = 88


def build_deck_workbook(
    sheets: list[tuple[str, list[ColumnSpec], list[dict]]],
    output_path: str | Path,
    *,
    period_label: str = "",
    focus_company: str = "银华基金",
) -> Path:
    """
    sheets: [(sheet_title, columns, rows), ...]
    """
    wb = Workbook()
    _readme_sheet(wb, period_label or "本期")
    sheet_ranges: dict[str, list[str]] = {}
    for title, columns, rows in sheets:
        name = title[:31]
        bi = _add_sheet(wb, name, columns, rows, focus_company=focus_company)
        if bi:
            sheet_ranges[name] = bi

    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    if sheet_ranges:
        patch_workbook_bidirectional_databars(out, sheet_ranges)
    return out
