"""期别档案：Q1 单口径 / Q3 单季主导双口径 / H1·Q4 期别主导双口径。"""

from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from typing import Literal


class PeriodMode(str, Enum):
    """叙事模式。"""

    SINGLE_QUARTER = "single_quarter"  # Q1：YTD≡单季
    QUARTER_LEAD = "quarter_lead"  # Q3：单季为主，YTD 并列
    PERIOD_LEAD = "period_lead"  # H1/Q2、Q4：半年度或全年为主


LeadKind = Literal["quarter", "period"]


@dataclass(frozen=True)
class PeriodProfile:
    mode: PeriodMode
    period_type: str  # half_year | quarter | year
    period_label: str
    ytd_tag: str
    quarter_tag: str
    span: str  # 口语窗口：上半年 / 全年 / Q1 / 前三季度 …
    ytd_phrase: str  # 增速/增量前缀
    dual_metrics: bool  # 成文是否双口径
    lead: LeadKind  # 双口径时主句顺序


def _yy_mm(current_date: str) -> tuple[str, int]:
    cur = (current_date or "").strip()
    yy = cur[2:4] if len(cur) >= 4 else "25"
    try:
        mm = int(cur[4:6]) if len(cur) >= 6 else 12
    except ValueError:
        mm = 12
    return yy, mm


def _quarter_num(mm: int) -> int:
    return (mm - 1) // 3 + 1


def _label_tokens(period_label: str) -> str:
    return (period_label or "").strip().upper().replace(" ", "")


def guess_period_from_date(current_date: str) -> tuple[str, str]:
    """从期末日期猜 (period_label, period_type)。"""
    yy, mm = _yy_mm(current_date)
    if mm in (1, 2, 3):
        return f"{yy}Q1", "quarter"
    if mm in (4, 5, 6):
        return f"{yy}H1", "half_year"
    if mm in (7, 8, 9):
        return f"{yy}Q3", "quarter"
    return f"{yy}Q4", "quarter"


def resolve_period_profile(
    period_label: str = "",
    period_type: str = "",
    current_date: str = "",
) -> PeriodProfile:
    """
    统一期别判定。

    优先级：label 中的 H1/H2/Q1–Q4 > period_type=half_year > 期末月份。
    """
    lab = _label_tokens(period_label)
    yy, mm = _yy_mm(current_date)
    qn = _quarter_num(mm)
    q_tag = f"Q{qn}"
    ptype = (period_type or "").strip().lower()

    # —— H1 / H2 / 显式半年度 ——
    if "H1" in lab or (ptype == "half_year" and "H2" not in lab and "Q" not in lab and mm <= 6):
        label = period_label.strip() if period_label.strip() else f"{yy}H1"
        if "H1" not in _label_tokens(label) and ptype == "half_year":
            label = f"{yy}H1"
        return PeriodProfile(
            mode=PeriodMode.PERIOD_LEAD,
            period_type="half_year",
            period_label=label,
            ytd_tag=label if "H" in _label_tokens(label) else f"{yy}H1",
            quarter_tag=q_tag,
            span="上半年",
            ytd_phrase=label if "H" in _label_tokens(label) else f"{yy}H1",
            dual_metrics=True,
            lead="period",
        )
    if "H2" in lab or (ptype == "half_year" and mm >= 7):
        label = period_label.strip() if period_label.strip() else f"{yy}H2"
        return PeriodProfile(
            mode=PeriodMode.PERIOD_LEAD,
            period_type="half_year",
            period_label=label,
            ytd_tag=label if "H" in _label_tokens(label) else f"{yy}H2",
            quarter_tag=q_tag,
            span="下半年",
            ytd_phrase=label if "H" in _label_tokens(label) else f"{yy}H2",
            dual_metrics=True,
            lead="period",
        )

    # —— Q1 ——
    if "Q1" in lab or (qn == 1 and "Q" not in lab and "H" not in lab):
        label = period_label.strip() if period_label.strip() else f"{yy}Q1"
        if "Q1" not in _label_tokens(label):
            label = f"{yy}Q1"
        return PeriodProfile(
            mode=PeriodMode.SINGLE_QUARTER,
            period_type="quarter",
            period_label=label,
            ytd_tag=label,
            quarter_tag="Q1",
            span="Q1",
            ytd_phrase=label,
            dual_metrics=False,
            lead="quarter",
        )

    # —— Q3 ——
    if "Q3" in lab or (qn == 3 and "Q" not in lab and "H" not in lab):
        label = period_label.strip() if period_label.strip() else f"{yy}Q3"
        if "Q3" not in _label_tokens(label):
            label = f"{yy}Q3"
        return PeriodProfile(
            mode=PeriodMode.QUARTER_LEAD,
            period_type="quarter",
            period_label=label,
            ytd_tag="前三季度",
            quarter_tag="Q3",
            span="前三季度",
            ytd_phrase="前三季度",
            dual_metrics=True,
            lead="quarter",
        )

    # —— Q2：与 H1 同族（半年度主导）——
    if "Q2" in lab or (qn == 2 and ptype != "quarter" and "Q" not in lab):
        # 显式 Q2 仍可按季度类型；若用户选 quarter+Q2，当作 period_lead 半年度口径
        label = period_label.strip() if period_label.strip() else f"{yy}H1"
        if "Q2" in lab:
            # 保留 Q2 标签，但叙事按半年度主导（案例：Q2 像 H1）
            ytd = label
            span = "上半年"
        else:
            ytd = f"{yy}H1"
            span = "上半年"
        return PeriodProfile(
            mode=PeriodMode.PERIOD_LEAD,
            period_type="half_year" if "H" in _label_tokens(label) else "quarter",
            period_label=label,
            ytd_tag=ytd,
            quarter_tag="Q2",
            span=span,
            ytd_phrase=ytd,
            dual_metrics=True,
            lead="period",
        )

    # —— Q4 / 年末 / 默认全年 ——
    if "Q4" in lab or qn == 4 or ptype == "year":
        label = period_label.strip() if period_label.strip() else f"{yy}Q4"
        if "Q4" not in _label_tokens(label) and ptype != "year":
            label = f"{yy}Q4"
        ytd = f"{yy}年"
        return PeriodProfile(
            mode=PeriodMode.PERIOD_LEAD,
            period_type="quarter" if "Q" in _label_tokens(label) else ("year" if ptype == "year" else "quarter"),
            period_label=label,
            ytd_tag=ytd,
            quarter_tag="Q4",
            span="全年",
            ytd_phrase=ytd,
            dual_metrics=True,
            lead="period",
        )

    # 兜底：按月份
    guessed_label, guessed_type = guess_period_from_date(current_date)
    return resolve_period_profile(guessed_label, guessed_type, current_date)
