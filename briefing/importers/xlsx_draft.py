"""从「简报底稿」xlsx 导入为标准 fund_metrics.csv。

支持工作簿（如 25H1简报底稿数据.xlsx / 25Q4简报底稿.xlsx）中的典型 sheet：
行业变化|1整体情况 / 总规模带格式|2管理人总规模 / 非货|3管理人非货 /
非货增量排名|4非货增量 / 主动权益排名|5主动权益 /
ETF含联接|6权益ETF（含联接） / 货币|7货币 / 固收排名|8固收 /
固收+排名|9固收+ / FOF|10 FOF / 被动权益排名 / Sheet1
"""

from __future__ import annotations

import re
import warnings
from pathlib import Path

import pandas as pd

# 行业变化「类型」列 → category key
_INDUSTRY_LABEL = {
    "合计": "total",
    "非货": "non_money",
    "货币": "money",
    "固收": "fixed_income",
    "被动权益": "passive_equity",
    "主动权益": "active_equity",
    "固收+": "fixed_income_plus",
    "FOF": "fof",
}

_INDUSTRY_SHEETS = ("行业变化", "1整体情况")

# 排名类 sheet → category（含 Q4 底稿别名）
_RANK_SHEETS = {
    "总规模带格式": "total",
    "总规模": "total",
    "2管理人总规模": "total",
    "非货": "non_money",
    "3管理人非货": "non_money",
    "主动权益排名": "active_equity",
    "5主动权益": "active_equity",
    "ETF含联接": "passive_equity_etf_with_link",
    "6权益ETF（含联接）": "passive_equity_etf_with_link",
    "ETF不含联接": "passive_equity_etf_no_link",
    "货币": "money",
    "7货币": "money",
    "固收排名": "fixed_income",
    "8固收": "fixed_income",
    "固收+排名": "fixed_income_plus",
    "9固收+": "fixed_income_plus",
    "FOF": "fof",
    "10 FOF": "fof",
    "10FOF": "fof",
    "被动权益排名": "passive_equity",
}

_TOTAL_SHEET_CANDIDATES = ("总规模带格式", "2管理人总规模", "总规模")
_INCREMENT_SHEETS = ("非货增量排名", "4非货增量")

# Sheet1 宽表列 → category
_SHEET1_COLS = {
    "总计": "total",
    "货币": "money",
    "主动权益": "active_equity",
    "被动权益": "passive_equity",
    "固收": "fixed_income",
    "固收+": "fixed_income_plus",
    "FOF": "fof",
}


def _norm_header(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = str(v).replace("\n", "").replace(" ", "").strip()
    return s


def _to_float(v, default: float = 0.0) -> float:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return default
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "").replace("%", "")
    if not s or s.lower() in {"nan", "none", "-"}:
        return default
    try:
        return float(s)
    except ValueError:
        return default


def _norm_company(name: str) -> str:
    """统一管理人简称，避免「银华」与「银华基金」各占一行挤掉 TopN。"""
    s = str(name).strip()
    for suffix in (
        "基金管理有限公司",
        "基金股份有限公司",
        "基金有限公司",
        "基金管理公司",
        "基金",
    ):
        if s.endswith(suffix) and len(s) > len(suffix):
            s = s[: -len(suffix)]
            break
    if s.endswith("资产管理") and len(s) > 4:
        s = s[:-4]
    return s


def _detect_dates_from_industry(raw: pd.DataFrame) -> tuple[str, str, str]:
    """从行业变化表头识别 期末/上季/年初。

    - 常规：3 列 YYYYMMDD（期末→上季末→年初）
    - Q1：可仅 2 列（期末→年初）；上季末回退为年初
    """
    header = [_norm_header(x) for x in raw.iloc[0].tolist()]
    dates = []
    for h in header:
        m = re.match(r"^(\d{8})", h)
        if m:
            dates.append(m.group(1))
    if len(dates) >= 3:
        return dates[0], dates[1], dates[2]
    if len(dates) == 2:
        # Q1：期末 + 年初
        return dates[0], dates[1], dates[1]
    raise ValueError(f"行业变化表头需至少 2 个 YYYYMMDD 日期列（期末[+上季末]+年初）: {header[:6]}")


def _parse_industry(path: Path, sheet: str, current: str, prev_q: str, year_start: str) -> list[dict]:
    raw = pd.read_excel(path, sheet_name=sheet, header=None)
    rows: list[dict] = []
    for i in range(1, len(raw)):
        label = raw.iloc[i, 0]
        if label is None or (isinstance(label, float) and pd.isna(label)):
            # 空行：若已读到主品类则结束（Q4 底稿在 FOF 后空行再接 REITs）
            if rows:
                break
            continue
        label = str(label).strip()
        cat = _INDUSTRY_LABEL.get(label)
        if not cat:
            # H1 底稿可能把 REITs/另类插在中间：跳过即可，勿中断
            continue
        aum_c = _to_float(raw.iloc[i, 1])
        aum_pq = _to_float(raw.iloc[i, 2])
        aum_ys = _to_float(raw.iloc[i, 3])
        ni = _to_float(raw.iloc[i, 6]) if raw.shape[1] > 6 else 0.0
        nc = _to_float(raw.iloc[i, 7]) if raw.shape[1] > 7 else 0.0
        hs = _to_float(raw.iloc[i, 8]) if raw.shape[1] > 8 else 0.0
        for d, aum, flows in [
            (current, aum_c, (ni, nc, hs)),
            (prev_q, aum_pq, (0.0, 0.0, 0.0)),
            (year_start, aum_ys, (0.0, 0.0, 0.0)),
        ]:
            rows.append(
                {
                    "company": "__industry__",
                    "category": cat,
                    "date": d,
                    "aum": aum,
                    "new_issue": flows[0],
                    "nav_change": flows[1],
                    "holding_sales": flows[2],
                }
            )
    return rows


def _find_col(headers: list[str], *candidates: str) -> int | None:
    for c in candidates:
        for i, h in enumerate(headers):
            if c and c in h:
                return i
    return None


def _find_ytd_inc_col(headers: list[str]) -> int | None:
    """优先匹配全年/半年度增量，避免误命中季度增量。"""
    return _find_col(
        headers,
        "25年增量",
        "上半年增量",
        "YTD增量",
        "规模增量",
        "非货增量",
        "年增量",
        "H1增量",
        "H2增量",
    )


def _find_q_inc_col(headers: list[str]) -> int | None:
    return _find_col(
        headers,
        "季度增量",
        "Q4增量",
        "Q3增量",
        "Q2增量",
        "Q1增量",
        "当季增量",
    )


def _parse_rank_sheet(
    path: Path,
    sheet: str,
    category: str,
    current: str,
    prev_q: str,
    year_start: str,
) -> list[dict]:
    raw = pd.read_excel(path, sheet_name=sheet, header=None)
    headers = [_norm_header(x) for x in raw.iloc[0].tolist()]
    # 部分表第 2 行是副表头，数据从第 2 或第 3 行开始
    start_row = 1
    if raw.shape[0] > 1:
        second = _norm_header(raw.iloc[1, 0])
        if second in {"", "排名变化", "排名"} or not str(raw.iloc[1, 0]).replace(".", "").isdigit():
            # 若第二行第一列不是数字排名，可能是副表头
            try:
                float(raw.iloc[1, 0])
            except (TypeError, ValueError):
                start_row = 2

    c_company = _find_col(headers, "基金公司", "公司", "管理人")
    c_aum = _find_col(headers, "总规模", "非货总计", "非货规模", "规模")
    c_inc = _find_ytd_inc_col(headers)
    if c_inc is None:
        # 兜底：含「增量」但不含「季/Q」的列
        for i, h in enumerate(headers):
            if "增量" in h and "季" not in h and not re.search(r"Q[1-4]", h):
                c_inc = i
                break
    c_q_inc = _find_q_inc_col(headers)
    c_ni = _find_col(headers, "新发")
    c_nc = _find_col(headers, "净值变化", "净值")
    c_hs = _find_col(headers, "持营")
    c_money = _find_col(headers, "货币规模", "货币") if category == "total" else None
    c_non = _find_col(headers, "非货规模", "非货") if category == "total" else None
    # 非货拆解规模
    sub_cols = {}
    if category == "non_money":
        for label, key in [
            ("主动权益规模", "active_equity"),
            ("被动权益规模", "passive_equity"),
            ("固收+规模", "fixed_income_plus"),
            ("固收规模", "fixed_income"),
            ("FOF规模", "fof"),
        ]:
            idx = _find_col(headers, label)
            if idx is not None:
                sub_cols[key] = idx

    if c_company is None or c_aum is None:
        return []

    rows: list[dict] = []
    seen: set[str] = set()

    def add_point(
        company: str,
        cat: str,
        aum_c: float,
        ytd_inc: float,
        q_inc: float | None,
        ni: float,
        nc: float,
        hs: float,
    ):
        aum_ys = aum_c - ytd_inc
        if q_inc is not None:
            aum_pq = aum_c - q_inc
        else:
            # 无上季数据时，用年初与期末中点近似
            aum_pq = (aum_ys + aum_c) / 2.0
        for d, aum, flows in [
            (current, aum_c, (ni, nc, hs)),
            (prev_q, aum_pq, (0.0, 0.0, 0.0)),
            (year_start, aum_ys, (0.0, 0.0, 0.0)),
        ]:
            rows.append(
                {
                    "company": company,
                    "category": cat,
                    "date": d,
                    "aum": round(aum, 4),
                    "new_issue": flows[0],
                    "nav_change": flows[1],
                    "holding_sales": flows[2],
                }
            )

    for i in range(start_row, len(raw)):
        company = raw.iloc[i, c_company]
        if company is None or (isinstance(company, float) and pd.isna(company)):
            continue
        company = str(company).strip()
        if not company or company in {"基金公司", "公司", "管理人", "nan"}:
            continue
        company = _norm_company(company)
        if not company or company in seen:
            continue
        # 排名列应为数字
        rank_val = raw.iloc[i, 0]
        try:
            float(rank_val)
        except (TypeError, ValueError):
            continue

        aum_c = _to_float(raw.iloc[i, c_aum])
        ytd_inc = _to_float(raw.iloc[i, c_inc]) if c_inc is not None else 0.0
        q_inc = _to_float(raw.iloc[i, c_q_inc]) if c_q_inc is not None else None
        ni = _to_float(raw.iloc[i, c_ni]) if c_ni is not None else 0.0
        nc = _to_float(raw.iloc[i, c_nc]) if c_nc is not None else 0.0
        hs = _to_float(raw.iloc[i, c_hs]) if c_hs is not None else 0.0
        seen.add(company)
        add_point(company, category, aum_c, ytd_inc, q_inc, ni, nc, hs)

        if category == "total":
            if c_money is not None:
                m = _to_float(raw.iloc[i, c_money])
                add_point(company, "money", m, 0.0, None, 0.0, 0.0, 0.0)
            if c_non is not None:
                n = _to_float(raw.iloc[i, c_non])
                add_point(company, "non_money", n, 0.0, None, 0.0, 0.0, 0.0)

        for sub_cat, idx in sub_cols.items():
            add_point(company, sub_cat, _to_float(raw.iloc[i, idx]), 0.0, None, 0.0, 0.0, 0.0)

    return rows


def _parse_sheet1(path: Path, current: str, year_start: str, prev_q: str) -> list[dict]:
    """全市场期末规模宽表；仅补齐排名表未覆盖的公司×品类。"""
    sheet_name = "Sheet1"
    try:
        raw = pd.read_excel(path, sheet_name=sheet_name, header=0)
    except ValueError:
        return []

    colmap = {}
    # 长标签优先，避免「固收+」被「固收」前缀误匹配
    labels_sorted = sorted(_SHEET1_COLS.items(), key=lambda kv: -len(kv[0]))
    for c in raw.columns:
        name = _norm_header(c)
        # 去掉 .1 后缀的排名列
        if name.endswith(".1"):
            continue
        for label, key in labels_sorted:
            if name == label:
                colmap[key] = c
                break
    name_col = raw.columns[0]
    rows: list[dict] = []
    for _, r in raw.iterrows():
        company = r[name_col]
        if company is None or (isinstance(company, float) and pd.isna(company)):
            continue
        company = str(company).strip()
        if not company or company in {"行标签", "总计", "合计"}:
            continue
        company = _norm_company(company)
        if not company:
            continue
        values = {}
        for key, col in colmap.items():
            values[key] = _to_float(r[col])
        # 非货 = 总计 - 货币（若两者都有）
        if "total" in values and "money" in values:
            values["non_money"] = values["total"] - values["money"]
        for cat, aum in values.items():
            if aum == 0:
                continue
            # Sheet1 无增量：三时点先同值，后续由排名表覆盖有增量的公司
            for d in (current, prev_q, year_start):
                rows.append(
                    {
                        "company": company,
                        "category": cat,
                        "date": d,
                        "aum": round(aum, 4),
                        "new_issue": 0.0,
                        "nav_change": 0.0,
                        "holding_sales": 0.0,
                    }
                )
    return rows


def _merge_rows(rows: list[dict]) -> pd.DataFrame:
    """合并同一 company/category/date：保留更高精度规模；流水优先非零。"""
    cols = ["company", "category", "date", "aum", "new_issue", "nav_change", "holding_sales"]
    if not rows:
        return pd.DataFrame(columns=cols)

    buckets: dict[tuple, dict] = {}
    for r in rows:
        key = (r["company"], r["category"], r["date"])
        aum = float(r["aum"])
        ni, nc, hs = float(r["new_issue"]), float(r["nav_change"]), float(r["holding_sales"])
        if key not in buckets:
            buckets[key] = {
                "company": r["company"],
                "category": r["category"],
                "date": r["date"],
                "aum": aum,
                "new_issue": ni,
                "nav_change": nc,
                "holding_sales": hs,
            }
            continue
        cur = buckets[key]
        old = float(cur["aum"])
        old_precise = abs(old - round(old)) > 1e-9
        new_precise = abs(aum - round(aum)) > 1e-9
        # Sheet1 高精度期末值，不被排名表整数覆盖
        if old_precise and not new_precise and abs(old - aum) < 1.5:
            pass
        else:
            cur["aum"] = aum
        for field, val in (("new_issue", ni), ("nav_change", nc), ("holding_sales", hs)):
            if abs(val) > 1e-12:
                cur[field] = val

    df = pd.DataFrame(list(buckets.values()), columns=cols)
    return df.sort_values(["category", "company", "date"]).reset_index(drop=True)


def import_draft_xlsx(
    xlsx_path: str | Path,
    output_csv: str | Path | None = None,
    prefer_total_sheet: str = "总规模带格式",
) -> pd.DataFrame:
    """
    读取底稿 xlsx，写出 fund_metrics.csv，并返回 DataFrame。

    日期默认从「行业变化 / 1整体情况」表头识别。
    """
    path = Path(xlsx_path)
    if not path.exists():
        raise FileNotFoundError(path)

    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        xl = pd.ExcelFile(path)
        sheets = set(xl.sheet_names)

        industry_sheet = next((s for s in _INDUSTRY_SHEETS if s in sheets), None)
        if industry_sheet is None:
            raise ValueError(f"未找到行业变化表，候选: {_INDUSTRY_SHEETS}，实际: {sorted(sheets)}")

        industry_raw = pd.read_excel(path, sheet_name=industry_sheet, header=None)
        current, prev_q, year_start = _detect_dates_from_industry(industry_raw)

        all_rows: list[dict] = []
        all_rows.extend(_parse_industry(path, industry_sheet, current, prev_q, year_start))

        # 先 Sheet1 打底，再排名表覆盖（含增量与流水）
        all_rows.extend(_parse_sheet1(path, current, year_start, prev_q))

        # 总规模：优先指定表，再按 Q4/H1 候选
        total_sheet = None
        if prefer_total_sheet in sheets:
            total_sheet = prefer_total_sheet
        else:
            total_sheet = next((s for s in _TOTAL_SHEET_CANDIDATES if s in sheets), None)
        if total_sheet:
            all_rows.extend(
                _parse_rank_sheet(path, total_sheet, "total", current, prev_q, year_start)
            )

        for sheet, cat in _RANK_SHEETS.items():
            if sheet in _TOTAL_SHEET_CANDIDATES or sheet == prefer_total_sheet:
                continue
            if sheet not in sheets:
                continue
            all_rows.extend(_parse_rank_sheet(path, sheet, cat, current, prev_q, year_start))

        # 非货增量排名：补非货增量（覆盖 non_money 的年初回推）
        for inc_sheet in _INCREMENT_SHEETS:
            if inc_sheet in sheets:
                all_rows.extend(
                    _parse_non_money_increment(path, inc_sheet, current, prev_q, year_start)
                )
                break

    df = _merge_rows(all_rows)
    if output_csv is not None:
        out = Path(output_csv)
        out.parent.mkdir(parents=True, exist_ok=True)
        df.to_csv(out, index=False, encoding="utf-8")
    return df


def _parse_non_money_increment(
    path: Path, sheet: str, current: str, prev_q: str, year_start: str
) -> list[dict]:
    raw = pd.read_excel(path, sheet_name=sheet, header=None)
    headers = [_norm_header(x) for x in raw.iloc[0].tolist()]
    c_company = _find_col(headers, "基金公司", "公司", "管理人")
    c_inc = _find_ytd_inc_col(headers) or _find_col(headers, "非货增量", "增量")
    c_q_inc = _find_q_inc_col(headers)
    if c_company is None or c_inc is None:
        return []

    # 从非货表取期末规模
    xl = pd.ExcelFile(path)
    nm_sheet = next((s for s in ("非货", "3管理人非货") if s in xl.sheet_names), None)
    if nm_sheet is None:
        return []
    nm = pd.read_excel(path, sheet_name=nm_sheet, header=None)
    nm_h = [_norm_header(x) for x in nm.iloc[0].tolist()]
    nm_co = _find_col(nm_h, "基金公司", "公司", "管理人")
    nm_aum = _find_col(nm_h, "非货总计", "非货规模", "规模")
    nm_inc = _find_ytd_inc_col(nm_h) or _find_col(nm_h, "上半年增量", "增量")
    nm_q = _find_q_inc_col(nm_h)
    aum_map = {}
    inc_map = {}
    q_map = {}
    start = 1
    try:
        float(nm.iloc[1, 0])
    except (TypeError, ValueError):
        start = 2
    if nm_co is not None and nm_aum is not None:
        for i in range(start, len(nm)):
            co = nm.iloc[i, nm_co]
            if co is None or (isinstance(co, float) and pd.isna(co)):
                continue
            co = str(co).strip()
            aum_map[_norm_company(co)] = _to_float(nm.iloc[i, nm_aum])
            if nm_inc is not None:
                inc_map[_norm_company(co)] = _to_float(nm.iloc[i, nm_inc])
            if nm_q is not None:
                q_map[_norm_company(co)] = _to_float(nm.iloc[i, nm_q])

    rows = []
    for i in range(1, len(raw)):
        co = raw.iloc[i, c_company]
        if co is None or (isinstance(co, float) and pd.isna(co)):
            continue
        co = _norm_company(str(co).strip())
        if co not in aum_map:
            continue
        aum_c = aum_map[co]
        inc = inc_map.get(co, _to_float(raw.iloc[i, c_inc]))
        aum_ys = aum_c - inc
        if co in q_map:
            aum_pq = aum_c - q_map[co]
        elif c_q_inc is not None:
            aum_pq = aum_c - _to_float(raw.iloc[i, c_q_inc])
        else:
            aum_pq = (aum_ys + aum_c) / 2.0
        for d, aum in [(current, aum_c), (prev_q, aum_pq), (year_start, aum_ys)]:
            rows.append(
                {
                    "company": co,
                    "category": "non_money",
                    "date": d,
                    "aum": round(aum, 4),
                    "new_issue": 0.0,
                    "nav_change": 0.0,
                    "holding_sales": 0.0,
                }
            )
    return rows


def detect_dates(xlsx_path: str | Path) -> dict[str, str]:
    path = Path(xlsx_path)
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        xl = pd.ExcelFile(path)
        sheets = set(xl.sheet_names)
        industry_sheet = next((s for s in _INDUSTRY_SHEETS if s in sheets), None)
        if industry_sheet is None:
            raise ValueError(f"未找到行业变化表，候选: {_INDUSTRY_SHEETS}")
        raw = pd.read_excel(path, sheet_name=industry_sheet, header=None)
    current, prev_q, year_start = _detect_dates_from_industry(raw)
    return {"current": current, "previous_quarter": prev_q, "year_start": year_start}
