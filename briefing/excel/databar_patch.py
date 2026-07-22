"""为 openpyxl 生成的数据条补上 Excel 双向轴：正右红、负左绿。"""

from __future__ import annotations

import re
import uuid
import zipfile
from pathlib import Path
from tempfile import TemporaryDirectory
from xml.etree import ElementTree as ET

X14_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
XM_NS = "http://schemas.microsoft.com/office/excel/2006/main"
MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# 与 Q4 底稿一致：正红 / 负绿；渐变
POS_COLOR = "FFFF555A"
NEG_COLOR = "FF00B050"
AXIS_COLOR = "FF000000"
ID_EXT_URI = "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}"
CF_EXT_URI = "{78C0D931-6437-407d-A8EE-F0AAD7539E65}"


def _guid() -> str:
    return "{" + str(uuid.uuid4()).upper() + "}"


def _sheet_paths_by_name(root: Path) -> dict[str, Path]:
    """sheet 显示名 → worksheet xml 路径。"""
    wb = root / "xl" / "workbook.xml"
    rels = root / "xl" / "_rels" / "workbook.xml.rels"
    if not wb.exists() or not rels.exists():
        return {}

    # 注册默认命名空间，方便找 tag
    ET.register_namespace("", MAIN_NS)
    ET.register_namespace("r", PKG_REL)
    wb_tree = ET.parse(wb)
    rel_tree = ET.parse(rels)
    rid_to_target: dict[str, str] = {}
    for rel in rel_tree.getroot():
        rid = rel.attrib.get("Id")
        target = rel.attrib.get("Target")
        if rid and target:
            rid_to_target[rid] = target

    out: dict[str, Path] = {}
    for sheet in wb_tree.getroot().findall(f"{{{MAIN_NS}}}sheets/{{{MAIN_NS}}}sheet"):
        name = sheet.attrib.get("name")
        rid = sheet.attrib.get(f"{{{PKG_REL}}}id")
        target = rid_to_target.get(rid or "", "")
        if not name or not target:
            continue
        # Target 可能是 "worksheets/sheet2.xml" 或 "/xl/worksheets/sheet2.xml"
        target = target.lstrip("/")
        if target.startswith("xl/"):
            path = root / target
        else:
            path = root / "xl" / target
        if path.exists():
            out[name] = path
    return out


def _inject_rule_ids(xml: str, sqrefs: list[str]) -> tuple[str, list[tuple[str, str]]]:
    """
    为每个 sqref 的 dataBar cfRule 注入 x14:id，返回 (新xml, [(sqref, guid)...])。
    每个 sqref 只处理第一条 dataBar 规则。
    """
    rule_ids: list[tuple[str, str]] = []
    for sqref in sqrefs:
        guid = _guid()
        pattern = (
            rf'(<conditionalFormatting[^>]*sqref="{re.escape(sqref)}"[^>]*>\s*'
            rf'<cfRule[^>]*type="dataBar"[^>]*>)(.*?)(</cfRule>)'
        )

        def _repl(m: re.Match, gid: str = guid) -> str:
            head, body, tail = m.group(1), m.group(2), m.group(3)
            # 修正 openpyxl 写出的 00XXXXXX 颜色为 FFXXXXXX
            body = re.sub(
                r'(<color rgb=")00([0-9A-Fa-f]{6}")',
                rf"\1FF\2",
                body,
            )
            # 统一正色为 POS_COLOR
            body = re.sub(
                r'<color rgb="[0-9A-Fa-f]{8}"/>',
                f'<color rgb="{POS_COLOR}"/>',
                body,
                count=1,
            )
            body = re.sub(r"<extLst>.*?</extLst>", "", body, flags=re.DOTALL)
            ext = (
                f'<extLst><ext uri="{ID_EXT_URI}" xmlns:x14="{X14_NS}">'
                f"<x14:id>{gid}</x14:id></ext></extLst>"
            )
            return head + body + ext + tail

        new_xml, n = re.subn(pattern, _repl, xml, count=1, flags=re.DOTALL)
        if n == 0:
            # 容错：sqref 属性顺序不同
            pattern2 = (
                rf'(<conditionalFormatting[^>]*sqref="{re.escape(sqref)}"[^>]*>)'
                rf'(.*?)(</conditionalFormatting>)'
            )

            def _repl2(m: re.Match, gid: str = guid) -> str:
                head, mid, tail = m.group(1), m.group(2), m.group(3)
                m2 = re.search(
                    r'(<cfRule[^>]*type="dataBar"[^>]*>)(.*?)(</cfRule>)',
                    mid,
                    flags=re.DOTALL,
                )
                if not m2:
                    return m.group(0)
                h, b, t = m2.group(1), m2.group(2), m2.group(3)
                b = re.sub(r'(<color rgb=")00([0-9A-Fa-f]{6}")', r"\1FF\2", b)
                b = re.sub(
                    r'<color rgb="[0-9A-Fa-f]{8}"/>',
                    f'<color rgb="{POS_COLOR}"/>',
                    b,
                    count=1,
                )
                b = re.sub(r"<extLst>.*?</extLst>", "", b, flags=re.DOTALL)
                ext = (
                    f'<extLst><ext uri="{ID_EXT_URI}" xmlns:x14="{X14_NS}">'
                    f"<x14:id>{gid}</x14:id></ext></extLst>"
                )
                mid2 = mid[: m2.start()] + h + b + ext + t + mid[m2.end() :]
                return head + mid2 + tail

            new_xml, n = re.subn(pattern2, _repl2, xml, count=1, flags=re.DOTALL)
            if n == 0:
                continue
        xml = new_xml
        rule_ids.append((sqref, guid))
    return xml, rule_ids


def _x14_block(sqref: str, guid: str) -> str:
    """
    对齐已验证可用的写法：
    - 正色只写在主 dataBar 的 <color>，x14 里不要再写 fillColor
      （否则 Mac Excel 常出现：负值仅边框变绿、填充仍为红）
    - negativeBarColorSameAsPositive="0" + negativeFillColor 实心绿
    - axisPosition=automatic：零轴按列内 min~max 比例偏移
      （正数跨度大 → 轴靠左；负数跨度大 → 轴靠右）
    """
    return (
        f'<x14:conditionalFormatting xmlns:xm="{XM_NS}">'
        f'<x14:cfRule type="dataBar" id="{guid}">'
        f'<x14:dataBar minLength="0" maxLength="100" border="1" gradient="1" '
        f'direction="leftToRight" axisPosition="automatic" '
        f'negativeBarColorSameAsPositive="0" '
        f'negativeBarBorderColorSameAsPositive="0">'
        f'<x14:cfvo type="autoMin"/>'
        f'<x14:cfvo type="autoMax"/>'
        f'<x14:borderColor rgb="{POS_COLOR}"/>'
        f'<x14:negativeFillColor rgb="{NEG_COLOR}"/>'
        f'<x14:negativeBorderColor rgb="{NEG_COLOR}"/>'
        f'<x14:axisColor rgb="{AXIS_COLOR}"/>'
        f"</x14:dataBar></x14:cfRule>"
        f"<xm:sqref>{sqref}</xm:sqref>"
        f"</x14:conditionalFormatting>"
    )


def _attach_x14_ext(xml: str, blocks: list[str]) -> str:
    if not blocks:
        return xml
    joined = "".join(blocks)
    # 已有同类 ext：追加到 conditionalFormattings 内
    if CF_EXT_URI in xml and "</x14:conditionalFormattings>" in xml:
        return xml.replace(
            "</x14:conditionalFormattings>",
            joined + "</x14:conditionalFormattings>",
            1,
        )

    x14_ext = (
        f'<ext uri="{CF_EXT_URI}" xmlns:x14="{X14_NS}">'
        f"<x14:conditionalFormattings>{joined}</x14:conditionalFormattings>"
        f"</ext>"
    )
    if re.search(r"</extLst>\s*</worksheet>", xml):
        return re.sub(
            r"</extLst>\s*</worksheet>",
            x14_ext + "</extLst></worksheet>",
            xml,
            count=1,
        )
    return xml.replace("</worksheet>", f"<extLst>{x14_ext}</extLst></worksheet>", 1)


def _inject_rule_ids_keep_color(xml: str, sqrefs: list[str]) -> tuple[str, list[tuple[str, str]]]:
    """注入 x14:id，修正 00→FF 前缀，但保留原填充色。"""
    rule_ids: list[tuple[str, str]] = []
    for sqref in sqrefs:
        guid = _guid()
        pattern = (
            rf'(<conditionalFormatting[^>]*sqref="{re.escape(sqref)}"[^>]*>\s*'
            rf'<cfRule[^>]*type="dataBar"[^>]*>)(.*?)(</cfRule>)'
        )

        def _repl(m: re.Match, gid: str = guid) -> str:
            head, body, tail = m.group(1), m.group(2), m.group(3)
            body = re.sub(
                r'(<color rgb=")00([0-9A-Fa-f]{6}")',
                rf"\1FF\2",
                body,
            )
            body = re.sub(r"<extLst>.*?</extLst>", "", body, flags=re.DOTALL)
            ext = (
                f'<extLst><ext uri="{ID_EXT_URI}" xmlns:x14="{X14_NS}">'
                f"<x14:id>{gid}</x14:id></ext></extLst>"
            )
            return head + body + ext + tail

        new_xml, n = re.subn(pattern, _repl, xml, count=1, flags=re.DOTALL)
        if n == 0:
            continue
        xml = new_xml
        rule_ids.append((sqref, guid))
    return xml, rule_ids


def _x14_gradient_block(sqref: str, guid: str, color: str) -> str:
    """单向数据条：渐变 + 保留指定颜色（红/蓝）。"""
    rgb = color if color.startswith("FF") or len(color) == 8 else f"FF{color}"
    return (
        f'<x14:conditionalFormatting xmlns:xm="{XM_NS}">'
        f'<x14:cfRule type="dataBar" id="{guid}">'
        f'<x14:dataBar minLength="0" maxLength="100" border="1" gradient="1" '
        f'direction="leftToRight" axisPosition="automatic">'
        f'<x14:cfvo type="autoMin"/>'
        f'<x14:cfvo type="autoMax"/>'
        f'<x14:borderColor rgb="{rgb}"/>'
        f"</x14:dataBar></x14:cfRule>"
        f"<xm:sqref>{sqref}</xm:sqref>"
        f"</x14:conditionalFormatting>"
    )


def patch_workbook_gradient_databars(
    xlsx_path: str | Path,
    sheet_ranges: dict[str, list[tuple[str, str]]],
) -> Path:
    """
    sheet_ranges: {sheet_name: [(sqref, rgb6_or_8), ...]}
    为单向数据条补渐变（不改正负轴逻辑）。
    """
    xlsx_path = Path(xlsx_path)
    if not sheet_ranges:
        return xlsx_path

    with TemporaryDirectory() as tmp:
        root = Path(tmp)
        with zipfile.ZipFile(xlsx_path, "r") as zin:
            zin.extractall(root)

        name_to_path = _sheet_paths_by_name(root)
        for sheet_name, ranges in sheet_ranges.items():
            path = name_to_path.get(sheet_name)
            if path is None or not ranges:
                continue
            xml = path.read_text(encoding="utf-8")
            sqrefs = [r for r, _ in ranges]
            color_by = {r: c for r, c in ranges}
            xml, rule_ids = _inject_rule_ids_keep_color(xml, sqrefs)
            blocks = [
                _x14_gradient_block(sqref, guid, color_by.get(sqref, POS_COLOR))
                for sqref, guid in rule_ids
            ]
            xml = _attach_x14_ext(xml, blocks)
            path.write_text(xml, encoding="utf-8")

        out_tmp = root / "_patched.xlsx"
        with zipfile.ZipFile(out_tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for f in root.rglob("*"):
                if f.is_file() and f != out_tmp:
                    zout.write(f, f.relative_to(root).as_posix())
        xlsx_path.write_bytes(out_tmp.read_bytes())

    return xlsx_path


def patch_workbook_bidirectional_databars(
    xlsx_path: str | Path,
    sheet_ranges: dict[str, list[str]],
) -> Path:
    """
    sheet_ranges: {sheet_name: ["E2:E31", "G2:G31", ...]}
    为正负混合列补双向轴 + 正红负绿。
    """
    xlsx_path = Path(xlsx_path)
    if not sheet_ranges:
        return xlsx_path

    with TemporaryDirectory() as tmp:
        root = Path(tmp)
        with zipfile.ZipFile(xlsx_path, "r") as zin:
            zin.extractall(root)

        name_to_path = _sheet_paths_by_name(root)
        for sheet_name, ranges in sheet_ranges.items():
            path = name_to_path.get(sheet_name)
            if path is None or not ranges:
                continue
            xml = path.read_text(encoding="utf-8")
            xml, rule_ids = _inject_rule_ids(xml, ranges)
            blocks = [_x14_block(sqref, guid) for sqref, guid in rule_ids]
            xml = _attach_x14_ext(xml, blocks)
            path.write_text(xml, encoding="utf-8")

        out_tmp = root / "_patched.xlsx"
        with zipfile.ZipFile(out_tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for f in root.rglob("*"):
                if f.is_file() and f != out_tmp:
                    zout.write(f, f.relative_to(root).as_posix())
        xlsx_path.write_bytes(out_tmp.read_bytes())

    return xlsx_path
