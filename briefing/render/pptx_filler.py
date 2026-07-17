"""将 payload 灌入 PPT 模版：整段替换文字 + 替换表图。"""

from __future__ import annotations

import shutil
from pathlib import Path
from typing import Any

import yaml
from pptx import Presentation
from pptx.util import Pt


def load_slide_map(path: str | Path) -> dict:
    with open(path, encoding="utf-8") as f:
        return yaml.safe_load(f)


def _find_shape(slide, name: str):
    for shape in slide.shapes:
        if shape.name == name:
            return shape
    raise KeyError(f"未找到形状: {name!r}")


def _set_paragraph_text(paragraph, text: str, *, font_size_pt: float | None = None) -> None:
    """整段替换：合并为单一 run，避免模版里数字被拆碎后难以维护。"""
    if not paragraph.runs:
        run = paragraph.add_run()
        run.text = text
        if font_size_pt:
            run.font.size = Pt(font_size_pt)
        return

    # 保留第一个 run 的字体属性，清空其余
    first = paragraph.runs[0]
    first.text = text
    if font_size_pt:
        first.font.size = Pt(font_size_pt)
    for run in paragraph.runs[1:]:
        run.text = ""


def replace_narrative_paragraphs(shape, paragraph_map: dict[int, str], font_size_pt: float = 12) -> None:
    """按段落索引写入文本。paragraph_map: {para_index: text}"""
    paragraphs = shape.text_frame.paragraphs
    for idx, text in paragraph_map.items():
        if idx >= len(paragraphs):
            continue
        if text is None:
            continue
        _set_paragraph_text(paragraphs[idx], text, font_size_pt=font_size_pt)


def replace_picture_fit(slide, shape_name: str, image_path: str | Path) -> None:
    """替换图片并保持原图宽高比，居中放入原图框（不拉伸）。"""
    from pptx.util import Emu

    shape = _find_shape(slide, shape_name)
    box_left, box_top = shape.left, shape.top
    box_w, box_h = shape.width, shape.height
    sp = shape._element
    sp.getparent().remove(sp)

    from PIL import Image

    with Image.open(image_path) as im:
        img_w, img_h = im.size
    img_aspect = img_w / img_h
    box_aspect = box_w / box_h

    if img_aspect > box_aspect:
        # 图更宽 → 以宽度为准
        new_w = box_w
        new_h = int(box_w / img_aspect)
    else:
        new_h = box_h
        new_w = int(box_h * img_aspect)

    left = box_left + (box_w - new_w) // 2
    top = box_top + (box_h - new_h) // 2
    slide.shapes.add_picture(str(image_path), left, top, width=new_w, height=new_h)


def replace_picture_inplace(slide, shape_name: str, image_path: str | Path, *, fit: bool = True) -> None:
    if fit:
        replace_picture_fit(slide, shape_name, image_path)
        return
    shape = _find_shape(slide, shape_name)
    left, top, width, height = shape.left, shape.top, shape.width, shape.height
    sp = shape._element
    sp.getparent().remove(sp)
    slide.shapes.add_picture(str(image_path), left, top, width=width, height=height)


def fill_pptx(
    template_path: str | Path,
    slide_map: dict,
    payload: dict[str, Any],
    output_path: str | Path,
    image_paths: dict[str, Path],
) -> Path:
    """
    payload 结构示例::

        {
          "overview": {
            "title": "...",
            "data_source_note": "...",
            ...
            "footnote": "..."
          },
          "total_ranking": { ... }
        }

    image_paths: {"category_summary": Path(...), "total_ranking_top30": Path(...)}
    """
    template_path = Path(template_path)
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy(template_path, output_path)

    prs = Presentation(str(output_path))

    for slide_cfg in slide_map.get("slides", []):
        slide = prs.slides[slide_cfg["index"]]
        section_id = slide_cfg["id"]
        section_payload = payload.get(section_id, {})

        # 叙事
        if "narrative_shape" in slide_cfg and "paragraphs" in slide_cfg:
            shape = _find_shape(slide, slide_cfg["narrative_shape"])
            para_map = {}
            for para_idx, key in slide_cfg["paragraphs"].items():
                # YAML may parse keys as int already
                idx = int(para_idx)
                if key in section_payload:
                    para_map[idx] = section_payload[key]
            replace_narrative_paragraphs(shape, para_map, font_size_pt=12)

        # 脚注
        if "footnote_shape" in slide_cfg:
            fk = slide_cfg.get("footnote_key", "footnote")
            if fk in section_payload:
                fn_shape = _find_shape(slide, slide_cfg["footnote_shape"])
                replace_narrative_paragraphs(fn_shape, {0: section_payload[fk]}, font_size_pt=10)

        # 表图
        table_key = slide_cfg.get("table_image")
        if table_key and table_key in image_paths and "picture_shape" in slide_cfg:
            replace_picture_inplace(slide, slide_cfg["picture_shape"], image_paths[table_key])

    prs.save(str(output_path))
    return output_path
