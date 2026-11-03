#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""解析 PPTX 模板的格式要素：画布、母版、版式、主题、字体、颜色、逐页排版。"""

import json
import sys
from collections import Counter, OrderedDict

from pptx import Presentation
from pptx.util import Emu

PATH = "/Users/00000489/Desktop/月报模板v3.pptx"
EMU_PER_PT = 12700
EMU_PER_CM = 360000


def pt(v):
    try:
        return round(int(v) / EMU_PER_PT, 1)
    except Exception:
        return None


def cm(v):
    try:
        return round(int(v) / EMU_PER_CM, 2)
    except Exception:
        return None


def color_of(fmt):
    """从 ColorFormat 提取可读颜色。"""
    try:
        t = fmt.type
    except Exception:
        return None
    try:
        if t is None:
            return None
        if str(t).startswith("MSO_THEME_COLOR") or "THEME" in str(t):
            return "theme:%s" % fmt.theme_color
        if fmt.rgb is not None:
            return "#%s" % str(fmt.rgb)
    except Exception:
        pass
    return None


def run_info(r):
    f = r.font
    out = OrderedDict()
    out["text"] = r.text
    out["size_pt"] = pt(f.size) if f.size else None
    out["bold"] = bool(f.bold) if f.bold is not None else None
    out["italic"] = bool(f.italic) if f.italic is not None else None
    out["font"] = f.name
    c = color_of(f.color)
    if c:
        out["color"] = c
    return out


def para_info(p):
    d = OrderedDict()
    d["align"] = str(p.alignment) if p.alignment is not None else None
    d["level"] = p.level
    if p.line_spacing is not None:
        ls = p.line_spacing
        d["line_spacing"] = round(float(ls), 2) if not hasattr(ls, "pt") else pt(ls)
    if p.space_before is not None:
        d["space_before_pt"] = pt(p.space_before)
    if p.space_after is not None:
        d["space_after_pt"] = pt(p.space_after)
    runs = [run_info(r) for r in p.runs]
    if runs:
        d["runs"] = runs
    return d


def shape_info(sh, depth=0):
    d = OrderedDict()
    d["type"] = str(sh.shape_type)
    d["name"] = sh.name
    try:
        d["pos_cm"] = [cm(sh.left), cm(sh.top)]
        d["size_cm"] = [cm(sh.width), cm(sh.height)]
    except Exception:
        pass
    if sh.has_text_frame:
        tfs = []
        for p in sh.text_frame.paragraphs:
            if p.runs or (p.text and p.text.strip()):
                tfs.append(para_info(p))
        if tfs:
            d["paragraphs"] = tfs
    if sh.has_table:
        rows = []
        for r in sh.table.rows:
            rows.append([" ".join(rr.text.split()) for rr in [c for c in r.cells]])
        d["table"] = rows
    if sh.shape_type == 6 and depth < 2:  # GROUP
        d["children"] = [shape_info(s, depth + 1) for s in sh.shapes]
    return d


def theme_fonts(prs):
    """读取 theme1.xml 的字体与配色。"""
    out = {}
    try:
        part = prs.part.package.part_related_by  # noqa
    except Exception:
        pass
    try:
        from lxml import etree
        for rel in prs.part.rels.values():
            if "theme" in rel.reltype:
                xml = etree.fromstring(rel.target_part.blob)
                ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
                scheme = xml.find(".//a:fontScheme", ns)
                if scheme is not None:
                    major = scheme.find(".//a:majorFont/a:latin", ns)
                    minor = scheme.find(".//a:minorFont/a:latin", ns)
                    ea_major = scheme.find(".//a:majorFont/a:ea", ns)
                    ea_minor = scheme.find(".//a:minorFont/a:ea", ns)
                    out["major_latin"] = major.get("typeface") if major is not None else None
                    out["minor_latin"] = minor.get("typeface") if minor is not None else None
                    out["major_ea"] = ea_major.get("typeface") if ea_major is not None else None
                    out["minor_ea"] = ea_minor.get("typeface") if ea_minor is not None else None
                clr = xml.find(".//a:clrScheme", ns)
                if clr is not None:
                    colors = OrderedDict()
                    for child in clr:
                        tag = etree.QName(child).localname
                        v = child.find("a:srgbClr", ns)
                        s = child.find("a:sysClr", ns)
                        if v is not None:
                            colors[tag] = "#" + v.get("val")
                        elif s is not None:
                            colors[tag] = "sys:" + str(s.get("lastClr"))
                    out["color_scheme"] = colors
                break
    except Exception as e:
        out["error"] = str(e)
    return out


def main():
    prs = Presentation(PATH)
    result = OrderedDict()

    # 画布
    result["canvas"] = {
        "width_cm": cm(prs.slide_width),
        "height_cm": cm(prs.slide_height),
        "width_pt": pt(prs.slide_width),
        "height_pt": pt(prs.slide_height),
        "ratio": round(prs.slide_width / prs.slide_height, 4),
        "slide_count": len(prs.slides),
    }

    # 主题
    result["theme"] = theme_fonts(prs)

    # 母版
    masters = []
    for m in prs.slide_masters:
        md = OrderedDict()
        md["name"] = m.name
        md["layouts"] = [l.name for l in m.slide_layouts]
        md["shapes"] = [shape_info(s) for s in m.shapes]
        masters.append(md)
    result["masters"] = masters

    # 逐页
    slides = []
    for i, s in enumerate(prs.slides, 1):
        sd = OrderedDict()
        sd["index"] = i
        sd["layout"] = s.slide_layout.name
        sd["shapes"] = [shape_info(sh) for sh in s.shapes]
        if s.has_notes_slide:
            nt = s.notes_slide.notes_text_frame.text
            if nt.strip():
                sd["notes"] = nt
        slides.append(sd)
    result["slides"] = slides

    # 字体使用统计
    fonts = Counter()
    sizes = Counter()
    colors = Counter()
    for sl in slides:
        def walk(shapes):
            for sh in shapes:
                for p in sh.get("paragraphs", []):
                    for r in p.get("runs", []):
                        if r.get("font"):
                            fonts[r["font"]] += 1
                        if r.get("size_pt"):
                            sizes[r["size_pt"]] += 1
                        if r.get("color"):
                            colors[r["color"]] += 1
                walk(sh.get("children", []))
        walk(sl["shapes"])
    result["usage"] = {
        "fonts": dict(fonts.most_common()),
        "sizes_pt": dict(sizes.most_common()),
        "colors": dict(colors.most_common()),
    }

    out = "/Users/00000489/WorkBuddy/2026-09-23-08-28-19/office-agent/templates/_pptx_format_raw.json"
    with open(out, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=1)
    print("已写出:", out)
    print("画布:", result["canvas"])
    print("字体统计:", result["usage"]["fonts"])
    print("字号统计:", result["usage"]["sizes_pt"])
    print("颜色统计:", dict(list(result["usage"]["colors"].items())[:12]))


if __name__ == "__main__":
    sys.exit(main())
