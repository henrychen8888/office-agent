#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""提取各版式占位符的完整字体规格（字号/加粗/颜色/字体）。"""

import glob
import json
import os
from lxml import etree

ROOT = "/tmp/pptx_x"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS = {"a": A, "p": P}

LAYOUT_LABEL = {
    "slideLayout1.xml": "1_首页", "slideLayout2.xml": "2_首页",
    "slideLayout3.xml": "2_目录页", "slideLayout10.xml": "2_目录页",
    "slideLayout4.xml": "3_节标题", "slideLayout11.xml": "3_节标题",
    "slideLayout12.xml": "4_内容页", "slideLayout5.xml": "6_结束页",
    "slideLayout14.xml": "6_结束页", "slideLayout7.xml": "标题和内容",
    "slideLayout16.xml": "标题和内容",
}


def rpr_spec(rpr):
    """从 defRPr / rPr 提取规格。"""
    if rpr is None:
        return None
    d = {}
    for k, key in (("sz", "size_pt"), ("b", "bold"), ("i", "italic"), ("spc", "letter_spacing")):
        v = rpr.get(k)
        if v is not None:
            if k == "sz":
                d[key] = int(v) / 100
            elif k in ("b", "i"):
                d[key] = v in ("1", "true")
            else:
                d[key] = v
    fill = rpr.find("a:solidFill/a:srgbClr", NS)
    if fill is not None:
        d["color"] = "#" + fill.get("val")
    for tag, key in (("latin", "latin_font"), ("ea", "ea_font")):
        n = rpr.find("a:" + tag, NS)
        if n is not None:
            d[key] = n.get("typeface")
    return d or None


def walk_levels(node):
    """收集 lstStyle 各级规格。"""
    out = {}
    if node is None:
        return out
    for child in node:
        tag = etree.QName(child).localname
        if tag.endswith("pPr"):
            rpr = child.find("a:defRPr", NS)
            spec = rpr_spec(rpr)
            if spec:
                marg = {}
                for k in ("marL", "indent"):
                    if child.get(k) is not None:
                        marg[k] = int(child.get(k)) / 360000
                if child.get("algn"):
                    marg["align"] = child.get("algn")
                if marg:
                    spec["layout_cm"] = marg
                out[tag] = spec
    return out


def main():
    result = {}
    for f in sorted(glob.glob(os.path.join(ROOT, "ppt/slideLayouts/slideLayout*.xml"))):
        base = os.path.basename(f)
        label = LAYOUT_LABEL.get(base)
        if not label:
            continue
        try:
            x = etree.parse(f).getroot()
        except Exception:
            continue
        shapes = []
        cSld = x.find(".//p:cSld", NS)
        if cSld is None:
            continue
        for sp in cSld.findall(".//p:sp", NS):
            nv = sp.find(".//p:cNvPr", NS)
            ph = sp.find(".//p:ph", NS)
            name = nv.get("name") if nv is not None else "?"
            entry = {"name": name, "ph_type": ph.get("type") if ph is not None else None}
            off = sp.find(".//a:off", NS)
            ext = sp.find(".//a:ext", NS)
            if off is not None and ext is not None:
                entry["pos_cm"] = [round(int(off.get("x")) / 360000, 2), round(int(off.get("y")) / 360000, 2)]
                entry["size_cm"] = [round(int(ext.get("cx")) / 360000, 2), round(int(ext.get("cy")) / 360000, 2)]
            bodyPr = sp.find(".//a:bodyPr", NS)
            if bodyPr is not None:
                anch = bodyPr.get("anchor")
                if anch:
                    entry["anchor"] = anch
            lst = walk_levels(sp.find(".//a:lstStyle", NS))
            if lst:
                entry["levels"] = lst
            # 直接挂在 run 上的规格
            run_rpr = sp.find(".//a:r/a:rPr", NS)
            rs = rpr_spec(run_rpr)
            if rs:
                entry["run_spec"] = rs
            shapes.append(entry)
        pics = []
        for pic in cSld.findall(".//p:pic", NS):
            nv = pic.find(".//p:cNvPr", NS)
            off = pic.find(".//a:off", NS)
            ext = pic.find(".//a:ext", NS)
            pics.append({
                "name": nv.get("name") if nv is not None else "?",
                "pos_cm": [round(int(off.get("x")) / 360000, 2), round(int(off.get("y")) / 360000, 2)] if off is not None else None,
                "size_cm": [round(int(ext.get("cx")) / 360000, 2), round(int(ext.get("cy")) / 360000, 2)] if ext is not None else None,
            })
        result.setdefault(label, {"file": base, "slides": []})
        result[label]["slides"].append({"shapes": shapes, "pictures": pics})

    out = "/Users/00000489/WorkBuddy/2026-09-23-08-28-19/office-agent/templates/_pptx_layouts.json"
    with open(out, "w", encoding="utf-8") as fp:
        json.dump(result, fp, ensure_ascii=False, indent=1)
    print("已写出:", out)
    for label, d in result.items():
        print("\n=== %s (%s) ===" % (label, d["file"]))
        for sl in d["slides"]:
            for s in sl["shapes"]:
                lv = s.get("levels", {}).get("lvl1pPr") or s.get("run_spec") or {}
                print("  %-16s %s lvl1: %spt b=%s %s" % (
                    s["name"], s.get("ph_type"), lv.get("size_pt"), lv.get("bold"), lv.get("latin_font")))
            for p in sl["pictures"]:
                print("  图片 %-12s %s %s" % (p["name"], p["pos_cm"], p["size_cm"]))


if __name__ == "__main__":
    main()
