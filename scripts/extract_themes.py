#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""从解包目录读取主题配色与字体。"""

import glob
import os
import re
import json
from lxml import etree

ROOT = "/tmp/pptx_x"
NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}


def scheme(path):
    try:
        x = etree.parse(path).getroot()
    except Exception as e:
        return {"error": str(e)}
    out = {}
    fs = x.find(".//a:fontScheme", NS)
    if fs is not None:
        for tag in ("majorFont", "minorFont"):
            node = fs.find("a:" + tag, NS)
            if node is None:
                continue
            for sub in ("latin", "ea", "cs"):
                n = node.find("a:" + sub, NS)
                if n is not None:
                    out["%s_%s" % (tag, sub)] = n.get("typeface")
    cs = x.find(".//a:clrScheme", NS)
    if cs is not None:
        colors = {}
        for child in cs:
            tag = etree.QName(child).localname
            s = child.find("a:srgbClr", NS)
            y = child.find("a:sysClr", NS)
            if s is not None:
                colors[tag] = "#" + s.get("val")
            elif y is not None:
                colors[tag] = "#" + str(y.get("lastClr"))
        out["colors"] = colors
    return out


def main():
    files = sorted(glob.glob(os.path.join(ROOT, "ppt/theme/theme*.xml")))
    result = {}
    for f in files:
        result[os.path.basename(f)] = scheme(f)

    with open("/Users/00000489/WorkBuddy/2026-09-23-08-28-19/office-agent/templates/_pptx_themes.json", "w", encoding="utf-8") as fp:
        json.dump(result, fp, ensure_ascii=False, indent=1)

    # theme1 是主主题
    t1 = result.get("theme1.xml", {})
    print("=== 主主题 theme1 ===")
    print("字体:", {k: v for k, v in t1.items() if "Font" in k})
    print("配色:")
    for k, v in (t1.get("colors") or {}).items():
        print("  %-10s %s" % (k, v))
    print()
    print("=== 主题总数 ===", len(files))
    for name, d in result.items():
        c = (d.get("colors") or {}).get("accent1")
        print("  %-14s accent1=%s minor_latin=%s" % (name, c, d.get("minorFont_latin")))


if __name__ == "__main__":
    main()
