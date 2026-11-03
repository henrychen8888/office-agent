#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""PPTX 模板与生成版的全细节对比：结构/母版/版式/主题/字体/颜色/排版/功能开关。"""

import glob
import io
import os
import re
import sys

SRC = "/tmp/cmp/src"
GEN = "/tmp/cmp/gen"
NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

from lxml import etree


def read(p):
    return io.open(p, encoding="utf-8").read()


def grp(title):
    print()
    print("=" * 74)
    print(title)
    print("=" * 74)


def same(a, b):
    return "✓ 一致" if a == b else "✗ 差异"


def cmp_val(label, va, vb, width=52):
    ok = same(va, vb)
    print("  %-22s %s" % (label, ok))
    if va != vb:
        print("      模板: %s" % str(va)[:width * 2])
        print("      生成: %s" % str(vb)[:width * 2])


# ── 1. 包结构 ──
def check_structure():
    grp("对比 1：包结构（母版 / 版式 / 主题 / 媒体 / 备注）")

    def count(pat):
        return len(glob.glob(os.path.join(SRC, pat))), len(glob.glob(os.path.join(GEN, pat)))

    for label, pat in [
        ("母版 slideMaster", "ppt/slideMasters/*.xml"),
        ("版式 slideLayout", "ppt/slideLayouts/*.xml"),
        ("主题 theme", "ppt/theme/*.xml"),
        ("媒体 media", "ppt/media/*"),
        ("备注页 notesSlide", "ppt/notesSlides/notesSlide*.xml"),
        ("备注母版 notesMaster", "ppt/notesMasters/*.xml"),
    ]:
        a, b = count(pat)
        cmp_val(label, a, b)

    # 版式名逐一对齐
    def layout_names(root):
        out = {}
        for f in sorted(glob.glob(os.path.join(root, "ppt/slideLayouts/slideLayout*.xml"))):
            m = re.search(r'name="([^"]*)"', read(f))
            out[os.path.basename(f)] = m.group(1) if m else ""
        return out

    la, lb = layout_names(SRC), layout_names(GEN)
    diff = [k for k in la if la.get(k) != lb.get(k)]
    print()
    print("  版式名对齐: %s" % ("✓ 全部一致（%d 个）" % len(la) if not diff else "✗ %d 个不同" % len(diff)))
    for k in diff[:6]:
        print("      %s 模板=%s 生成=%s" % (k, la.get(k), lb.get(k)))


# ── 2. 母版内容 ──
def check_masters():
    grp("对比 2：母版内容（占位符 / 尺寸 / 位置）")

    def master_sig(p):
        try:
            x = etree.parse(p).getroot()
        except Exception as e:
            return [("解析失败", str(e))]
        out = []
        cs = x.find(".//{http://schemas.openxmlformats.org/presentationml/2006/main}cSld")
        if cs is None:
            return out
        for sp in cs.findall(".//{http://schemas.openxmlformats.org/presentationml/2006/main}sp"):
            nv = sp.find(".//{http://schemas.openxmlformats.org/presentationml/2006/main}cNvPr")
            off = sp.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}off")
            ext = sp.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}ext")
            ph = sp.find(".//{http://schemas.openxmlformats.org/presentationml/2006/main}ph")
            pos = ""
            if off is not None and ext is not None:
                pos = "(%s,%s)x(%s,%s)" % (off.get("x"), off.get("y"), ext.get("cx"), ext.get("cy"))
            out.append((nv.get("name") if nv is not None else "?",
                        ph.get("type") if ph is not None else "-", pos))
        return out

    files = sorted(os.path.basename(f) for f in glob.glob(os.path.join(SRC, "ppt/slideMasters/*.xml")))
    bad = 0
    for f in files:
        a = master_sig(os.path.join(SRC, "ppt/slideMasters", f))
        b = master_sig(os.path.join(GEN, "ppt/slideMasters", f))
        if a != b:
            bad += 1
            print("  ✗ %s 结构不同" % f)
            for i in range(max(len(a), len(b))):
                va = a[i] if i < len(a) else None
                vb = b[i] if i < len(b) else None
                if va != vb:
                    print("      %s  vs  %s" % (va, vb))
    print("  母版占位符结构: %s" % ("✓ 全部一致（%d 个母版）" % len(files) if not bad else "✗ %d 个不同" % bad))


# ── 3. 主题 ──
def check_themes():
    grp("对比 3：主题（配色方案 / 字体方案）")
    A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

    def sig(p):
        try:
            x = etree.parse(p).getroot()
        except Exception:
            return None
        d = {}
        cs = x.find(".//%sclrScheme" % A)
        if cs is not None:
            for c in cs:
                t = etree.QName(c).localname
                s = c.find("%ssrgbClr" % A)
                y = c.find("%ssysClr" % A)
                d[t] = "#" + s.get("val") if s is not None else ("sys:" + str(y.get("lastClr")) if y is not None else "?")
        fs = x.find(".//%sfontScheme" % A)
        if fs is not None:
            for tag in ("majorFont", "minorFont"):
                n = fs.find("%s%s" % (A, tag))
                if n is None:
                    continue
                for sub in ("latin", "ea"):
                    e = n.find("%s%s" % (A, sub))
                    if e is not None:
                        d["%s_%s" % (tag, sub)] = e.get("typeface")
        return d

    files = sorted(os.path.basename(f) for f in glob.glob(os.path.join(SRC, "ppt/theme/theme*.xml")))
    bad = 0
    for f in files:
        a = sig(os.path.join(SRC, "ppt/theme", f))
        b = sig(os.path.join(GEN, "ppt/theme", f))
        if a != b:
            bad += 1
            print("  ✗ %s 不同" % f)
            for k in set(list(a or {}) + list(b or {})):
                if (a or {}).get(k) != (b or {}).get(k):
                    print("      %s: %s vs %s" % (k, (a or {}).get(k), (b or {}).get(k)))
    print("  主题配色与字体: %s" % ("✓ 全部一致（%d 个主题）" % len(files) if not bad else "✗ %d 个不同" % bad))


# ── 4. 版式细节 ──
def check_layouts():
    grp("对比 4：版式细节（占位符 / 背景 / 继承关系）")
    P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"

    def sig(p):
        try:
            x = etree.parse(p).getroot()
        except Exception:
            return {}
        d = {"bg": None, "shapes": [], "pics": []}
        bg = x.find(".//%scSld/%sbg" % (P, P))
        d["bg"] = etree.tostring(bg)[:60].decode() if bg is not None else "无"
        cs = x.find(".//%scSld" % P)
        if cs is None:
            return d
        for sp in cs.findall(".//%ssp" % P):
            nv = sp.find(".//%scNvPr" % P)
            d["shapes"].append(nv.get("name") if nv is not None else "?")
        for pic in cs.findall(".//%spic" % P):
            nv = pic.find(".//%scNvPr" % P)
            d["pics"].append(nv.get("name") if nv is not None else "?")
        return d

    files = sorted(os.path.basename(f) for f in glob.glob(os.path.join(SRC, "ppt/slideLayouts/*.xml")))
    bad = 0
    for f in files:
        a = sig(os.path.join(SRC, "ppt/slideLayouts", f))
        b = sig(os.path.join(GEN, "ppt/slideLayouts", f))
        if a != b:
            bad += 1
            print("  ✗ %s" % f)
            if a.get("bg") != b.get("bg"):
                print("      背景差异")
            if a.get("shapes") != b.get("shapes"):
                print("      图形: %s vs %s" % (a.get("shapes"), b.get("shapes")))
            if a.get("pics") != b.get("pics"):
                print("      图片: %s vs %s" % (a.get("pics"), b.get("pics")))
    print("  版式结构: %s" % ("✓ 全部一致（%d 个版式）" % len(files) if not bad else "✗ %d 个不同" % bad))


# ── 5. 应用了版式的内容页 ──
def check_slide_layout_usage():
    grp("对比 5：内容页所用版式（继承链路）")
    from pptx import Presentation
    A = Presentation("/Users/00000489/Desktop/月报模板v3.pptx")
    B = Presentation("/Users/00000489/WorkBuddy/2026-09-23-08-28-19/office-agent/preview/资产轮巡制度可行性方案.pptx")
    print("  模板 10 页:")
    for i, s in enumerate(A.slides, 1):
        print("    P%-2d %s" % (i, s.slide_layout.name))
    print("  生成 17 页:")
    for i, s in enumerate(B.slides, 1):
        print("    P%-2d %s" % (i, s.slide_layout.name))

    ta = {s.slide_layout.name for s in A.slides}
    tb = {s.slide_layout.name for s in B.slides}
    print()
    print("  模板用到的版式: %s" % "、".join(sorted(ta)))
    print("  生成用到的版式: %s" % "、".join(sorted(tb)))
    extra = tb - ta
    print("  额外使用: %s" % ("无 ✓" if not extra else "、".join(extra)))


# ── 6. 字体与颜色 ──
def check_font_color():
    grp("对比 6：字体与字体颜色（逐 run 统计）")
    from collections import Counter
    from pptx import Presentation

    def stat(path):
        p = Presentation(path)
        fonts, colors, sizes, bolds = Counter(), Counter(), Counter(), Counter()
        for s in p.slides:
            for sh in s.shapes:
                if not sh.has_text_frame:
                    continue
                for para in sh.text_frame.paragraphs:
                    for r in para.runs:
                        if not r.text.strip():
                            continue
                        fonts[r.font.name or "(继承)"] += 1
                        if r.font.size:
                            sizes[r.font.size.pt] += 1
                        bolds[bool(r.font.bold)] += 1
                        try:
                            if r.font.color and r.font.color.type is not None and r.font.color.rgb:
                                colors["#" + str(r.font.color.rgb)] += 1
                        except Exception:
                            pass
        return fonts, colors, sizes, bolds

    fa, ca, sa, ba = stat("/Users/00000489/Desktop/月报模板v3.pptx")
    fb, cb, sb, bb = stat("/Users/00000489/WorkBuddy/2026-09-23-08-28-19/office-agent/preview/资产轮巡制度可行性方案.pptx")

    print("  【字体】")
    print("    模板: %s" % dict(fa))
    print("    生成: %s" % dict(fb))
    extra_fonts = set(fb) - set(fa)
    print("    生成版新增字体: %s" % ("无 ✓" if not extra_fonts else "、".join(extra_fonts)))

    print()
    print("  【字号】")
    print("    模板: %s" % dict(sorted(sa.items(), key=lambda x: -x[0])))
    print("    生成: %s" % dict(sorted(sb.items(), key=lambda x: -x[0])))
    extra_sizes = set(sb) - set(sa)
    print("    生成版新增字号: %s" % ("无 ✓" if not extra_sizes else sorted(extra_sizes)))

    print()
    print("  【字色】")
    print("    模板: %s" % dict(ca.most_common(8)))
    print("    生成: %s" % dict(cb.most_common(8)))
    extra_colors = set(cb) - set(ca)
    print("    生成版新增字色: %s" % ("无 ✓" if not extra_colors else sorted(extra_colors)))

    print()
    print("  【加粗分布】")
    print("    模板: %s" % dict(ba))
    print("    生成: %s" % dict(bb))


# ── 7. 功能开关 ──
def check_features():
    grp("对比 7：功能开关（母版 / 版式的行为设置）")
    P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"

    def feat(p):
        try:
            x = etree.parse(p).getroot()
        except Exception:
            return {}
        d = {}
        for tag in ("hf", "tbg", "txStyles", "sldLayoutIdLst", "clrMap"):
            n = x.find(".//%s%s" % (P, tag))
            if n is not None:
                d[tag] = etree.tostring(n)[:100].decode()
        return d

    for f in sorted(os.path.basename(x) for x in glob.glob(os.path.join(SRC, "ppt/slideMasters/*.xml")))[:3]:
        a = feat(os.path.join(SRC, "ppt/slideMasters", f))
        b = feat(os.path.join(GEN, "ppt/slideMasters", f))
        print("  %s" % f)
        for k in sorted(set(list(a) + list(b))):
            v = same(a.get(k), b.get(k))
            if "✗" in v:
                print("    %-14s %s" % (k, v))
        if all(a.get(k) == b.get(k) for k in set(list(a) + list(b))):
            print("    ✓ 全部一致")


def main():
    check_structure()
    check_masters()
    check_themes()
    check_layouts()
    check_slide_layout_usage()
    check_font_color()
    check_features()
    print()
    print("=" * 74)
    print("对比完成")
    print("=" * 74)


if __name__ == "__main__":
    main()
