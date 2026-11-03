#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
验证：以 月报模板v3.pptx 为基底，套用格式规范生成一份月报 PPT。

做法：复制模板 → 定位占位符 → 按 pptx_style 常量填充文字。
      保留全部母版、版式、背景素材，仅替换内容。
"""

import os
import shutil
import sys

from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.dml.color import RGBColor

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pptx_style as S  # noqa: E402

SRC = "/Users/00000489/Desktop/月报模板v3.pptx"
DST = "/Users/00000489/WorkBuddy/2026-09-23-08-28-19/office-agent/preview/月报_格式验证.pptx"


def set_ph_text(shape, text, kind=None, color=None):
    tf = shape.text_frame
    tf.text = text
    for p in tf.paragraphs:
        for r in p.runs:
            if kind:
                S.apply_run(r, kind, color)
    return shape


def find_shape(slide, name_contains):
    for sh in slide.shapes:
        if name_contains in sh.name:
            return sh
    return None


def main():
    os.makedirs(os.path.dirname(DST), exist_ok=True)
    shutil.copy(SRC, DST)

    prs = Presentation(DST)
    print("基底模板: %s" % os.path.basename(SRC))
    print("画布: %.2f x %.2f cm" % (prs.slide_width / 360000, prs.slide_height / 360000))
    print("页数: %d   母版: %d   版式总数: %d" % (
        len(prs.slides), len(prs.slide_masters),
        sum(len(m.slide_layouts) for m in prs.slide_masters)))
    print()

    # ── 第1页 封面：只替换可变信息，保留背景与装饰 ──
    s1 = prs.slides[0]
    tgt = {
        "标题": ("工作月报", "cover_title"),
        "内容占位符 3": (S.COVER["unit"], "cover_info"),
        "内容占位符 5": (S.COVER["presenter"], "cover_info"),
        "内容占位符 4": ("2026 年 09 月", "cover_info"),
    }
    print("第1页 封面:")
    for sh in s1.shapes:
        key = None
        for k in tgt:
            if sh.name.strip() == k:
                key = k
                break
        if key and sh.has_text_frame:
            txt, kind = tgt[key]
            before = sh.text_frame.text.replace("\n", " / ")
            set_ph_text(sh, txt, kind)
            print("  [%s] %s -> %s" % (sh.name, before[:24], txt))

    # ── 第3页 节标题：验证大号序号 ──
    s3 = prs.slides[2]
    print("\n第3页 节标题:")
    for sh in s3.shapes:
        if not sh.has_text_frame:
            continue
        if "标题" in sh.name:
            print("  [%s] -> %s" % (sh.name, S.SECTIONS[0][1]))
            set_ph_text(sh, S.SECTIONS[0][1], "section_title")
        elif "文本占位符" in sh.name:
            print("  [%s] -> %s (%.1fpt)" % (sh.name, S.SECTIONS[0][0], S.SIZE["section_number"]))
            set_ph_text(sh, S.SECTIONS[0][0], "section_number")

    # ── 第4页 内容页：填充指标表 ──
    s4 = prs.slides[3]
    print("\n第4页 内容页:")
    for sh in s4.shapes:
        if "标题" in sh.name and sh.has_text_frame:
            t = "数智技术研发处 – 2026 年 09 月 OGSM 指标达成状态"
            set_ph_text(sh, t, "content_title")
            print("  [%s] -> %s" % (sh.name, t))
        if sh.has_table:
            tb = sh.table
            print("  表格 %d行 x %d列" % (len(tb.rows), len(tb.columns)))
            # 表头
            for ci, col in enumerate(S.METRIC_COLUMNS):
                if ci < len(tb.columns):
                    cell = tb.cell(0, ci)
                    cell.text = col
                    for p in cell.text_frame.paragraphs:
                        for r in p.runs:
                            r.font.size = Pt(S.SIZE["table_value"])
                            r.font.bold = True
                            r.font.name = S.FONT_CN
            print("    表头: %s" % " / ".join(S.METRIC_COLUMNS))

    prs.save(DST)
    print("\n已生成: %s" % DST)
    print("大小: %.1f KB" % (os.path.getsize(DST) / 1024))

    # 复核
    chk = Presentation(DST)
    print("\n复核: 画布 %.2fx%.2f cm, 页数 %d, 母版 %d, 版式 %d" % (
        chk.slide_width / 360000, chk.slide_height / 360000,
        len(chk.slides), len(chk.slide_masters),
        sum(len(m.slide_layouts) for m in chk.slide_masters)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
