#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
月报 PPTX 样式常量库（提取自 月报模板v3.pptx）

用途：生成新 PPT 时统一引用本模块的颜色/字体/字号/版式常量，
      确保所有产出与模板格式一致。

配套：以 `月报模板v3.pptx` 为基底生成时，直接复用其版式，
      再用本模块的常量覆盖文字规格。
"""

# ── 画布 ──────────────────────────────────────────────
SLIDE_W_CM = 25.40
SLIDE_H_CM = 14.29
SLIDE_W_EMU = 9144000
SLIDE_H_EMU = 5143500
RATIO = 16 / 9

# ── 主题色板（theme1）────────────────────────────────
COLORS = {
    "dk1": "000000",
    "lt1": "FFFFFF",
    "dk2": "0E2841",
    "lt2": "E8E8E8",
    "accent1": "156082",
    "accent2": "E97132",
    "accent3": "196B24",
    "accent4": "0F9ED5",
    "accent5": "A02B93",
    "accent6": "4EA72E",
    "hlink": "467886",
    "folHlink": "96607D",
}

# 实测用色
TEXT_MAIN = "1A1A1A"      # 正文
TEXT_TITLE = "333333"     # 首页大标题
HIGHLIGHT_YELLOW = "FFFF00"  # 达标率强调
HIGHLIGHT_BLUE = "0432FF"    # 关键数值 / 分割符

# ── 字体 ──────────────────────────────────────────────
# 实测：模板版式层用中文名「微软雅黑」，内容页正文层用英文名「Microsoft YaHei」。
# 两者在渲染上等价（同字体不同写法），但为最大程度贴合模板，按层区分：
FONT_CN = "Microsoft YaHei"   # 内容层（对标模板内容页 P4/P6/P7 的写法）
FONT_LAYOUT = "微软雅黑"       # 版式层（对标模板 slideLayout 的写法）
FONT_END = "Arial"            # 仅结束页
FONT_LATIN = "Microsoft YaHei"  # 西文同用

# ── 字号（pt）─────────────────────────────────────────
# 模板实测：内容页表格内文字统一 12pt 粗，进度条数值 10.5pt 粗。
# 为对齐模板的字号阶梯，内容页正文统一收敛到 12 / 10.5 / 10.5 三档，
# 不再散用 9~11.5 的碎档位。
SIZE = {
    "cover_title": 27.0,      # 首页主标题（粗）· 模板实测 2700
    "cover_subtitle": 18.0,   # 首页副标题 · 模板实测 1800
    "cover_info": 15.0,       # 首页单位/汇报人/年月 · 模板实测 1500
    "toc_item": 18.0,         # 目录项（粗）· 模板实测 1800
    "section_title": 27.0,    # 节标题（粗）· 模板实测 2700
    "section_number": 103.5,  # 节标题大号序号（粗）· 模板实测 10350
    "content_title": 19.5,    # 内容页标题（粗）· 模板实测 1950
    "content_body": 12.0,     # 内容页正文（粗）· 模板表格实测 1200
    "table_value_bold": 12.0,  # 表格强调数字（粗）· 模板实测 1200
    "table_value": 10.5,       # 表格普通数字（粗）· 模板实测 1050
    "card_title": 12.0,        # 卡片标题（粗）
    "card_desc": 10.5,         # 卡片说明
    "note": 10.5,              # 注释/脚注
    "end_text": 24.0,          # 结束页文字（粗）· 模板实测 2400
    "part_decoration": 80.0,   # Part 装饰字 · 模板实测 8000
}

# ── 版式名（须与模板中的名称一致）─────────────────────
LAYOUT = {
    "cover": "2_首页",
    "toc": "2_目录页",
    "section": "3_节标题",
    "content": "标题和内容",
    "end": "6_结束页",
}

# ── 关键位置（cm）────────────────────────────────────
POS = {
    "cover_title": (1.20, 3.01),
    "cover_subtitle": (1.20, 4.80),
    "cover_unit": (2.35, 9.32),
    "cover_presenter": (2.35, 10.55),
    "cover_date": (2.35, 11.79),
    "toc_items": [(13.84, 2.63), (13.84, 4.60), (13.84, 6.50), (13.84, 8.40)],
    "section_title": (5.89, 5.83),
    "section_number": (2.98, 5.16),
    "content_title": (4.32, 0.42),
    "content_table": (0.92, 2.38),
    "content_card": (0.92, 7.14),
    "end_text": (10.00, 5.50),
}

# ── 垂直网格（cm）· 实测自原模板 ─────────────────────
# 模板规律：标题 y=0.42 h=0.98；正文固定 y=2.37 起；页脚 y=13.24；
#          底部留白收敛在 1.3~2.4，版面饱满。
GRID = {
    "title_y": 0.42,          # 内容页标题顶
    "title_h": 0.98,          # 内容页标题高
    "body_y": 2.38,           # 正文区起始（严格跟随模板）
    "body_left": 0.92,        # 左边距
    "body_w": 24.08,          # 正文区宽度
    "footer_y": 13.24,        # 页脚三件套
    "body_bottom": 12.6,      # 正文区安全下界
    "max_bottom_gap": 2.4,    # 底部留白上限（超出即需加高行距或均分）
}

# ── 尺寸（cm）────────────────────────────────────────
SIZES_CM = {
    "content_table_w": 24.08,   # 内容页主表格宽度（固定）
    "content_card_w": 23.44,    # 圆角矩形容器宽
    "progress_bar_h": 0.45,     # 进度条高
    "toc_item_w": 6.60,
    "toc_item_h": 0.90,
    "gap_card": 0.28,           # 卡片间距
    "gap_section": 0.55,        # 区块间距
}

# ── 封面固定信息 ──────────────────────────────────────
COVER = {
    "unit": "数智技术研发处",
    "presenter": "吴永生",
}

# ── 章节结构（固定三章）───────────────────────────────
SECTIONS = [
    ("01", "OGSM指标达成状态"),
    ("02", "指标达成相关重点工作"),
    ("03", "其他关注事项"),
]

# ── 指标表列定义（固定 7 列）──────────────────────────
METRIC_COLUMNS = ["指标名称", "目标值", "单位", "指标值", "得分", "占比", "加权得分"]

# ── 改善说明表列定义（3 列）───────────────────────────
IMPROVE_COLUMNS = ["未达标原因分类", "原因说明", "改善说明"]

# ── 页脚位置（cm）────────────────────────────────────
FOOTER = {
    "date": (1.75, 13.24),
    "footer": (8.41, 13.24),
    "page_no": (17.94, 13.24),
}


def vgrid_rows(n, top=None, bottom=None, gap=None, min_h=0.8):
    """把正文区按 n 行均分，返回 [(y, h), ...]。

    用于让内容铺满到模板的密集度——模板正文固定 y=2.38 起、下界约 12.6，
    底部留白收敛在 1.3~2.4。行数少时自动加高单行，避免「上挤下空」。
    """
    top = GRID["body_y"] if top is None else top
    bottom = GRID["body_bottom"] if bottom is None else bottom
    gap = SIZES_CM["gap_card"] if gap is None else gap
    avail = bottom - top - gap * (n - 1)
    h = max(min_h, avail / n)
    # 若加高后超出下界，回退为紧凑排布
    if top + h * n + gap * (n - 1) > bottom + 0.01:
        h = max(min_h, (bottom - top - gap * (n - 1)) / n)
    return [(top + i * (h + gap), h) for i in range(n)]


def content_title_y():
    """内容页标题的 (x, y, w, h)。"""
    return (4.32, GRID["title_y"], 20.67, GRID["title_h"])


def apply_run(run, kind, color=None):
    from pptx.util import Pt
    from pptx.dml.color import RGBColor

    spec_size = {
        "cover_title": SIZE["cover_title"],
        "cover_subtitle": SIZE["cover_subtitle"],
        "cover_info": SIZE["cover_info"],
        "toc_item": SIZE["toc_item"],
        "section_title": SIZE["section_title"],
        "section_number": SIZE["section_number"],
        "content_title": SIZE["content_title"],
        "content_body": SIZE["content_body"],
        "table_value_bold": SIZE["table_value_bold"],
        "table_value": SIZE["table_value"],
        "end_text": SIZE["end_text"],
    }
    bold_kinds = {
        "cover_title", "toc_item", "section_title", "section_number",
        "content_title", "table_value_bold", "table_value", "end_text",
    }

    if kind in spec_size:
        run.font.size = Pt(spec_size[kind])
    run.font.bold = kind in bold_kinds

    if kind == "end_text":
        run.font.name = FONT_END
    else:
        run.font.name = FONT_CN

    if color:
        run.font.color.rgb = RGBColor.from_string(color.lstrip("#"))
    elif kind == "cover_title":
        run.font.color.rgb = RGBColor.from_string(TEXT_TITLE)
    elif kind == "table_value_bold":
        run.font.color.rgb = RGBColor.from_string(HIGHLIGHT_YELLOW)
    elif kind in ("content_body", "table_value"):
        run.font.color.rgb = RGBColor.from_string(TEXT_MAIN)


if __name__ == "__main__":
    print("画布: %.2f x %.2f cm  (16:9)" % (SLIDE_W_CM, SLIDE_H_CM))
    print("字体: %s / 结束页 %s" % (FONT_CN, FONT_END))
    print("字号档位: %d 个" % len(SIZE))
    print("版式: %s" % "、".join(LAYOUT.values()))
    print("章节: %s" % " | ".join("%s %s" % s for s in SECTIONS))
    print("指标表列: %s" % " / ".join(METRIC_COLUMNS))
