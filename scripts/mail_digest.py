#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
企业微信邮箱 · 收件摘要生成器

用途：拉取最近 N 天（默认 6 天）的邮件，按「未读 / 含附件 / 其他」分类，
      生成一份 Markdown 收件摘要，落到 office-agent/inbox/ 目录。

背景限制（实测确认，勿再试探）：
  1. 邮件读取范围硬限制为「当前时刻前后 7 天以内」，
     超出会返回 680220，所以本脚本的回看天数最大只能到 7。
  2. is_not_reminder 字段在该邮箱中恒为 false，即「重要邮件」筛选
     命中全部邮件、无区分度，故摘要不采用「重要」这一分类。

用法：
  python3 mail_digest.py            # 默认回看 6 天
  python3 mail_digest.py --days 3   # 回看 3 天
  python3 mail_digest.py --stdout   # 只打印，不写文件
"""

import argparse
import datetime
import json
import os
import subprocess
import sys

SKILL_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_DIR = os.path.join(SKILL_ROOT, "inbox")
NODE_BIN = "/Users/00000489/.workbuddy/binaries/node/versions/22.22.2-3/bin"
MAX_LOOKBACK_DAYS = 7


def build_env():
    env = os.environ.copy()
    env["PATH"] = NODE_BIN + os.pathsep + env.get("PATH", "")
    return env


def run_search(params, page_count=2):
    """调用 mail search，返回解析后的 dict；失败抛 RuntimeError。"""
    payload = json.dumps(params, ensure_ascii=False)
    cmd = ["wecom-cli", "mail", "search", "--json", payload, "--page-count", str(page_count)]
    proc = subprocess.run(cmd, capture_output=True, text=True, env=build_env())
    raw = (proc.stdout or "").strip()
    if not raw:
        raise RuntimeError("接口无输出：%s" % (proc.stderr or "").strip())

    data = None
    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        for line in reversed(raw.splitlines()):
            line = line.strip()
            if line.startswith("{"):
                try:
                    data = json.loads(line)
                    break
                except json.JSONDecodeError:
                    continue
    if data is None:
        raise RuntimeError("无法解析接口返回内容")

    if "error" in data and data["error"]:
        err = data["error"]
        raise RuntimeError(err.get("message") or "接口调用失败")
    return data


def collect(window_start, window_end):
    base = {"begin_time": window_start, "end_time": window_end, "limit": 100}
    all_mail = run_search(base, page_count=3).get("mails", [])
    with_attach = run_search(dict(base, has_attachments=True), page_count=3).get("mails", [])
    unread = run_search(dict(base, only_unread=True), page_count=3).get("mails", [])

    attach_ids = {m["mail_id"] for m in with_attach}
    unread_ids = {m["mail_id"] for m in unread}

    for m in all_mail:
        m["_has_attach"] = m["mail_id"] in attach_ids
        m["_unread"] = m["mail_id"] in unread_ids
    return all_mail


def fmt_time(s):
    return (s or "").strip()[:16]


def table(rows, cols):
    if not rows:
        return "_（无）_\n"
    out = ["| " + " | ".join(cols) + " |", "|" + "---|" * len(cols)]
    out += rows
    return "\n".join(out) + "\n"


def build_markdown(mails, window_start, window_end, generated_at):
    inbox = [m for m in mails if m.get("folder_name") != "已发送"]
    unread = [m for m in inbox if m.get("_unread")]
    attach = [m for m in inbox if not m.get("_unread") and m.get("_has_attach")]
    rest = [m for m in inbox if not m.get("_unread") and not m.get("_has_attach")]

    def rows(items, idx_start=1):
        r = []
        for i, m in enumerate(items, idx_start):
            subject = " ".join((m.get("subject") or "(无主题)").split())
            sender = (m.get("sender") or {}).get("name") or "-"
            folder = m.get("folder_name") or "-"
            tag = "含附件" if m.get("_has_attach") else "—"
            r.append("| %d | %s | %s | %s | %s |" % (i, sender, subject, fmt_time(m.get("send_time")), tag))
        return r

    cols = ["#", "发件人", "主题", "时间", "备注"]
    lines = []
    lines.append("# 收件摘要 · %s" % generated_at[:10])
    lines.append("")
    lines.append("**统计范围**：%s 至 %s（收件箱口径，已排除「已发送」）" % (window_start, window_end))
    lines.append("")
    lines.append("| 指标 | 数量 |")
    lines.append("|---|---|")
    lines.append("| 收件总数 | %d |" % len(inbox))
    lines.append("| 未读 | %d |" % len(unread))
    lines.append("| 含附件 | %d |" % len(attach))
    lines.append("| 其他 | %d |" % len(rest))
    lines.append("")
    lines.append("> 说明：邮件读取窗口受接口限制，最长仅可回看 7 天。")
    lines.append("")
    lines.append("## 一、未读邮件")
    lines.append("")
    lines.append(table(rows(unread), cols))
    lines.append("## 二、含附件邮件")
    lines.append("")
    lines.append(table(rows(attach, len(unread) + 1), cols))
    lines.append("## 三、其他邮件")
    lines.append("")
    lines.append(table(rows(rest, len(unread) + len(attach) + 1), cols))
    lines.append("---")
    lines.append("")
    lines.append("_本摘要由个人办公智能体生成，数据取自企业微信邮箱，仅本地留存。_")
    lines.append("")
    return "\n".join(lines)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--days", type=int, default=6, help="回看天数，最大 7")
    ap.add_argument("--stdout", action="store_true", help="只打印，不写文件")
    args = ap.parse_args()

    days = max(1, min(args.days, MAX_LOOKBACK_DAYS))
    now = datetime.datetime.now()
    start = (now - datetime.timedelta(days=days)).replace(hour=0, minute=0, second=0, microsecond=0)
    window_start = start.strftime("%Y-%m-%d %H:%M:%S")
    window_end = now.strftime("%Y-%m-%d %H:%M:%S")
    generated_at = now.strftime("%Y-%m-%d %H:%M")

    try:
        mails = collect(window_start, window_end)
    except RuntimeError as e:
        print("收件失败：%s" % e, file=sys.stderr)
        return 1

    md = build_markdown(mails, window_start, window_end, generated_at)

    if args.stdout:
        print(md)
        return 0

    os.makedirs(OUT_DIR, exist_ok=True)
    path = os.path.join(OUT_DIR, "收件摘要-%s.md" % now.strftime("%Y-%m-%d"))
    with open(path, "w", encoding="utf-8") as f:
        f.write(md)

    inbox = [m for m in mails if m.get("folder_name") != "已发送"]
    unread = len([m for m in inbox if m.get("_unread")])
    print("已生成：%s" % path)
    print("收件 %d 封，其中未读 %d 封" % (len(inbox), unread))
    return 0


if __name__ == "__main__":
    sys.exit(main())
