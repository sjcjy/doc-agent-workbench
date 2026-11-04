# -*- coding: utf-8 -*-
"""工具层冒烟自查：不依赖大模型，直接调用五个工具跑一遍。"""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from core import config
from core.tools import compose_tool, convert_tool, edit_tool, format_tool, inspect_tool


def main():
    src = str(config.SAMPLE_DIR / "样例1_项目周报.docx")
    sid = "smoke"
    print("源文件:", src)

    r = inspect_tool.run(path=src, session_id=sid)
    print("1) inspect      :", r.ok, r.code, "|", r.message[:70])

    r2 = compose_tool.run(
        topic="季度总结", title="季度工作总结", session_id=sid,
        outline="# 季度工作总结\n\n## 主要进展\n\n- 完成工具层五个模块\n- 完成界面与接口\n\n正文示例段落，用于验证排版效果。\n",
    )
    print("2) compose      :", r2.ok, r2.code, "|", r2.message[:70])

    r3 = edit_tool.run(path=src, action="replace",
                       old_text="联系邮箱：project@example.com",
                       new_text="联系邮箱：team@example.com", session_id=sid)
    print("3) edit/replace :", r3.ok, r3.code, "|", r3.message[:70])

    r4 = format_tool.run(path=src, action="align", target="headings",
                         alignment="center", session_id=sid)
    print("4) format/align :", r4.ok, r4.code, "|", r4.message[:70], "| 变更段数:", len(r4.changes))

    r5 = convert_tool.run(path=src, target="md", session_id=sid)
    print("5) convert/md   :", r5.ok, r5.code, "|", r5.message[:70])

    r6 = convert_tool.run(path=src, target="pdf", session_id=sid)
    print("6) convert/pdf  :", r6.ok, r6.code, "|", r6.message[:140])
    if r6.ok:
        print("   后端:", r6.data.get("backend"), "| 输出:", Path(r6.output_file).name)


if __name__ == "__main__":
    main()
