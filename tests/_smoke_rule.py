# -*- coding: utf-8 -*-
"""规则路由自查：验证自然语言能否被翻译成正确的工具调用。"""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from core import config
from core.agent import rule_agent

CASES = [
    "查看这份文档",
    "把「联系邮箱：project@example.com」改成「联系邮箱：team@example.com」",
    "把所有标题居中",
    "把正文字体改成宋体，字号设为小四",
    "行距设为 1.5 倍",
    "首行缩进 2 字符",
    "转成 PDF",
    "导出 Markdown",
    "在「三、下周计划」后面新增章节，标题为「四、风险提示」，内容为「需要关注接口稳定性」",
    "追加：本周五前完成联调",
    "写一份项目总结",
    "一键套用报告排版",
    "这句话不太可能被理解对吧",
]


def main():
    src = str(config.SAMPLE_DIR / "样例1_项目周报.docx")
    for q in CASES:
        plan = rule_agent.route(q, src)
        if "error" in plan:
            print(f"x {q[:34]:<36} -> ERROR {plan['error'][:50]}")
        else:
            print(f"v {q[:34]:<36} -> {plan['name']:<22} {str(plan['args'])[:90]}")


if __name__ == "__main__":
    main()
