from __future__ import annotations

"""文档工作台智能体 - 结果与异常的统一数据结构。

设计原则（源自计划书 5.x）：
- 所有工具返回同一种结构，界面层无需为每个工具写特判；
- 失败如实返回，绝不伪造成功。任何异常都落到 ok=False + 明确的 code；
- 工具不负责把结果打印出来，只负责把「发生了什么」说清楚。
"""

from dataclasses import dataclass, field, asdict
from typing import Any

# 统一的错误码。界面层用它给出不同的提示语气与建议。
CODE_OK = "OK"
CODE_INVALID_INPUT = "INVALID_INPUT"          # 空输入 / 参数缺失
CODE_FILE_NOT_FOUND = "FILE_NOT_FOUND"
CODE_UNSUPPORTED_FORMAT = "UNSUPPORTED_FORMAT"
CODE_FILE_CORRUPTED = "FILE_CORRUPTED"        # 打不开的 docx（实为 zip 损坏）
CODE_NOT_FOUND_IN_DOC = "NOT_FOUND_IN_DOC"    # 要改的内容在文档里找不到
CODE_NO_CHANGE = "NO_CHANGE"                  # 执行了但没有产生任何变化
CODE_RENDER_FAILED = "RENDER_FAILED"          # PDF 渲染后端缺失或失败
CODE_MODEL_FAILED = "MODEL_FAILED"            # 大模型调用失败（繁忙 / Key 无效 / 额度）
CODE_SAFETY_BLOCKED = "SAFETY_BLOCKED"        # 路径越界 / 试图写回原文件
CODE_INTERNAL_ERROR = "INTERNAL_ERROR"

_MESSAGE_BY_CODE = {
    CODE_INVALID_INPUT: "输入不合法或缺少必要参数",
    CODE_FILE_NOT_FOUND: "文件不存在",
    CODE_UNSUPPORTED_FORMAT: "不支持的文件格式",
    CODE_FILE_CORRUPTED: "文件已损坏，无法解析为 DOCX",
    CODE_NOT_FOUND_IN_DOC: "在文档中未找到目标内容",
    CODE_NO_CHANGE: "操作未产生任何改动",
    CODE_RENDER_FAILED: "渲染后端不可用，转换未生成文件",
    CODE_MODEL_FAILED: "大模型调用失败",
    CODE_SAFETY_BLOCKED: "操作被安全策略拦截",
    CODE_INTERNAL_ERROR: "内部错误",
}


def code_text(code: str) -> str:
    """把错误码翻译成中文短语，用于界面展示。"""
    return _MESSAGE_BY_CODE.get(code, "未知状态")


@dataclass
class ToolResult:
    """所有工具的统一返回值。"""

    ok: bool
    tool: str                      # inspect / compose / edit / format / convert
    action: str                    # 具体动作，如 replace_text / set_font / to_pdf
    code: str = CODE_OK
    message: str = ""
    output_file: str | None = None  # 输出文件的绝对路径（一定是新文件，绝不覆盖输入）
    changes: list[dict] = field(default_factory=list)   # 变动摘要，供 diff 展示
    data: dict = field(default_factory=dict)            # 工具各自的详细数据
    elapsed_ms: int = 0

    def __post_init__(self) -> None:
        # 失败时若未给 message，用错误码兜底，避免界面出现空白提示
        if not self.message:
            self.message = code_text(self.code) if not self.ok else "操作成功"

    def to_dict(self) -> dict[str, Any]:
        d = asdict(self)
        d["code_text"] = code_text(self.code)
        return d

    # ---------- 便捷构造器 ----------
    @classmethod
    def fail(cls, tool: str, action: str, code: str, message: str = "", **kw) -> "ToolResult":
        return cls(ok=False, tool=tool, action=action, code=code, message=message, **kw)

    @classmethod
    def ok_result(cls, tool: str, action: str, message: str = "操作成功", **kw) -> "ToolResult":
        return cls(ok=True, tool=tool, action=action, code=CODE_OK, message=message, **kw)
