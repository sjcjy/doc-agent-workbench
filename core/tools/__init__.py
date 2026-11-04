from __future__ import annotations

"""工具注册表：LangChain 工具与 FastAPI / 规则路由共用同一份清单。"""

from core.tools import (
    compose_tool,
    convert_tool,
    edit_tool,
    format_tool,
    inspect_tool,
    structure_tool,
)

MODULES = [inspect_tool, compose_tool, edit_tool, format_tool, convert_tool, structure_tool]

TOOL_NAMES = [m.TOOL_NAME for m in MODULES]


def build_tools() -> list:
    """构造 LangChain 可用的工具列表。"""
    return [m.build_tool() for m in MODULES]


def tool_descriptions() -> list[dict]:
    """供界面展示的工具清单。"""
    return [{"name": m.TOOL_NAME, "desc": m.TOOL_DESC} for m in MODULES]


def dispatch(name: str, **kwargs):
    """按工具名直接调用核心实现，返回 ToolResult。FastAPI 与批处理脚本用它绕过模型。"""
    table = {
        inspect_tool.TOOL_NAME: inspect_tool.run,
        compose_tool.TOOL_NAME: compose_tool.run,
        edit_tool.TOOL_NAME: edit_tool.run,
        format_tool.TOOL_NAME: format_tool.run,
        convert_tool.TOOL_NAME: convert_tool.run,
        structure_tool.TOOL_NAME: structure_tool.run,
    }
    fn = table.get(name)
    if fn is None:
        raise KeyError(f"未知工具：{name}，可选：{', '.join(table)}")
    return fn(**kwargs)


__all__ = ["build_tools", "dispatch", "tool_descriptions", "TOOL_NAMES",
           "inspect_tool", "compose_tool", "edit_tool", "format_tool", "convert_tool",
           "structure_tool"]
