from __future__ import annotations

"""会话上下文。

用 ContextVar 传递 session_id，工具函数就不必在多传一个参数，
LangChain 也不会把它当成模型可见的工具参数暴露出去。
"""

from contextvars import ContextVar

_SESSION: ContextVar[str] = ContextVar("doc_session_id", default="")


def set_session_id(value: str) -> None:
    _SESSION.set(value or "")


def get_session_id() -> str:
    return _SESSION.get()


def new(value: str) -> str:
    set_session_id(value)
    return value
