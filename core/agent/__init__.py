from __future__ import annotations

"""Agent 统一入口：上层界面（Streamlit / FastAPI / 未来的小程序）只跟这里对话。

策略：
- AGENT_BACKEND=llm   → 只走大模型
- AGENT_BACKEND=rule  → 只走规则路由
- AGENT_BACKEND=auto  → 优先大模型，失败（无 Key / 超时 / 鉴权失败）自动回退规则路由
"""

import time

from core import config
from core.ctx import get_session_id
from core.logger import log_call


def execute(query: str, file_path: str = "", session_id: str = "", backend: str | None = None) -> dict:
    """执行一次用户任务。返回结构：
    {ok, backend, reply, steps, outputs, error, suggestion, elapsed_ms}
    """
    sid = session_id or get_session_id()
    mode = (backend or config.AGENT_BACKEND or "auto").lower()

    from core.agent import llm_agent, rule_agent

    attempted: list[str] = []

    # file_path 里用 | 串了多份文件时，直接走规则路由：大模型拿到的只是一个拼接字符串，
    # 容易把 "a|b" 整个当成一个路径塞进 path 参数，反而必错。这类操作本身就是确定性的。
    if "|" in str(file_path or ""):
        result = rule_agent.run(query, file_path, sid)
        result["backend"] = "rule(multi-file)"
        return result

    if mode in ("auto", "llm"):
        try:
            return llm_agent.run(query, file_path, sid)
        except Exception as e:
            attempted.append(str(e)[:200])
            if mode == "llm":
                log_call(sid, "agent", "fallback", "skip", {"query": query[:120]}, False,
                         "MODEL_FAILED", f"指定 llm 但调用失败：{e}", backend="llm")
                return {
                    "ok": False, "backend": "llm(failed)", "reply": "",
                    "steps": [], "outputs": [],
                    "error": f"大模型调用失败：{e}",
                    "suggestion": "检查 .env 中的 ZHIPU_API_KEY 是否正确、额度是否充足，或改用规则路由模式",
                    "elapsed_ms": 0,
                }

    result = rule_agent.run(query, file_path, sid)
    if attempted:
        result["suggestion"] = (result.get("suggestion", "") or "") + \
                               f"（大模型不可用，已使用规则路由：{attempted[0][:120]}）"
        result["backend"] = "rule(fallback)"
    return result


def plan_only(query: str, file_path: str = "") -> dict:
    """只做意图识别、不执行，供界面预览「打算调用哪个工具」。"""
    from core.agent import rule_agent
    return rule_agent.route(query, file_path)
