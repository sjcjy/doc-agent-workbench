from __future__ import annotations

"""LangChain + 智谱 GLM 驱动的主 Agent。

要点：
- 用 create_agent（LangChain 1.x 的入口，内部带工具循环）；
- 把会话里的每一次工具调用从 messages 中还原出来，作为「调用记录」交给界面；
- 任何异常都抛给上层，由上层决定是否回退到规则路由。
"""

import json
import time
from typing import Any

from langchain.agents import create_agent

from core import config
from core.agent.prompt import SYSTEM_PROMPT, build_user_message
from core.logger import log_call
from core.models import ToolResult


class ModelUnavailable(Exception):
    """模型不可用（缺 Key、超时、鉴权失败、额度不足等）。"""


def build_model():
    """构造聊天模型实例。默认走智谱的 OpenAI 兼容端点。"""
    provider = (config.MODEL_PROVIDER or "zhipu-openai").lower()

    if provider in ("zhipu-openai", "zhipu", "openai"):
        from langchain_openai import ChatOpenAI

        if provider.startswith("zhipu"):
            if not config.ZHIPU_API_KEY:
                raise ModelUnavailable("未配置 ZHIPU_API_KEY，无法调用智谱模型")
            api_key, base_url, model = config.ZHIPU_API_KEY, config.ZHIPU_BASE_URL, config.ZHIPU_MODEL
        else:
            if not config.OPENAI_API_KEY:
                raise ModelUnavailable("未配置 OPENAI_API_KEY")
            api_key, base_url, model = config.OPENAI_API_KEY, config.OPENAI_BASE_URL or None, config.OPENAI_MODEL
        return ChatOpenAI(
            model=model,
            api_key=api_key,
            base_url=base_url,
            temperature=config.MODEL_TEMPERATURE,
            timeout=config.MODEL_TIMEOUT_SEC,
            max_retries=config.MODEL_MAX_RETRIES,
        )

    # 预留：官方接口的原生接入方式
    raise ModelUnavailable(f"暂不支持的 MODEL_PROVIDER：{provider}")


def _to_text(content: Any) -> str:
    """把不同形态的 message content 统一成字符串。"""
    if isinstance(content, str):
        return content
    if isinstance(content, list):
        buf = []
        for c in content:
            if isinstance(c, str):
                buf.append(c)
            elif isinstance(c, dict) and c.get("type") == "text":
                buf.append(str(c.get("text", "")))
        return "".join(buf)
    return str(content or "")


def _parse_result(raw: str) -> dict:
    """工具返回的是 JSON 字符串，解析成 dict；解析不了就包一层。"""
    try:
        return json.loads(raw)
    except Exception:
        return {"ok": True, "tool": "unknown", "action": "", "message": raw[:200], "code": "OK",
                "output_file": None, "changes": [], "data": {}}


def run(query: str, file_path: str = "", session_id: str = "", extra_context: str = "") -> dict:
    """执行一次任务。返回统一结构：{ok, backend, reply, steps, outputs, error}。"""
    t0 = time.perf_counter()

    from core.tools import build_tools

    model = build_model()
    tools = build_tools()
    agent = create_agent(model=model, tools=tools, system_prompt=SYSTEM_PROMPT)

    msg = build_user_message(query, file_path, extra_context)
    try:
        result = agent.invoke(
            {"messages": [{"role": "user", "content": msg}]},
            config={"recursion_limit": max(10, config.AGENT_MAX_STEPS * 3)},
        )
    except Exception as e:  # 网络、鉴权、限流都在这里汇合
        elapsed = int((time.perf_counter() - t0) * 1000)
        reason = f"{type(e).__name__}: {e}"
        log_call(session_id, "model", "glm", "invoke", {"query": query[:200]}, False, "MODEL_FAILED", reason, elapsed_ms=elapsed)
        raise ModelUnavailable(reason)

    messages = result.get("messages", []) if isinstance(result, dict) else []
    steps: list[dict] = []
    outputs: list[str] = []
    reply = ""

    seq = 0
    # 第一遍：把所有工具返回按 tool_call_id 收好（消息顺序是 AI(tool_calls) 在前、
    # ToolMessage 在后，必须先把结果收齐，再解析 AI 调用，否则 output_file 会全部落空）。
    tool_outputs: dict[str, str] = {}
    for m in messages:
        role = getattr(m, "role", None) or getattr(m, "type", "")
        if role in ("tool", "ToolMessage"):
            tid = getattr(m, "tool_call_id", None) or getattr(m, "id", None)
            if tid:
                tool_outputs[str(tid)] = _to_text(getattr(m, "content", ""))

    # 第二遍：按顺序解析 AI 的工具调用，对齐各自的结果
    for m in messages:
        role = getattr(m, "role", None) or getattr(m, "type", "")
        if role in ("ai", "AIMessage"):
            calls = getattr(m, "tool_calls", None) or []
            if calls:
                for c in calls:
                    seq += 1
                    name = c.get("name") if isinstance(c, dict) else getattr(c, "name", "")
                    args = c.get("args") if isinstance(c, dict) else getattr(c, "args", {})
                    cid = c.get("id") if isinstance(c, dict) else getattr(c, "id", "")
                    raw_res = tool_outputs.get(str(cid), "")
                    parsed = _parse_result(raw_res) if raw_res else {}
                    if parsed.get("output_file"):
                        outputs.append(parsed["output_file"])
                    steps.append({"seq": seq, "tool": name, "args": args, "result": parsed})
            txt = _to_text(getattr(m, "content", "")).strip()
            if txt:
                reply = txt  # 保留最后一条有文字的回复

    elapsed = int((time.perf_counter() - t0) * 1000)
    log_call(session_id, "agent", "langchain", "run", {"query": query[:200]}, True, "OK",
             f"调用工具 {len(steps)} 次", output_file=(outputs[-1] if outputs else None),
             elapsed_ms=elapsed, backend="glm", extra={"model": getattr(model, "model_name", "")})

    return {
        "ok": True,
        "backend": f"glm:{getattr(model, 'model_name', 'unknown')}",
        "reply": reply or f"已完成任务，共调用工具 {len(steps)} 次。",
        "steps": steps,
        "outputs": outputs,
        "error": None,
        "elapsed_ms": elapsed,
    }
