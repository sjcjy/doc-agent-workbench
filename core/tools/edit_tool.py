from __future__ import annotations

"""工具三：修改文档内容（edit）。

五个动作：
- replace      精确替换文本（支持跨 run 合并后替换，解决中文文档最常见的「替换不生效」）
- insert_section 在指定标题/段落后新增章节
- rewrite     整段改写
- delete      删除段落
- append      文末追加段落

铁律：永远读源文件、写新文件，源文件路径不作为保存目标。
"""

import json
import time

from pydantic import BaseModel, Field
from typing_extensions import Literal

from core.ctx import get_session_id
from core.logger import log_call
from core.models import (
    CODE_INTERNAL_ERROR,
    CODE_INVALID_INPUT,
    CODE_NOT_FOUND_IN_DOC,
    CODE_NO_CHANGE,
    ToolResult,
)
from core.safety import InputGuard, build_output_path, validate_input_file
from core.tools import _docxkit as dk

TOOL_NAME = "edit_document"
TOOL_DESC = (
    "修改一份已有 DOCX 的内容，输出为新文件（源文件保持只读）。"
    "action=replace 精确替换文字，需要 old_text 与 new_text；"
    "action=insert_section 在 anchor 之后新增章节，需要 heading 与 body；"
    "action=rewrite 整段改写，需要 anchor 与 new_text；"
    "action=delete 删除段落，需要 anchor；"
    "action=append 在文末追加段落，需要 new_text。"
    "anchor 可以是段落序号（从 0 开始）或该段落中包含的文字。"
    "修改前建议先调用 inspect_document 确认段落序号。"
)

EditAction = Literal["replace", "insert_section", "rewrite", "delete", "append"]


class EditArgs(BaseModel):
    path: str = Field(description="源 DOCX 文件绝对路径（只读，绝不会被覆盖）")
    action: EditAction = Field(description="修改动作")
    old_text: str = Field(default="", description="replace 时的原文字，必须与文档中完全一致")
    new_text: str = Field(default="", description="替换后 / 追加 / 改写后的文字")
    anchor: str = Field(default="", description="定位用：段落序号（如 3）或段落中包含的文字")
    heading: str = Field(default="", description="insert_section 时的章节标题")
    body: str = Field(default="", description="insert_section 时的章节正文，多行用 \\n 分隔会拆成多段")
    occurrence: int = Field(default=1, description="replace 时替换第几处；0 表示全部替换")
    match_last: bool = Field(default=False, description="anchor 用文字匹配时，是否取最后一个匹配（默认取第一个）")


def _resolve_anchor(doc, anchor: str, match_last: bool = False) -> int | None:
    """把 anchor 解析成段落序号。支持数字序号或文本包含匹配。"""
    a = str(anchor or "").strip()
    if a == "":
        return None
    if a.lstrip("-").isdigit():
        idx = int(a)
        paras = doc.paragraphs
        if 0 <= idx < len(paras):
            return idx
        if idx < 0 and len(paras) + idx >= 0:
            return len(paras) + idx
        return None
    hits = [i for i, p in enumerate(doc.paragraphs) if a in dk.para_text(p)]
    if not hits:
        return None
    return hits[-1] if match_last else hits[0]


def run(
    path: str,
    action: str,
    old_text: str = "",
    new_text: str = "",
    anchor: str = "",
    heading: str = "",
    body: str = "",
    occurrence: int = 1,
    match_last: bool = False,
    session_id: str = "",
) -> ToolResult:
    session_id = session_id or get_session_id()
    t0 = time.perf_counter()
    tool = "edit"
    args = {"path": path, "action": action, "old_text": old_text, "new_text": new_text,
            "anchor": anchor, "heading": heading, "occurrence": occurrence}

    ok_path, err = validate_input_file(path, tool, action)
    if err is not None:
        err.elapsed_ms = int((time.perf_counter() - t0) * 1000)
        log_call(session_id, "tool", TOOL_NAME, action, args, False, err.code, err.message, elapsed_ms=err.elapsed_ms)
        return err
    src = ok_path

    if action not in ("replace", "insert_section", "rewrite", "delete", "append"):
        r = ToolResult.fail(tool, action, CODE_INVALID_INPUT,
                            f"不支持的修改动作：{action}（可选 replace/insert_section/rewrite/delete/append）",
                            elapsed_ms=int((time.perf_counter() - t0) * 1000))
        log_call(session_id, "tool", TOOL_NAME, action, args, False, r.code, r.message, elapsed_ms=r.elapsed_ms)
        return r

    with InputGuard(src) as guard:
        try:
            doc = dk.open_docx(src)
            changes: list[dict] = []
            out_path = build_output_path(src, action, ".docx")

            if action == "replace":
                if not old_text:
                    r = ToolResult.fail(tool, action, CODE_INVALID_INPUT, "replace 需要提供 old_text（要被替换的原文字）",
                                        elapsed_ms=int((time.perf_counter() - t0) * 1000))
                    log_call(session_id, "tool", TOOL_NAME, action, args, False, r.code, r.message, elapsed_ms=r.elapsed_ms)
                    return r
                hits = 0
                for i, p in enumerate(doc.paragraphs):
                    before = dk.para_text(p)
                    n = dk.replace_in_paragraph(p, old_text, new_text, occurrence)
                    if n:
                        hits += n
                        changes.append({"type": "replace", "index": i,
                                        "before": before.strip()[:160], "after": dk.para_text(p).strip()[:160]})
                        if occurrence > 0 and hits >= occurrence:
                            break
                if occurrence == 0:
                    for ti, tb in enumerate(doc.tables):
                        if old_text in tb._tbl.xml:
                            n = dk.replace_in_table(tb, old_text, new_text, 0)
                            hits += n
                            changes.append({"type": "replace_table", "table": ti, "count": n})
                if hits == 0:
                    r = ToolResult.fail(tool, action, CODE_NOT_FOUND_IN_DOC,
                                        f"未在文档中找到「{old_text[:40]}」，未做任何修改（注意是否有空格或标点差异）",
                                        data={"hint": "建议先调用 inspect_document 查看原文"},
                                        elapsed_ms=int((time.perf_counter() - t0) * 1000))
                    log_call(session_id, "tool", TOOL_NAME, action, args, False, r.code, r.message, elapsed_ms=r.elapsed_ms)
                    return r
                summary = f"完成 {hits} 处替换"

            elif action == "insert_section":
                idx = _resolve_anchor(doc, anchor, match_last)
                if idx is None:
                    paras = doc.paragraphs
                    if not anchor:
                        idx = len(paras) - 1 if paras else None
                    if idx is None:
                        r = ToolResult.fail(tool, action, CODE_NOT_FOUND_IN_DOC,
                                            f"未定位到锚点「{anchor}」，请换一个段落序号或包含其中的文字",
                                            elapsed_ms=int((time.perf_counter() - t0) * 1000))
                        log_call(session_id, "tool", TOOL_NAME, action, args, False, r.code, r.message, elapsed_ms=r.elapsed_ms)
                        return r
                if not heading and not body:
                    r = ToolResult.fail(tool, action, CODE_INVALID_INPUT, "insert_section 需要 heading 或 body 至少一项",
                                        elapsed_ms=int((time.perf_counter() - t0) * 1000))
                    log_call(session_id, "tool", TOOL_NAME, action, args, False, r.code, r.message, elapsed_ms=r.elapsed_ms)
                    return r
                anchor_p = doc.paragraphs[idx]
                cursor = anchor_p
                if heading:
                    hp = dk.insert_paragraph_after(cursor, heading)
                    try:
                        hp.style = "Heading 2"
                    except KeyError:
                        dk.set_run_font(hp.runs[0], size_pt=15, bold=True) if hp.runs else None
                    changes.append({"type": "insert_heading", "after_index": idx, "text": heading})
                    cursor = hp
                for b in [x for x in (body or "").replace("\\n", "\n").split("\n") if x.strip()]:
                    bp = dk.insert_paragraph_after(cursor, b.strip())
                    dk.set_line_spacing(bp, 1.5)
                    dk.set_first_line_indent(bp, chars=2)
                    changes.append({"type": "insert_paragraph", "text": b.strip()[:80]})
                    cursor = bp
                summary = f"在第 {idx} 段后新增章节「{heading or body[:20]}」，共插入 {len(changes)} 段"

            elif action == "rewrite":
                idx = _resolve_anchor(doc, anchor, match_last)
                if idx is None:
                    r = ToolResult.fail(tool, action, CODE_NOT_FOUND_IN_DOC,
                                        f"未定位到目标段落（anchor={anchor or '空'}）",
                                        elapsed_ms=int((time.perf_counter() - t0) * 1000))
                    log_call(session_id, "tool", TOOL_NAME, action, args, False, r.code, r.message, elapsed_ms=r.elapsed_ms)
                    return r
                if not new_text:
                    r = ToolResult.fail(tool, action, CODE_INVALID_INPUT, "rewrite 需要提供 new_text（改写后的内容）",
                                        elapsed_ms=int((time.perf_counter() - t0) * 1000))
                    log_call(session_id, "tool", TOOL_NAME, action, args, False, r.code, r.message, elapsed_ms=r.elapsed_ms)
                    return r
                p = doc.paragraphs[idx]
                before = dk.para_text(p).strip()
                if before == new_text.strip():
                    r = ToolResult.fail(tool, action, CODE_NO_CHANGE, "新内容与原文一致，无需改写",
                                        elapsed_ms=int((time.perf_counter() - t0) * 1000))
                    log_call(session_id, "tool", TOOL_NAME, action, args, False, r.code, r.message, elapsed_ms=r.elapsed_ms)
                    return r
                dk.merge_runs(p)
                if p.runs:
                    p.runs[0].text = new_text
                else:
                    p.add_run(new_text)
                changes.append({"type": "rewrite", "index": idx, "before": before[:160], "after": new_text[:160]})
                summary = f"第 {idx} 段已改写"

            elif action == "delete":
                idx = _resolve_anchor(doc, anchor, match_last)
                if idx is None:
                    r = ToolResult.fail(tool, action, CODE_NOT_FOUND_IN_DOC,
                                        f"未定位到要删除的段落（anchor={anchor or '空'}）",
                                        elapsed_ms=int((time.perf_counter() - t0) * 1000))
                    log_call(session_id, "tool", TOOL_NAME, action, args, False, r.code, r.message, elapsed_ms=r.elapsed_ms)
                    return r
                p = doc.paragraphs[idx]
                txt = dk.para_text(p).strip()
                el = p._element
                el.getparent().remove(el)
                changes.append({"type": "delete", "index": idx, "before": txt[:160]})
                summary = f"已删除第 {idx} 段"

            else:  # append
                if not new_text:
                    r = ToolResult.fail(tool, action, CODE_INVALID_INPUT, "append 需要提供 new_text",
                                        elapsed_ms=int((time.perf_counter() - t0) * 1000))
                    log_call(session_id, "tool", TOOL_NAME, action, args, False, r.code, r.message, elapsed_ms=r.elapsed_ms)
                    return r
                cnt = 0
                for line in [x for x in new_text.replace("\\n", "\n").split("\n") if x.strip()]:
                    np_ = doc.add_paragraph(line.strip())
                    dk.set_line_spacing(np_, 1.5)
                    dk.set_first_line_indent(np_, chars=2)
                    cnt += 1
                    changes.append({"type": "append", "text": line.strip()[:80]})
                summary = f"文末追加了 {cnt} 段"

            doc.save(str(out_path))
            r = ToolResult.ok_result(tool, action, f"{summary}，已另存为 {out_path.name}",
                                     output_file=str(out_path), changes=changes,
                                     data={"source": str(src), "change_count": len(changes)},
                                     elapsed_ms=int((time.perf_counter() - t0) * 1000))
            return r

        except Exception as e:
            msg = f"修改文档时出错：{type(e).__name__}: {e}"
            r = ToolResult.fail(tool, action, CODE_INTERNAL_ERROR, msg, elapsed_ms=int((time.perf_counter() - t0) * 1000))
            log_call(session_id, "tool", TOOL_NAME, action, args, False, r.code, r.message, elapsed_ms=r.elapsed_ms)
            return r

    if guard.violated:
        return ToolResult.fail(tool, action, "SAFETY_BLOCKED", "源文件在处理过程中被改动，已中止",
                               elapsed_ms=int((time.perf_counter() - t0) * 1000))
    log_call(session_id, "tool", TOOL_NAME, action, args, r.ok, r.code, r.message,
             output_file=r.output_file, elapsed_ms=r.elapsed_ms,
             extra={"guard": guard.report(), "changes": len(r.changes)})
    return r


def _lc_entry(path: str, action: str, old_text: str = "", new_text: str = "", anchor: str = "",
              heading: str = "", body: str = "", occurrence: int = 1, match_last: bool = False) -> str:
    return json.dumps(run(path, action, old_text, new_text, anchor, heading, body, occurrence, match_last).to_dict(),
                      ensure_ascii=False)


def build_tool():
    from langchain_core.tools import StructuredTool
    return StructuredTool.from_function(
        func=_lc_entry, name=TOOL_NAME, description=TOOL_DESC, args_schema=EditArgs,
    )
