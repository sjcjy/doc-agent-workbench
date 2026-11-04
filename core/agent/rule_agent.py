from __future__ import annotations

"""规则降级 Agent：不依赖大模型的自然语言 → 工具调用路由。

存在意义：演示现场最怕 API 繁忙或 Key 失效。这个模块在没有模型的情况下
依然能跑通「撰写 / 修改 / 格式 / 转换」四项功能，保证验收不翻车。
它只做模式匹配与参数抽取，最终仍然把活交给同一套确定性工具。

排在后端的第二顺位，有 Key 时不影响主路径。
"""

import json
import re
import time

from core.logger import log_call
from core.tools import dispatch

# 中文印刷字号 → 磅
CN_SIZE_MAP = {
    "初号": 42, "小初": 36, "一号": 26, "小一": 24, "二号": 22, "小二": 18,
    "三号": 16, "小三": 15, "四号": 14, "小四": 12, "五号": 10.5, "小五": 9,
    "六号": 7.5, "小六": 6.5,
}
FONT_WORDS = ("宋体", "黑体", "微软雅黑", "仿宋", "楷体", "华文宋体", "华文楷体",
              "隶书", "等线", "Times New Roman", "Arial", "Calibri")

_ROUTE_FAILED = {
    "ok": False,
    "backend": "rule",
    "reply": "",
    "steps": [],
    "outputs": [],
    "error": "未能理解这条指令",
    "suggestion": "",
    "elapsed_ms": 0,
}


def _fail(reason: str, suggestion: str = "", elapsed_ms: int = 0) -> dict:
    d = dict(_ROUTE_FAILED)
    d["error"] = reason
    d["suggestion"] = suggestion
    d["elapsed_ms"] = elapsed_ms
    return d


def _need_file(file_path: str | None) -> bool:
    return bool(file_path and str(file_path).strip())


# ---------------------------------------------------------------- 参数抽取
def _strip_quotes(s: str) -> str:
    return re.sub(r'^["“‘\'（《<]+|["”’\'）》>]+$', "", (s or "").strip()).strip()


def extract_replace_pair(q: str) -> tuple[str, str] | None:
    """从「把 A 改成 B」这类指令中抽取替换前后文本。"""
    patterns = [
        r'把\s*(.{1,60}?)\s*(?:改成|改为|替换成|替换为|修改成|修改为|替换为|替换成)\s*(.{1,80}?)(?:[。！；;，,]|$)',
        r'将\s*(.{1,60}?)\s*(?:改成|改为|替换成|替换为|修改成|修改为)\s*(.{1,80}?)(?:[。！；;，,]|$)',
        r'(?:替换|替换掉)\s*(.{1,60}?)\s*(?:为|成)\s*(.{1,80}?)(?:[。！；;，,]|$)',
    ]
    for pat in patterns:
        m = re.search(pat, q)
        if m:
            old = _strip_quotes(m.group(1))
            new = _strip_quotes(m.group(2))
            if old and new and old != new:
                return old, new
    return None


def extract_size(q: str) -> float | None:
    for word, pt in CN_SIZE_MAP.items():
        if word in q:
            return pt
    m = re.search(r'(\d+(?:\.\d+)?)\s*(?:磅|pt|PT|号字)', q)
    if m:
        try:
            v = float(m.group(1))
            if 5 <= v <= 72:
                return v
        except ValueError:
            pass
    return None


def extract_font(q: str) -> str | None:
    for f in FONT_WORDS:
        if f in q:
            return f
    m = re.search(r'字体\s*(?:改成|改为|设为|设置为|换成)?\s*([\u4e00-\u9fa5A-Za-z ]{2,12})', q)
    if m:
        cand = m.group(1).strip()
        if cand and cand not in ("是什么", "怎么样"):
            return cand
    return None


def extract_spacing(q: str) -> tuple[float, bool] | None:
    """返回 (行距, 是否固定值)。"""
    if re.search(r'单倍行距|一倍行距|1\.0倍行距', q):
        return 1.0, False
    m = re.search(r'固定(?:值)?\s*(\d+(?:\.\d+)?)\s*(?:磅|pt)', q)
    if m:
        return float(m.group(1)), True
    m = re.search(r'(\d+(?:\.\d+)?)\s*倍(?:行距)?', q)
    if m:
        v = float(m.group(1))
        if 0.5 <= v <= 5:
            return v, False
    # 「设置行距」但没给数值 → 默认 1.5
    if re.search(r'行距', q):
        return 1.5, False
    return None


def extract_indent(q: str) -> float | None:
    m = re.search(r'(?:首行)?缩进\s*(\d+(?:\.\d+)?)\s*字符', q)
    if m:
        return float(m.group(1))
    m = re.search(r'缩进\s*(\d+(?:\.\d+)?)\s*(?:磅|pt)', q)
    if m:
        return -float(m.group(1))  # 负数交给工具按磅处理
    if "首行缩进" in q or "缩进" in q:
        return 2.0
    return None


_TARGET_FORMAT_WORDS = {
    "pdf": "pdf",
    "markdown": "md", "md": "md",
    "docx": "docx", "word": "docx",
    "纯文本": "txt", "txt": "txt", "文本": "txt",
}
_FORMAT_WORD_RE = r'(pdf|markdown|docx|word|纯文本|txt|md|文本)'


def extract_target_format(q: str) -> str | None:
    """识别「要转成什么格式」。

    关键是取「转成 / 导出」后面的目标词，而不是看句子里出现了哪个格式名——
    否则「把 PDF 转成 Word」会因为句子里有 pdf 而被误判成要转 PDF。
    """
    low = q.lower()

    # 1) 转（换）成/为/到 … <目标>；也覆盖「PDF 转 Word」这种省略「成」的说法
    m = re.search(r'(?:转换|转|另存)\s*(?:换成|换为|成|为|到|化成|化)?[^，,。；;]{0,8}?'
                  + _FORMAT_WORD_RE, low)
    if m:
        return _TARGET_FORMAT_WORDS.get(m.group(1))

    # 2) 导出 / 生成 / 输出 … <目标>
    m = re.search(r'(?:导出|生成|输出)\s*(?:为|成)?\s*' + _FORMAT_WORD_RE, low)
    if m:
        return _TARGET_FORMAT_WORDS.get(m.group(1))

    # 3) 兜底：没有明确动词时按单一关键词判断
    if re.search(r'pdf\s*(?:版|版本)', low):
        return "pdf"
    if "markdown" in low:
        return "md"
    if "纯文本" in low:
        return "txt"
    return None


def extract_insert(q: str) -> dict | None:
    """抽取新增章节的意图：{heading, body, anchor}。"""
    kw = ("新增章节", "增加章节", "添加章节", "插入章节", "新增一节", "增加一节", "新增一段", "添加一段")
    if not any(k in q for k in kw):
        return None
    heading, body, anchor = "", "", ""
    m = re.search(r'(?:标题|章节名)\s*[:：是为]?\s*["“]?([^"”，,。；;]{1,40})', q)
    if m:
        heading = _strip_quotes(m.group(1))
    m2 = re.search(r'(?:在|于)\s*(.{1,30}?)\s*(?:之后|后面|下面)', q)
    if m2:
        anchor = _strip_quotes(m2.group(1))
    m3 = re.search(r'(?:内容|正文)\s*[:：是为]?\s*["“]?([^"”]{1,200})', q)
    if m3:
        body = _strip_quotes(m3.group(1))
    if not body and not heading:
        # 指令里引号中的第一段当成标题
        m4 = re.search(r'["“]([^"”]{1,40})["”]', q)
        if m4:
            heading = m4.group(1)
    return {"heading": heading, "body": body, "anchor": anchor}


def extract_topic(q: str) -> str:
    """从「写一份 X」中抽取主题 X。"""
    # 「内容包括…」之后的部分是正文要点，不属于标题
    head = re.split(r'(?:内容包括|内容包含|包含|要点|包括)', q)[0]
    m = re.search(
        r'(?:写|生成|新建|创建|起草|做)\s*(?:一份|一篇|一个|个)?\s*(?:关于)?\s*'
        r'["“]?([^"”，,。！；;？?]{1,30})',
        head,
    )
    if m:
        t = _strip_quotes(m.group(1)).strip()
        t = re.sub(r'^(的|关于)', "", t).strip()
        if t and len(t) <= 30:
            return t
    return "新建文档"


def _path_list(file_path: str) -> list[str]:
    """file_path 允许用 | 分隔多个文件，供合并与对比使用。"""
    if not file_path:
        return []
    return [p.strip() for p in str(file_path).split("|") if p.strip()]


_CN_NUM = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "1": 1, "2": 2, "3": 3, "4": 4, "5": 5}


def _split_level(q: str) -> int:
    m = re.search(r'([一二三四五1-5])\s*级标题', q)
    if m:
        return _CN_NUM.get(m.group(1), 1)
    return 1


# ---------------------------------------------------------------- 主路由
def route(query: str, file_path: str = "") -> dict:
    """把一条自然语言指令翻译成一次工具调用。

    返回 {"name": 工具名, "args": {...}} 或 {"error": 原因, "suggestion": 建议}。
    """
    q = (query or "").strip()
    if not q:
        return {"error": "任务内容为空", "suggestion": "请输入一句自然语言任务，例如「把这份文档的标题居中」"}

    has_file = _need_file(file_path)
    paths = _path_list(file_path)

    # ---- 0) 结构操作：合并 / 拆分 / 提取表格 / 对比
    # 放在最前面，避免「对比两份文档的差异」被「查看」规则先吃掉。
    if re.search(r'合并|拼接|合到一块|合成一份|汇总成一份', q):
        if len(paths) < 2:
            return {"error": "合并需要两份以上文档",
                    "suggestion": "展开「📎 多份文档操作」面板，勾选文件后点「合并」；"
                                  "接口调用时可用 | 把多个路径连起来"}
        return {"name": "structure_document", "args": {
            "path": paths[0], "action": "merge",
            "other_paths": "|".join(paths[1:]),
            "add_page_break": not bool(re.search(r'不分页|不要分页|连续排', q))}}

    if re.search(r'(拆分|分割|拆成|切成|分开)', q):
        if not paths:
            return {"error": "拆分需要先选中或上传一份文档"}
        m = re.search(r'(?:从|在)\s*["“]?(.{1,30}?)["”]?\s*(?:开始|处|这一?节?)\s*(?:拆分|拆分|切|分)', q)
        return {"name": "structure_document", "args": {
            "path": paths[0], "action": "split",
            "level": _split_level(q), "anchor": _strip_quotes(m.group(1)) if m else ""}}

    if re.search(r'(提取|导出|抽取|另存).{0,8}表格|表格.{0,8}(提取|导出|另存|存成|保存成)', q):
        if not paths:
            return {"error": "提取表格需要先选中或上传一份文档"}
        return {"name": "structure_document", "args": {
            "path": paths[0], "action": "extract_tables"}}

    if re.search(r'(对比|比对|比较|差异|diff)', q) and re.search(r'(文档|版本|两版|差异|对比|比对|比较|不同|区别|改了什么)', q):
        if len(paths) < 2:
            return {"error": "对比需要两份文档",
                    "suggestion": "展开「📎 多份文档操作」面板，勾选两份文件后点「对比差异」"}
        return {"name": "structure_document", "args": {
            "path": paths[0], "action": "diff", "other_paths": paths[1]}}

    # ---- 1) 检查 / 查看
    if re.search(r'^(看看|查看|看一下|看下|检查| Inspect|结构|大纲|有哪些|什么内容)', q) or \
       re.search(r'(查看|看看|检查|统计|检索|搜索)\s*(?:一下)?\s*(?:这份|这个|文档|结构|大纲)', q):
        if not has_file:
            return {"error": "查看文档需要先选中或上传一份文档"}
        kw = ""
        m = re.search(r'(?:检索|搜索|查找|找一下|有没有)\s*["“]?([^"” ，,。]{1,30})', q)
        if m:
            kw = _strip_quotes(m.group(1))
        return {"name": "inspect_document", "args": {"path": file_path, **({"keyword": kw} if kw else {})}}

    # ---- 2) 转换（优先于格式，避免「导出 pdf」被别的规则吃掉）
    fmt = extract_target_format(q)
    if fmt:
        if not has_file:
            return {"error": f"转换成 {fmt.upper()} 需要先选中或上传源文件"}
        return {"name": "convert_document", "args": {"path": file_path, "target": fmt}}

    # ---- 3) 精准替换
    # 注意：句子里出现格式关键词时，优先按格式处理，
    # 否则「把正文字体改成宋体」会被误判成「把『正文字体』替换成『宋体』」。
    _FMT_KW = ("字体", "字号", "行距", "缩进", "居中", "对齐", "加粗", "粗体", "排版", "磅")
    pair = None if any(k in q for k in _FMT_KW) else extract_replace_pair(q)
    if pair:
        if not has_file:
            return {"error": "修改文档内容需要先选中或上传一份文档"}
        old, new = pair
        all_match = bool(re.search(r'全部|所有|每一处|处处', q))
        return {"name": "edit_document", "args": {
            "path": file_path, "action": "replace", "old_text": old, "new_text": new,
            "occurrence": 0 if all_match else 1}}

    # ---- 4) 新增章节 / 追加
    ins = extract_insert(q)
    if ins:
        if not has_file:
            return {"error": "新增章节需要先选中或上传一份文档"}
        if not ins["heading"] and not ins["body"]:
            return {"error": "没能识别出新增章节的标题或内容",
                    "suggestion": "可以这样说：在「三、下周计划」后面新增章节，标题为「四、风险提示」，内容为「需要关注接口稳定性」"}
        return {"name": "edit_document", "args": {
            "path": file_path, "action": "insert_section",
            "heading": ins["heading"], "body": ins["body"], "anchor": ins["anchor"]}}

    if re.search(r'(追加|末尾添加|文末添加|在最后添加|在末尾加)', q):
        if not has_file:
            return {"error": "追加内容需要先选中或上传一份文档"}
        m = re.search(r'(?:追加|添加|加上)\s*[:：]?\s*["“]?([^"”]{1,200})', q) or \
            re.search(r'["“]([^"”]{1,200})["”]', q)
        body = _strip_quotes(m.group(1)) if m else ""
        if not body:
            return {"error": "没能识别出要追加的内容", "suggestion": "试试：追加：本周五前完成联调"}
        return {"name": "edit_document", "args": {"path": file_path, "action": "append", "new_text": body}}

    # ---- 5) 删除
    if re.search(r'(删除|删掉|去掉)\s*(第?\s*\d+\s*段|包含.{1,20}的段落)', q) or \
       re.search(r'删除\s*["“]?([^"”]{1,40})', q):
        if not has_file:
            return {"error": "删除段落需要先选中或上传一份文档"}
        m = re.search(r'第?\s*(\d+)\s*段', q)
        anchor = m.group(1) if m else ""
        if not anchor:
            m2 = re.search(r'删除\s*["“]?([^"”]{1,40})', q)
            anchor = _strip_quotes(m2.group(1)) if m2 else ""
        if not anchor:
            return {"error": "没能识别出要删除哪一段", "suggestion": "试试：删除第 3 段，或 删除包含「备注」的段落"}
        return {"name": "edit_document", "args": {"path": file_path, "action": "delete", "anchor": anchor}}

    # ---- 6) 格式
    if re.search(r'(标题|一级标题|所有标题|全文)?\s*居中', q):
        if not has_file:
            return {"error": "设置居中需要先选中或上传一份文档"}
        target = "headings" if re.search(r'标题', q) else "all"
        return {"name": "format_document", "args": {"path": file_path, "action": "align",
                                                    "target": target, "alignment": "center"}}

    if re.search(r'(左对齐|右对齐|两端对齐)', q):
        if not has_file:
            return {"error": "设置对齐需要先选中或上传一份文档"}
        al = "left" if "左对齐" in q else ("right" if "右对齐" in q else "justify")
        target = "headings" if re.search(r'标题', q) else "all"
        return {"name": "format_document", "args": {"path": file_path, "action": "align",
                                                    "target": target, "alignment": al}}

    if re.search(r'行距', q) and not re.search(r'行距是|行距多少', q):
        if not has_file:
            return {"error": "调整行距需要先选中或上传一份文档"}
        sp = extract_spacing(q)
        if not sp:
            return {"error": "没能识别出行距数值", "suggestion": "试试：把行距设为 1.5 倍"}
        val, exact = sp
        return {"name": "format_document", "args": {
            "path": file_path, "action": "line_spacing",
            "target": "body" if re.search(r'正文', q) else "all",
            "line_spacing": val, "line_spacing_exact": exact}}

    if re.search(r'(首行)?缩进', q):
        if not has_file:
            return {"error": "设置缩进需要先选中或上传一份文档"}
        val = extract_indent(q)
        if val is None:
            return {"error": "没能识别出缩进数值"}
        args = {"path": file_path, "action": "indent", "target": "all"}
        if val < 0:
            args["indent_pt"] = -val
        else:
            args["indent_chars"] = val
        return {"name": "format_document", "args": args}

    if re.search(r'(一键排版|报告排版|规范排版|套用模板|论文排版)', q):
        if not has_file:
            return {"error": "一键排版需要先选中或上传一份文档"}
        return {"name": "format_document", "args": {"path": file_path, "action": "preset_report"}}

    if re.search(r'(字体|字号|加粗|粗体)', q):
        if not has_file:
            return {"error": "设置字体需要先选中或上传一份文档"}
        font = extract_font(q)
        size = extract_size(q)
        bold = True if re.search(r'加粗|粗体', q) else None
        if not font and not size and bold is None:
            return {"error": "没能识别出字体或字号", "suggestion": "试试：把正文字体改成宋体，字号设为小四"}
        return {"name": "format_document", "args": {
            "path": file_path, "action": "font",
            "target": "headings" if re.search(r'标题', q) else ("body" if re.search(r'正文', q) else "all"),
            "font_name": font or "", "size_pt": size or 0, "bold": bold}}

    # ---- 7) 撰写
    if re.search(r'(写|生成|新建|创建|起草|做)\s*(一份|一篇|一个)?', q) and not re.search(r'(这段|该段|第\s*\d+\s*段)', q):
        topic = extract_topic(q)
        title = topic
        outline = _template_outline(topic, q)
        return {"name": "compose_document", "args": {
            "topic": topic, "title": title, "outline": outline,
            "with_cover": bool(re.search(r'封面', q))}}

    return {"error": "没有匹配到任何可执行的操作",
            "suggestion": "试试这些说法：查看这份文档 / 把「张三」改成「李四」/ 标题居中 / "
                          "正文改成宋体小四 / 行距设为 1.5 倍 / 转成 PDF / 写一份项目总结"}


def _template_outline(topic: str, query: str) -> str:
    """降级模式下的正文模板。

    诚实地说明这是模板：没有大模型时只能给出结构框架，内容需要用户补充。
    """
    points = []
    m = re.search(r'(?:内容包括|包含|要点[:：]?|包括)\s*(.{1,200})', query)
    if m:
        raw = m.group(1)
        points = [p.strip(" ，,。；;") for p in re.split(r'[、,，。;；/]', raw) if p.strip()][:8]

    lines = [f"# {topic}", ""]
    if points:
        lines += ["## 主要内容", ""]
        for p in points:
            lines += [f"- {p}"]
    else:
        lines += [
            "## 主要内容", "",
            "（降级模式生成）此处填写第一部分的核心内容。",
            "", "## 工作安排", "",
            "（降级模式生成）此处填写具体的执行安排与责任人。",
            "", "## 风险与对策", "",
            "（降级模式生成）此处填写可能的风险与对应的应对措施。",
        ]
    lines += ["", "## 说明", "", "本文档由文档工作台的降级模式生成，"
              "仅提供结构框架，正文内容请在使用时补充完善。", ""]
    return "\n".join(lines)


# ---------------------------------------------------------------- 执行
def run(query: str, file_path: str = "", session_id: str = "", extra_context: str = "") -> dict:
    """按规则路由执行一次任务，返回与 llm_agent 一致的结构。"""
    t0 = time.perf_counter()

    plan = route(query, file_path)
    if "error" in plan:
        elapsed = int((time.perf_counter() - t0) * 1000)
        log_call(session_id, "agent", "rule_agent", "route", {"query": query[:200]}, False,
                 "ROUTE_FAILED", plan["error"], elapsed_ms=elapsed, backend="rule")
        return _fail(plan["error"], plan.get("suggestion", ""), elapsed)

    name, args = plan["name"], plan["args"]
    log_call(session_id, "agent", "rule_agent", "route", {"query": query[:200]}, True, "OK",
             f"路由到 {name}", backend="rule", elapsed_ms=int((time.perf_counter() - t0) * 1000))

    try:
        res = dispatch(name, **args, session_id=session_id)
    except Exception as e:
        elapsed = int((time.perf_counter() - t0) * 1000)
        return _fail(f"工具执行异常：{type(e).__name__}: {e}", "", elapsed)

    elapsed = int((time.perf_counter() - t0) * 1000)
    outputs = [res.output_file] if res.output_file else []
    reply = res.message
    if res.ok and res.output_file:
        from pathlib import Path
        reply = f"{res.message}\n\n输出文件：{Path(res.output_file).name}"
    elif not res.ok:
        reply = f"未能完成：{res.message}"

    log_call(session_id, "agent", "rule_agent", "run", {"query": query[:200]}, res.ok, res.code,
             res.message, output_file=res.output_file, elapsed_ms=elapsed, backend="rule")

    return {
        "ok": res.ok,
        "backend": "rule",
        "reply": reply,
        "steps": [{"seq": 1, "tool": name, "args": args, "result": res.to_dict()}],
        "outputs": outputs,
        "error": None if res.ok else res.message,
        "suggestion": "" if res.ok else (res.data.get("hint", "") or "可先用「查看这份文档」确认内容后再操作"),
        "elapsed_ms": elapsed,
    }
