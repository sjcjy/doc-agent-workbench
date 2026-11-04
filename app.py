# -*- coding: utf-8 -*-
"""文档工作台智能体 · Streamlit 界面。

运行：streamlit run app.py
界面只做三件事：收集输入、调用核心层、展示结果。
所有业务逻辑都在 core/ 里，后续转成小程序时可以直接复用 FastAPI 那一套接口。
"""

import re
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

import streamlit as st

from core import config
from core.agent import execute
from core.ctx import set_session_id
from core.logger import new_session_id, read_calls
from core.safety import InputGuard, save_upload, sha256_of
from core.tools import convert_tool, dispatch

st.set_page_config(page_title="文档工作台智能体", page_icon="📄", layout="wide")

# 快捷指令：按四大功能 + 查看分类。text 必须与 rule_agent 能识别的措辞一致，
# 这样在没有大模型（无 Key）的降级模式下，点了也一定能执行成功。
QUICK_GROUPS = [
    {
        "key": "inspect",
        "label": "📋 查看",
        "items": [
            {"label": "查看文档结构", "text": "查看这份文档的结构"},
            {"label": "统计字数段落", "text": "统计这份文档的字数和段落数"},
            {"label": "查找关键词", "text": "查看文档有没有“关键词”"},
        ],
    },
    {
        "key": "compose",
        "label": "✍️ 撰写",
        "items": [
            {"label": "项目周报", "text": "写一份项目周报，内容包括本周进展、遇到的问题、下周计划"},
            {"label": "会议纪要", "text": "写一份会议纪要，内容包括会议主题、参会人员、讨论要点、待办事项"},
            {"label": "工作总结", "text": "写一份工作总结，内容包括主要成果、存在问题、改进措施"},
            {"label": "实验报告", "text": "写一份实验报告，内容包括实验目的、实验步骤、实验结果、结论"},
        ],
    },
    {
        "key": "edit",
        "label": "✏️ 修改",
        "items": [
            {"label": "替换文本", "text": "把“原文”改成“新文本”"},
            {"label": "全部替换", "text": "把“原文”改成“新文本”，全部替换"},
            {"label": "追加一段", "text": "追加：本周已完成接口联调"},
            {"label": "新增章节", "text": "新增章节，标题：风险提示，内容：需要关注接口稳定性"},
        ],
    },
    {
        "key": "format",
        "label": "🎨 格式",
        "items": [
            {"label": "标题居中", "text": "把所有标题居中"},
            {"label": "正文宋体小四", "text": "把正文字体改成宋体，字号设为小四"},
            {"label": "1.5 倍行距", "text": "把行距设为 1.5 倍"},
            {"label": "首行缩进 2 字符", "text": "首行缩进 2 字符"},
            {"label": "一键报告排版", "text": "一键套用报告排版"},
            {"label": "标题加粗", "text": "把标题加粗"},
        ],
    },
    {
        "key": "convert",
        "label": "🔄 转换",
        "items": [
            {"label": "转 Markdown", "text": "转成 Markdown"},
            {"label": "转 PDF", "text": "转成 PDF"},
            {"label": "转 TXT", "text": "转成 TXT"},
            {"label": "PDF 转 Word", "text": "把 PDF 转成 Word"},
        ],
    },
    {
        "key": "structure",
        "label": "🧩 结构",
        "items": [
            {"label": "合并多份文档", "text": "把选中的多份文档合并成一份"},
            {"label": "按章节拆分", "text": "按一级标题把这份文档拆成多份"},
            {"label": "提取表格到 Excel", "text": "把文档里的表格提取出来存成 Excel"},
            {"label": "对比两版差异", "text": "对比这两份文档有哪些差异"},
        ],
    },
]


# ---------------------------------------------------------------- 会话状态
def init_state() -> None:
    st.session_state.setdefault("session_id", new_session_id())
    st.session_state.setdefault("current_file", "")
    st.session_state.setdefault("hash_before", "")
    st.session_state.setdefault("history", [])
    st.session_state.setdefault("last_result", None)


init_state()


# ---------------------------------------------------------------- 文件选择
def list_docs() -> dict[str, Path]:
    """列出可选文档：DOCX 与 PDF（PDF 需先转成 Word 才能编辑）。

    输出目录只取最近 12 个，避免历史产物把下拉框撑到几十项。
    """
    files: dict[str, Path] = {}
    for label, root, limit in (("样例", config.SAMPLE_DIR, None),
                               ("上传", config.UPLOAD_DIR, None),
                               ("输出", config.OUTPUT_DIR, 12)):
        if not root.exists():
            continue
        items: list[Path] = []
        for ext in ("*.docx", "*.pdf"):
            items.extend(root.glob(ext))
        items.sort(key=lambda x: x.stat().st_mtime, reverse=True)
        if limit:
            items = items[:limit]
        for p in items:
            files[f"[{label}] {p.name}"] = p
    return files


def render_file_download(result_row: dict) -> None:
    """为一次工具结果渲染下载按钮。"""
    path = result_row.get("output_file")
    if not path or not Path(path).exists():
        return
    p = Path(path)
    mime = {".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            ".pdf": "application/pdf", ".md": "text/markdown", ".txt": "text/plain"}.get(p.suffix, "application/octet-stream")
    size_kb = p.stat().st_size / 1024
    st.download_button(
        label=f"⬇ 下载 {p.name}（{size_kb:.0f} KB）",
        data=p.read_bytes(),
        file_name=p.name,
        mime=mime,
        key=f"dl_{p.name}_{result_row.get('_uid', '')}",
    )


def _direct_result(res, task_label: str, args: dict) -> dict:
    """把直接调用的 ToolResult 包成与 agent 一致的结构，复用同一套结果展示。"""
    outputs = [res.output_file] if getattr(res, "output_file", None) else []
    reply = res.message
    if res.ok and res.output_file:
        reply = f"{res.message}\n\n输出文件：{Path(res.output_file).name}"
    elif not res.ok:
        reply = f"未能完成：{res.message}"
    tip = "" if res.ok else "请确认选中的文件是否正确，或换一种拆分依据再试。"
    return {
        "ok": res.ok,
        "backend": "direct",
        "reply": reply,
        "steps": [{"seq": 1, "tool": res.tool, "args": args, "result": res.to_dict()}],
        "outputs": outputs,
        "error": None if res.ok else res.message,
        "suggestion": tip,
        "elapsed_ms": res.elapsed_ms,
        "_task": task_label,
    }


# ---------------------------------------------------------------- 侧边栏
with st.sidebar:
    st.title("📄 文档工作台")
    st.caption("选一份文档，用一句话说出你想做什么")

    st.subheader("1. 选择文档")
    uploaded = st.file_uploader(
        "上传 DOCX / PDF / TXT / MD",
        type=["docx", "pdf", "txt", "md", "markdown"],
        label_visibility="collapsed",
    )
    if uploaded is not None:
        target, err = save_upload(uploaded.getvalue(), uploaded.name)
        if err is not None:
            st.error(f"上传失败：{err.message}")
        else:
            st.session_state.current_file = str(target)
            st.session_state.hash_before = sha256_of(target)
            st.success(f"已上传：{target.name}")

    docs = list_docs()
    if docs:
        cur = st.session_state.current_file
        names = list(docs.keys())
        idx = 0
        for i, n in enumerate(names):
            if str(docs[n]) == cur:
                idx = i
                break
        picked = st.selectbox("或从已有文件中选择", names, index=idx)
        if str(docs[picked]) != st.session_state.current_file:
            st.session_state.current_file = str(docs[picked])
            st.session_state.hash_before = sha256_of(docs[picked])

    if st.session_state.current_file:
        p = Path(st.session_state.current_file)
        st.info(f"当前文档：{p.name}")
        if p.suffix.lower() == ".pdf":
            st.caption("这是 PDF。想编辑内容，请先点「🔄 转换 → PDF 转 Word」，转成 Word 后再操作。")

    st.subheader("2. 高级设置")
    with st.expander("⚙️ 技术信息（开发者 / 验收用）", expanded=False):
        _default_mode = config.AGENT_BACKEND if config.AGENT_BACKEND in ("auto", "llm", "rule") else "auto"
        mode = st.radio(
            "Agent 后端",
            ["auto", "llm", "rule"],
            format_func=lambda x: {"auto": "自动（优先大模型，失败降级）",
                                   "llm": "仅大模型（智谱 GLM）",
                                   "rule": "仅规则路由（无需 Key）"}[x],
            index=["auto", "llm", "rule"].index(_default_mode),
        )

        has_key = bool(config.ZHIPU_API_KEY)
        if has_key:
            st.success(f"已检测到 ZHIPU_API_KEY · 模型 {config.ZHIPU_MODEL}")
        else:
            st.warning("未配置 ZHIPU_API_KEY，大模型不可用，将使用规则路由")
        st.caption(f"MODEL_PROVIDER = {config.MODEL_PROVIDER}")

        be = convert_tool.pdf_backends()
        st.caption("PDF 生成后端：" + "、".join(f"{'✓' if v else '✗'} {k}" for k, v in be.items()))
        st.caption(f"会话 ID：{st.session_state.session_id}")

    if st.button("🧹 清空调用记录", use_container_width=True):
        try:
            Path(config.CALL_LOG_PATH).unlink(missing_ok=True)
            st.session_state.history = []
            st.rerun()
        except OSError:
            pass


# ---------------------------------------------------------------- 主区
st.header("文档工作台智能体")
st.caption("上传或选择一份文档，用一句话说出你想怎么改；也可以直接让它写一份新文档。")

if not st.session_state.current_file:
    st.warning("请先在左侧选择或上传一份文档（撰写类任务可跳过）。")

# ---------------------------------------------------------------- 任务输入
st.session_state.setdefault("task_query", "")


def _fill_query(text: str) -> None:
    """点击快捷指令 → 把指令填入输入框，用户可再修改后执行。"""
    st.session_state["task_query"] = text


task = st.text_area(
    "用一句话说出你的任务",
    key="task_query",
    height=110,
    placeholder="例如：把「联系邮箱：project@example.com」改成「联系邮箱：team@example.com」",
)

# ---------------------------------------------------------------- 快捷指令
st.caption("常用需求 · 点击填入输入框（可修改后再执行），鼠标悬停可看完整指令")
_tabs = st.tabs([g["label"] for g in QUICK_GROUPS])
for _tab, _group in zip(_tabs, QUICK_GROUPS):
    with _tab:
        _items = _group["items"]
        _cols = st.columns(min(len(_items), 4))
        for _i, _item in enumerate(_items):
            with _cols[_i % len(_cols)]:
                st.button(
                    _item["label"],
                    help=_item["text"],
                    on_click=_fill_query,
                    args=(_item["text"],),
                    key=f"qk_{_group['key']}_{_i}",
                    use_container_width=True,
                )

run_btn = st.button("🚀 执行任务", type="primary", use_container_width=True)

# ---------------------------------------------------------------- 多文档操作
with st.expander("📎 多份文档操作（合并 / 拆分 / 提取表格 / 对比）", expanded=False):
    st.caption("这些操作要指定多份文件，直接在这里选好点按钮即可，不用打字。")
    _all_docs = list_docs()
    if not _all_docs:
        st.info("还没有可用文档，请先在左侧上传。")
    else:
        # 默认只列样例与上传的原件：输出目录里是历史产物，全列出来会把可选列表撑爆、
        # 也容易误选（比如把刚拆出来的碎片又合并回去）。需要时勾一下即可。
        _src_docs = {k: v for k, v in _all_docs.items() if not k.startswith("[输出]")}
        _show_out = st.checkbox("也包含输出目录的产物", value=False, key="multi_show_out")
        _pool = _all_docs if _show_out else _src_docs
        _labels = list(_pool.keys())
        _picked = st.multiselect(
            "选择参与操作的文件（合并按这里的先后顺序拼接）",
            _labels,
            key="multi_pick",
            help="拆分和提取表格只用第一份；合并至少两份；对比恰好两份。",
        )
        _paths = [str(_pool[x]) for x in _picked]
        # 供「合并 / 对比」类对话指令复用：这些动作天然需要多份文件，
        # 让用户在面板选好，再在输入框里说一句「合并这些文档」就能执行。
        st.session_state.multi_paths = _paths

        _m1, _m2, _m3, _m4 = st.columns(4)
        _do_merge = _m1.button("🔗 合并", use_container_width=True,
                               disabled=len(_paths) < 2,
                               help="把多份文档按顺序合并成一份，可选分页")
        _do_split = _m2.button("✂️ 按标题拆分", use_container_width=True,
                               disabled=not _paths,
                               help="按标题层级把第一份文档拆成多份")
        _do_tables = _m3.button("📊 提取表格", use_container_width=True,
                                disabled=not _paths,
                                help="把第一份文档里的表格导出成 Excel")
        _do_diff = _m4.button("🔍 对比差异", use_container_width=True,
                              disabled=len(_paths) != 2,
                              help="对比两份文档，输出新增 / 删除 / 修改")

        _o1, _o2, _o3 = st.columns(3)
        _split_level = _o1.selectbox("拆分依据", [1, 2, 3], index=0,
                                     format_func=lambda v: f"按 {v} 级标题",
                                     key="split_level")
        _split_anchor = _o2.text_input("或从某个标题处切开", key="split_anchor",
                                       placeholder="留空则按层级全拆")
        _page_break = _o3.checkbox("合并时分页", value=True, key="merge_brk")

        _job = None
        if _do_merge:
            _job = ("merge", "合并多份文档", {
                "path": _paths[0], "action": "merge",
                "other_paths": "|".join(_paths[1:]),
                "add_page_break": _page_break,
            })
        elif _do_split:
            _job = ("split", "按标题拆分文档", {
                "path": _paths[0], "action": "split",
                "level": int(_split_level), "anchor": _split_anchor.strip(),
            })
        elif _do_tables:
            _job = ("tables", "提取表格到 Excel", {
                "path": _paths[0], "action": "extract_tables",
            })
        elif _do_diff:
            _job = ("diff", "对比两版文档", {
                "path": _paths[0], "action": "diff", "other_paths": _paths[1],
            })

        if _job:
            _sid = st.session_state.session_id
            set_session_id(_sid)
            _args = _job[2]
            with st.spinner("正在处理…"):
                try:
                    _res = dispatch("structure_document", **_args)
                except Exception as _e:  # 兜底：不让界面白屏
                    st.error(f"执行失败：{type(_e).__name__}: {_e}")
                    _res = None
            if _res is not None:
                st.session_state.last_result = _direct_result(_res, _job[1], _args)

# ---------------------------------------------------------------- 执行
if run_btn:
    if not task or not task.strip():
        st.error("请先输入任务描述（空输入已被输入校验拦截）")
    else:
        sid = st.session_state.session_id
        set_session_id(sid)
        src_path = st.session_state.current_file or ""
        # 合并 / 对比这类动作天然要两份以上文件：如果用户已经在面板里勾好了，
        # 就直接用那批文件，不用再回面板点一次按钮。单个文件的操作不受影响。
        _multi = st.session_state.get("multi_paths") or []
        if len(_multi) >= 2 and re.search(r"合并|拆分|对比|比对|差异|提取表格", task):
            src_path = "|".join(_multi)
        hash_before = ""
        if src_path:
            try:
                hash_before = sha256_of(src_path)
            except OSError:
                hash_before = ""

        with st.spinner("正在执行…"):
            t0 = time.time()
            result = execute(task.strip(), src_path, sid, backend=mode)
            elapsed = time.time() - t0

        result["_task"] = task.strip()
        result["_elapsed"] = elapsed
        st.session_state.last_result = result

        # 原文件保护自检
        guard_note = None
        if src_path and hash_before:
            try:
                hash_after = sha256_of(src_path)
                guard_note = (hash_after == hash_before, hash_before, hash_after)
            except OSError:
                pass
        result["_guard"] = guard_note

        if result.get("ok"):
            st.success(f"✅ 已完成（用时 {elapsed:.1f} 秒）")
        else:
            st.error("这次没有完成，可以按下面的建议调整说法再试一次。")

# ---------------------------------------------------------------- 结果展示
res = st.session_state.last_result
if res:
    st.divider()
    left, right = st.columns([2, 1])

    with left:
        st.subheader("结果")
        if res.get("reply"):
            st.write(res["reply"])
        if not res.get("ok"):
            st.error(res.get("error") or "任务失败")
            if res.get("suggestion"):
                st.info(f"建议：{res['suggestion']}")

        steps = res.get("steps") or []
        with st.expander("🔍 查看处理过程（验收 / 排查时展开）", expanded=False):
            if not steps:
                st.info("本次任务没有产生工具调用。")
            for i, s in enumerate(steps, 1):
                r = s.get("result", {}) or {}
                title = f"第 {i} 步 · {s.get('tool')} · " + ("成功" if r.get("ok") else f"失败[{r.get('code')}]")
                with st.expander(title, expanded=False):
                    st.markdown("**入参**")
                    st.json(s.get("args", {}), expanded=False)
                    st.markdown("**返回**")
                    st.write(r.get("message", ""))
                    if r.get("changes"):
                        st.markdown(f"**变更明细（{len(r['changes'])} 项）**")
                        shown = [{k: (str(v)[:120]) for k, v in c.items()} for c in r["changes"][:10]]
                        st.table(shown)
                    data = r.get("data") or {}
                    if data.get("guard"):
                        g = data["guard"]
                        st.caption(f"源文件哈希：{g.get('sha256_before','')} → {g.get('sha256_after','')} "
                                   f"({'未变化 ✓' if g.get('unchanged') else '已变化 ✗'})")
                    if data.get("outline"):
                        st.markdown("**文档大纲**")
                        st.dataframe(data["outline"], use_container_width=True, hide_index=True)
                    r["_uid"] = f"{i}_{id(res)}"
                    render_file_download(r)

            gn = res.get("_guard")
            if gn:
                ok, before, after = gn
                st.markdown("**源文件保护自检**")
                if ok:
                    st.success("源文件哈希未变化 ✓")
                else:
                    st.error("源文件哈希发生变化 ✗")
                st.caption(f"{before[:16]}… → {after[:16]}…")

    with right:
        st.subheader("输出文件")
        outs = list(res.get("outputs") or [])
        # 拆分会产生多份文件，它们放在 data["files"] 里，这里一并列出
        for s in res.get("steps") or []:
            d = (s.get("result") or {}).get("data") or {}
            for f in (d.get("files") or []):
                if f not in outs:
                    outs.append(f)
        if outs:
            shown = 0
            for i, o in enumerate(outs):
                p = Path(o)
                if p.exists():
                    render_file_download({"output_file": o, "_uid": f"out{i}"})
                    shown += 1
            if shown > 1:
                st.caption(f"共 {shown} 个文件，可逐个下载")
        else:
            st.caption("无输出文件")

# ---------------------------------------------------------------- 调用记录
st.divider()
with st.expander("📋 调用记录（验收 / 排查问题时展开）", expanded=False):
    records = read_calls(st.session_state.session_id, limit=200)
    if records:
        import pandas as pd
        df = pd.DataFrame([{
            "时间": r.get("ts", "").split("T")[-1][:12],
            "层": r.get("layer"),
            "名称": r.get("name"),
            "动作": r.get("action"),
            "后端": r.get("backend") or "-",
            "结果": "✓" if r.get("ok") else f"✗ {r.get('code')}",
            "耗时ms": r.get("elapsed_ms"),
            "说明": (r.get("message") or "")[:60],
            "输出": Path(r["output_file"]).name if r.get("output_file") else "",
        } for r in records])
        st.dataframe(df, use_container_width=True, hide_index=True)
        st.download_button(
            "⬇ 导出调用记录 CSV",
            df.to_csv(index=False).encode("utf-8-sig"),
            file_name=f"调用记录_{st.session_state.session_id}.csv",
            mime="text/csv",
        )
    else:
        st.caption("暂无调用记录，执行任务后这里会出现每一层调用的明细。")
