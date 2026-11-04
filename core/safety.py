from __future__ import annotations

"""安全层：原文件保护、路径校验、输出唯一命名、哈希快照。

计划书里最硬的一条是「全程不覆盖原文件」「原文件哈希不变」，这一层就是它的实现：
1. 所有输入经 validate_input_file 校验后才允许读取；
2. 工具永远不写回输入路径，输出必须由 build_output_path 生成唯一新文件名；
3. 任务前后对输入文件做 sha256 快照，任何变化都会被 hash 校验发现。
"""

import hashlib
import re
import secrets
import zipfile
from datetime import datetime
from pathlib import Path

from core import config
from core.models import (
    CODE_FILE_CORRUPTED,
    CODE_FILE_NOT_FOUND,
    CODE_INVALID_INPUT,
    CODE_SAFETY_BLOCKED,
    CODE_UNSUPPORTED_FORMAT,
    ToolResult,
)


# ---------------------------------------------------------------- 基础工具
def sha256_of(path: str | Path, chunk: int = 1 << 20) -> str:
    """计算文件 sha256，用于原文件保护校验。"""
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for blk in iter(lambda: f.read(chunk), b""):
            h.update(blk)
    return h.hexdigest()


def _within(path: Path, roots: tuple[Path, ...]) -> bool:
    """判断路径是否落在允许的目录白名单内（解析 .. 后比较，防路径穿越）。"""
    try:
        resolved = path.resolve()
    except OSError:
        return False
    for root in roots:
        try:
            resolved.relative_to(root.resolve())
            return True
        except ValueError:
            continue
    return False


def looks_like_docx(path: Path) -> bool:
    """DOCX 本质是 zip。用 zipfile 快速判断是否可解析，避免半程崩溃。

    计划书里的失败情形「损坏文件」就靠这里拦下来。
    """
    try:
        if not zipfile.is_zipfile(path):
            return False
        with zipfile.ZipFile(path) as z:
            names = set(z.namelist())
            return "word/document.xml" in names and "[Content_Types].xml" in names
    except Exception:
        return False


def looks_like_pdf(path: Path) -> bool:
    """PDF 文件头必须是 %PDF-。用前 8 字节快速判断，避免损坏文件进到解析器里崩。"""
    try:
        with open(path, "rb") as f:
            return f.read(8).startswith(b"%PDF-")
    except Exception:
        return False


# ---------------------------------------------------------------- 输入校验
def validate_input_file(raw_path: str | None, tool: str = "common", action: str = "read") -> tuple[Path | None, ToolResult | None]:
    """校验输入文件。返回 (合法路径, None) 或 (None, 失败结果)。

    覆盖的失败情形：空输入、路径越界、文件不存在、不支持的格式、损坏的 docx。
    """
    if raw_path is None or not str(raw_path).strip():
        return None, ToolResult.fail(tool, action, CODE_INVALID_INPUT, "文件路径为空，请先上传或选择一份文档")
    if len(str(raw_path)) > 500:
        return None, ToolResult.fail(tool, action, CODE_INVALID_INPUT, "文件路径过长，疑似异常输入")

    try:
        p = Path(str(raw_path).strip()).expanduser()
    except (ValueError, OSError):
        return None, ToolResult.fail(tool, action, CODE_INVALID_INPUT, "文件路径无法解析")

    if not _within(p, config.READABLE_ROOTS):
        return None, ToolResult.fail(
            tool, action, CODE_SAFETY_BLOCKED,
            f"路径不在允许的目录内（仅限 data/uploads、data/work、data/outputs、samples）：{p}"
        )
    if not p.exists():
        return None, ToolResult.fail(tool, action, CODE_FILE_NOT_FOUND, f"文件不存在：{p.name}")
    if not p.is_file():
        return None, ToolResult.fail(tool, action, CODE_INVALID_INPUT, f"不是一个文件：{p.name}")

    try:
        size = p.stat().st_size
    except OSError:
        return None, ToolResult.fail(tool, action, CODE_INVALID_INPUT, f"无法读取文件信息：{p.name}")
    if size == 0:
        return None, ToolResult.fail(tool, action, CODE_INVALID_INPUT, f"文件为空：{p.name}")
    if size > config.MAX_UPLOAD_BYTES:
        return None, ToolResult.fail(
            tool, action, CODE_UNSUPPORTED_FORMAT,
            f"文件过大（{size / 1048576:.1f}MB），上限 {config.MAX_UPLOAD_MB}MB"
        )

    suffix = p.suffix.lower()
    if suffix == ".docx":
        if not looks_like_docx(p):
            return None, ToolResult.fail(
                tool, action, CODE_FILE_CORRUPTED,
                f"文件无法解析为 DOCX（结构损坏或实为其它格式）：{p.name}"
            )
        return p, None
    if suffix in config.SUPPORTED_TEXT_EXT:
        return p, None
    if suffix == ".pdf":
        if not looks_like_pdf(p):
            return None, ToolResult.fail(
                tool, action, CODE_FILE_CORRUPTED,
                f"文件无法解析为 PDF（结构损坏或实为其它格式）：{p.name}"
            )
        return p, None

    return None, ToolResult.fail(tool, action, CODE_UNSUPPORTED_FORMAT, f"不支持的文件格式：{suffix or '(无扩展名)'}")


# ---------------------------------------------------------------- 输出命名
_SAFE_NAME = re.compile(r"[^\w\u4e00-\u9fa5.\-]+", re.UNICODE)


def sanitize_stem(text: str, max_len: int = 40, fallback: str = "doc") -> str:
    """把任意文本整理成可安全用作文件名的主干。"""
    s = _SAFE_NAME.sub("_", str(text or "").strip()).strip("._")
    return (s[:max_len] or fallback)


def build_output_path(input_path: str | Path | None, action: str, ext: str, hint: str = "") -> Path:
    """生成唯一的输出路径：原名__动作__时间戳_随机串.扩展名

    唯一命名是「不覆盖」的双保险：即使同一份文档、同一动作连跑两次，
    也不会互相覆盖，历史产物可追溯。
    """
    if input_path is not None:
        try:
            base = sanitize_stem(Path(input_path).stem)
        except Exception:
            base = ""
    else:
        base = ""
    if not base:
        base = sanitize_stem(hint, fallback="doc")

    act = sanitize_stem(action, max_len=16, fallback="out")
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    rand = secrets.token_hex(3)
    ext = ext if ext.startswith(".") else f".{ext}"
    out = config.OUTPUT_DIR / f"{base}__{act}__{ts}_{rand}{ext}"

    # 极低概率撞名时再补一位随机
    while out.exists():
        out = config.OUTPUT_DIR / f"{base}__{act}__{ts}_{secrets.token_hex(4)}{ext}"
    return out


def assert_not_input(input_path: str | Path, output_path: str | Path) -> None:
    """确保输出不会写回输入。违反即抛异常，由上层转成 SAFETY_BLOCKED。"""
    try:
        same = Path(input_path).resolve() == Path(output_path).resolve()
    except OSError:
        same = False
    if same:
        raise PermissionError("输出路径与输入路径相同，已拦截（原文件必须保持只读）")


# ---------------------------------------------------------------- 原文件保护
class InputGuard:
    """上下文管理器：进入时记录输入文件哈希，退出时校验是否被动过。

    用法：
        with InputGuard(src) as g:
            ...工具逻辑...
        if g.violated: 报告并计入失败
    """

    def __init__(self, input_path: str | Path):
        self.path = Path(input_path)
        self.before = ""
        self.after = ""
        self.violated = False
        self.error = ""

    def __enter__(self) -> "InputGuard":
        try:
            self.before = sha256_of(self.path)
        except Exception as e:  # 读不到就算了，不算破坏
            self.error = f"{type(e).__name__}: {e}"
        return self

    def __exit__(self, exc_type, exc, tb) -> bool:
        if self.before:
            try:
                self.after = sha256_of(self.path)
                self.violated = self.after != self.before
            except Exception as e:
                self.error = f"{type(e).__name__}: {e}"
        return False  # 不吞异常，交给上层处理

    def report(self) -> dict:
        return {
            "file": str(self.path),
            "sha256_before": self.before[:16] + "..." if self.before else "",
            "sha256_after": self.after[:16] + "..." if self.after else "",
            "unchanged": bool(self.before and self.after and not self.violated),
            "error": self.error,
        }


def save_upload(file_bytes: bytes, filename: str) -> tuple[Path | None, ToolResult | None]:
    """把上传内容落到 data/uploads，同名自动加后缀，绝不覆盖已存在的原文件。"""
    from core.models import ToolResult as _TR  # 局部导入避免循环

    if not file_bytes:
        return None, _TR.fail("upload", "save", CODE_INVALID_INPUT, "上传内容为空")
    name = Path(filename or "").name
    if not name:
        return None, _TR.fail("upload", "save", CODE_INVALID_INPUT, "文件名缺失")
    ext = Path(name).suffix.lower()
    if ext not in config.SUPPORTED_UPLOAD_EXT:
        return None, _TR.fail("upload", "save", CODE_UNSUPPORTED_FORMAT, f"不支持上传的类型：{ext or '(无扩展名)'}")
    if len(file_bytes) > config.MAX_UPLOAD_BYTES:
        return None, _TR.fail(
            "upload", "save", CODE_UNSUPPORTED_FORMAT,
            f"文件 {len(file_bytes) / 1048576:.1f}MB 超过上限 {config.MAX_UPLOAD_MB}MB"
        )

    target = config.UPLOAD_DIR / name
    if target.exists():
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        target = config.UPLOAD_DIR / f"{Path(name).stem}_{stamp}{ext}"
    target.write_bytes(file_bytes)
    return target, None
