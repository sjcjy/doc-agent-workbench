from __future__ import annotations

"""全局配置与路径常量。

所有路径都以项目根目录为基准推导，方便整体迁移（后续转小程序时用同一套核心层，
只需换掉上层界面，路径规则保持一致）。
"""

import os
from pathlib import Path

from dotenv import load_dotenv

# ---------------------------------------------------------------- 目录
ROOT_DIR = Path(__file__).resolve().parent.parent

DATA_DIR = ROOT_DIR / "data"
UPLOAD_DIR = DATA_DIR / "uploads"      # 用户上传的原文件（只读）
WORK_DIR = DATA_DIR / "work"           # 工具中间产物（同样只读看待）
OUTPUT_DIR = DATA_DIR / "outputs"      # 唯一命名的输出文件
LOG_DIR = ROOT_DIR / "logs"
SAMPLE_DIR = ROOT_DIR / "samples"
EVIDENCE_DIR = ROOT_DIR / "evidence"

for _d in (UPLOAD_DIR, WORK_DIR, OUTPUT_DIR, LOG_DIR, SAMPLE_DIR, EVIDENCE_DIR):
    _d.mkdir(parents=True, exist_ok=True)

CALL_LOG_PATH = LOG_DIR / "calls.jsonl"

# ---------------------------------------------------------------- 环境变量
load_dotenv(ROOT_DIR / ".env", override=False)

# 大模型
ZHIPU_API_KEY = os.getenv("ZHIPU_API_KEY", "").strip()
ZHIPU_BASE_URL = os.getenv("ZHIPU_BASE_URL", "https://open.bigmodel.cn/api/paas/v4").strip()
ZHIPU_MODEL = os.getenv("ZHIPU_MODEL", "glm-4-flash").strip()
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "").strip()
OPENAI_BASE_URL = os.getenv("OPENAI_BASE_URL", "").strip()
OPENAI_MODEL = os.getenv("OPENAI_MODEL", "gpt-4o-mini").strip()

# auto：有 Key 走大模型，失败自动回退规则路由；llm：只用大模型；rule：只用规则路由
AGENT_BACKEND = os.getenv("AGENT_BACKEND", "auto").strip().lower()
# zhipu-openai（智谱 OpenAI 兼容端点，推荐）| zhipu-native | openai
MODEL_PROVIDER = os.getenv("MODEL_PROVIDER", "zhipu-openai").strip().lower()

MODEL_TIMEOUT_SEC = int(os.getenv("MODEL_TIMEOUT_SEC", "60"))
MODEL_MAX_RETRIES = int(os.getenv("MODEL_MAX_RETRIES", "3"))
MODEL_TEMPERATURE = float(os.getenv("MODEL_TEMPERATURE", "0.2"))
AGENT_MAX_STEPS = int(os.getenv("AGENT_MAX_STEPS", "8"))

# ---------------------------------------------------------------- 文件规则
SUPPORTED_DOCX_EXT = {".docx"}
SUPPORTED_TEXT_EXT = {".txt", ".md", ".markdown"}
SUPPORTED_INPUT_EXT = SUPPORTED_DOCX_EXT | SUPPORTED_TEXT_EXT
SUPPORTED_UPLOAD_EXT = SUPPORTED_INPUT_EXT | {".pdf"}

MAX_UPLOAD_MB = int(os.getenv("MAX_UPLOAD_MB", "20"))
MAX_UPLOAD_BYTES = MAX_UPLOAD_MB * 1024 * 1024

# 允许读取的目录白名单（防路径穿越）。输出目录外的写入一律拦截。
READABLE_ROOTS = (UPLOAD_DIR, WORK_DIR, OUTPUT_DIR, SAMPLE_DIR)
WRITABLE_ROOTS = (OUTPUT_DIR, WORK_DIR)

# 中文默认字体与常用格式预设
DEFAULT_CN_FONT = os.getenv("DEFAULT_CN_FONT", "宋体")
DEFAULT_EN_FONT = os.getenv("DEFAULT_EN_FONT", "Times New Roman")
DEFAULT_BODY_SIZE_PT = 12.0
DEFAULT_LINE_SPACING = 1.5


def ensure_dir(path: str | Path) -> Path:
    p = Path(path)
    p.mkdir(parents=True, exist_ok=True)
    return p


def has_model_key() -> bool:
    """判断当前是否具备调用大模型的凭据。"""
    if MODEL_PROVIDER.startswith("zhipu"):
        return bool(ZHIPU_API_KEY)
    return bool(OPENAI_API_KEY)
