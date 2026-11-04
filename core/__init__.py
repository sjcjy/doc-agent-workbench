"""文档工作台智能体 - 核心层。

这一层不依赖任何界面框架（Streamlit / FastAPI / 未来的小程序都复用它），
职责边界：模型只负责理解与决策，文件操作全部由确定性工具完成。
"""

from core.models import ToolResult, code_text  # noqa: F401

__all__ = ["ToolResult", "code_text"]
__version__ = "1.0.0"
