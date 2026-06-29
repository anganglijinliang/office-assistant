# -*- coding: utf-8 -*-
"""万能办公助手 v6.2 — 工具函数模块（Utils）"""

import os
import sys
import json
import logging
import traceback
from pathlib import Path
from datetime import datetime
from typing import Optional, Any, Callable

import tkinter as tk
from tkinter import ttk

# ── 数据目录 ──────────────────────────────────────────────────────────────
DATA_DIR = Path.home() / '.office_assistant'
CONFIG_FILE = DATA_DIR / 'config.json'

# ── 库可用性检查 ──────────────────────────────────────────────────────────
DOCX_AVAILABLE = False
OPENPYL_AVAILABLE = False
PIL_AVAILABLE = False
PDF_AVAILABLE = False
REPORTLAB_AVAILABLE = False
PYPDFIUM_AVAILABLE = False
WORD_COM_AVAILABLE = False
_DOCX_ERROR = ""
_OPENPYL_ERROR = ""
_PIL_ERROR = ""
_PDF_ERROR = ""
_REPORTLAB_ERROR = ""
_PYPDFIUM_ERROR = ""

try:
    import docx  # noqa: F401
    DOCX_AVAILABLE = True
except ImportError as e:
    _DOCX_ERROR = str(e)

try:
    import openpyxl  # noqa: F401
    OPENPYL_AVAILABLE = True
except ImportError as e:
    _OPENPYL_ERROR = str(e)

try:
    from PIL import Image  # noqa: F401
    PIL_AVAILABLE = True
except ImportError as e:
    _PIL_ERROR = str(e)

# PDF: try PyMuPDF first, then PyPDF2
try:
    import fitz  # PyMuPDF  # noqa: F401
    PDF_AVAILABLE = True
except ImportError:
    try:
        import PyPDF2  # noqa: F401
        PDF_AVAILABLE = True
    except ImportError as e:
        _PDF_ERROR = str(e)

try:
    from reportlab.lib.pagesizes import A4  # noqa: F401
    REPORTLAB_AVAILABLE = True
except ImportError as e:
    _REPORTLAB_ERROR = str(e)

try:
    import pypdfium2  # noqa: F401
    PYPDFIUM_AVAILABLE = True
except ImportError as e:
    _PYPDFIUM_ERROR = str(e)

try:
    import win32com.client  # noqa: F401
    WORD_COM_AVAILABLE = True
except ImportError:
    pass


# ── 公共工具函数 ──────────────────────────────────────────────────────────

def safe_str(obj: Any, default: str = '') -> str:
    """将任意对象安全转换为字符串，转换失败时返回 default。"""
    try:
        return str(obj)
    except Exception:
        return default


def esc_xml(text: str) -> str:
    """XML 转义，只转义 & < >（不转义引号）。"""
    text = text.replace('&', '&amp;')
    text = text.replace('<', '&lt;')
    text = text.replace('>', '&gt;')
    return text


def get_font(size: int = 10) -> tuple:
    """返回 Tkinter 可用的中文字体元组 (family, size)。
    优先使用已注册的中文字体，回退到 TkDefaultFont。"""
    name = get_chinese_font_name()
    return (name, size)


def register_chinese_font() -> None:
    """注册系统常见中文字体供 reportlab / PIL 使用（reportlab 环境）。"""
    try:
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
    except ImportError:
        return

    # 常见中文字体路径
    candidates = [
        # Windows
        'C:/Windows/Fonts/msyh.ttc',       # 微软雅黑
        'C:/Windows/Fonts/simsun.ttc',      # 宋体
        'C:/Windows/Fonts/simhei.ttf',      # 黑体
        'C:/Windows/Fonts/simkai.ttf',      # 楷体
        'C:/Windows/Fonts/fangsong.ttf',    # 仿宋
        # macOS / Linux 常见路径
        '/System/Library/Fonts/PingFang.ttc',
        '/System/Library/Fonts/STHeiti Light.ttc',
        '/usr/share/fonts/truetype/wqy/wqy-microhei.ttc',
        '/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf',
    ]
    for path in candidates:
        fp = Path(path)
        if fp.exists():
            try:
                pdfmetrics.registerFont(TTFont(fp.stem, str(fp)))
            except Exception:
                continue


def get_chinese_font_name() -> str:
    """获取当前系统可用中文字体名称，优先返回 Tkinter 可用字体。"""
    families = []
    try:
        families = list(tk.font.families())
    except Exception:
        pass

    # 按优先级排列候选字体
    preferred = ['微软雅黑', 'Microsoft YaHei', 'SimSun', 'SimHei',
                 'PingFang SC', 'STHeiti', 'Noto Sans CJK SC',
                 'WenQuanYi Micro Hei', 'WenQuanYi Zen Hei',
                 'Droid Sans Fallback', 'Source Han Sans CN']
    for name in preferred:
        if name in families:
            return name

    # Windows 通用回退
    for f in families:
        lower = f.lower()
        if 'yahei' in lower or '雅黑' in f:
            return f
        if 'simsun' in lower or '宋体' in f:
            return f
        if 'simhei' in lower or '黑体' in f:
            return f
    return 'TkDefaultFont'


def check_ocr_available() -> bool:
    """检查 OCR（pytesseract + tesseract）是否可用。"""
    try:
        import pytesseract  # noqa: F401
        from PIL import Image  # noqa: F401
        # 尝试调用 tesseract 版本检测
        try:
            pytesseract.get_tesseract_version()
            return True
        except Exception:
            return False
    except ImportError:
        return False


def get_ocr_install_guide() -> str:
    """返回 OCR 安装指南文本"""
    return (
        "OCR 识别需要安装 Tesseract OCR 引擎和 pytesseract 库：\n\n"
        "1. 安装 Tesseract 引擎:\n"
        "   下载: https://github.com/UB-Mannheim/tesseract/wiki\n"
        "   安装时勾选「简体中文(chinese-simplified)」语言包\n\n"
        "2. 安装 pytesseract:\n"
        "   pip install pytesseract\n\n"
        "3. 如果 Tesseract 不在系统 PATH 中:\n"
        "   在代码中设置: pytesseract.pytesseract.tesseract_cmd = r'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'"
    )


def safe_cond_check(val: Any, cond: Optional[str] = None) -> bool:
    """带条件检查的值验证，安全处理异常。

    参数:
        val: 待检查的值
        cond: 条件字符串，如 '>0', '==True', 'is None' 等

    返回:
        bool 值
    """
    if cond is None:
        return bool(val)
    try:
        # 安全构建表达式
        allowed_ops = {'>', '<', '==', '!=', '>=', '<=', 'is', 'in', 'not'}
        cond = cond.strip()
        # 使用 eval 但限制在安全范围内
        result = eval(f"val {cond}", {'__builtins__': {}}, {'val': val})
        return bool(result)
    except Exception:
        return False


def get_lib_status_text() -> str:
    """返回各依赖库的安装状态文本，用于诊断显示。"""
    lines = [
        "━" * 40,
        "  库依赖状态",
        "━" * 40,
        f"  python-docx     : {'✅' if DOCX_AVAILABLE else '❌'}  {'已安装' if DOCX_AVAILABLE else '未安装'}",
        f"  openpyxl        : {'✅' if OPENPYL_AVAILABLE else '❌'}  {'已安装' if OPENPYL_AVAILABLE else '未安装'}",
        f"  Pillow (PIL)    : {'✅' if PIL_AVAILABLE else '❌'}  {'已安装' if PIL_AVAILABLE else '未安装'}",
        f"  PyMuPDF (fitz)  : {'✅' if PDF_AVAILABLE else '❌'}  {'已安装' if PDF_AVAILABLE else '未安装'}",
        f"  reportlab       : {'✅' if REPORTLAB_AVAILABLE else '❌'}  {'已安装' if REPORTLAB_AVAILABLE else '未安装'}",
        f"  pypdfium2       : {'✅' if PYPDFIUM_AVAILABLE else '❌'}  {'已安装' if PYPDFIUM_AVAILABLE else '未安装'}",
        f"  win32com        : {'✅' if WORD_COM_AVAILABLE else '❌'}  {'已安装' if WORD_COM_AVAILABLE else '未安装'}",
        f"  pytesseract     : {'✅' if check_ocr_available() else '❌'}  {'已安装' if check_ocr_available() else '未安装'}",
        "━" * 40,
    ]
    return '\n'.join(lines)


def write_diagnostic_log(message: str, logger: Optional[logging.Logger] = None) -> None:
    """写入诊断日志到文件和控制台。

    日志保存在 DATA_DIR/diagnostics/ 目录下，按日期命名。

    参数:
        message: 日志消息
        logger: 可选 logging.Logger 实例，同时输出到该 logger
    """
    try:
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        diag_dir = DATA_DIR / 'diagnostics'
        diag_dir.mkdir(exist_ok=True)
        log_file = diag_dir / f"diag_{datetime.now().strftime('%Y%m%d')}.log"
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        line = f"[{timestamp}] {message}"
        with open(log_file, 'a', encoding='utf-8') as f:
            f.write(line + '\n')
        if logger:
            logger.info(message)
        else:
            print(line, file=sys.stderr)
    except Exception:
        pass  # 静默失败，避免连锁异常


def load_config() -> dict:
    """加载配置文件 config.json。

    返回:
        dict 配置字典，文件不存在或解析失败时返回空 dict
    """
    try:
        if CONFIG_FILE.exists():
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
    except (json.JSONDecodeError, OSError) as e:
        write_diagnostic_log(f"加载配置文件失败: {e}")
    return {}


def save_config(config: dict) -> bool:
    """保存配置到 config.json。

    参数:
        config: 配置字典

    返回:
        bool 是否保存成功
    """
    try:
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        return True
    except (OSError, TypeError) as e:
        write_diagnostic_log(f"保存配置文件失败: {e}")
        return False


# ── 进度对话框 ────────────────────────────────────────────────────────────

class ProgressDialog:
    """模态进度对话框，带进度条和状态信息。

    用法:
        dlg = ProgressDialog(parent, "正在处理...", "准备就绪")
        try:
            dlg.update("处理中...", 50)
            # ... 执行任务
            dlg.finish(ok=True, extra="完成！")
        except Exception as e:
            dlg.finish(ok=False, fail=str(e))
        finally:
            dlg.close()
    """

    def __init__(self, parent: Optional[tk.Widget] = None,
                 title: str = "处理中",
                 msg: str = "请稍候...",
                 progress: int = 0):
        """初始化进度对话框。

        参数:
            parent: 父窗口，None 则无父窗口
            title: 窗口标题
            msg: 初始状态文本
            progress: 初始进度值 (0-100)
        """
        self._window: Optional[tk.Toplevel] = None
        self._label: Optional[ttk.Label] = None
        self._progress_bar: Optional[ttk.Progressbar] = None
        self._status_var: tk.StringVar = tk.StringVar(value=msg)
        self._result: Optional[bool] = None

        try:
            win = tk.Toplevel(parent) if parent else tk.Toplevel()
            win.title(title)
            win.resizable(False, False)
            win.transient(parent) if parent else None
            win.grab_set()  # 模态

            # 窗口居中
            win.update_idletasks()
            w, h = 420, 140
            sw = win.winfo_screenwidth()
            sh = win.winfo_screenheight()
            x = (sw - w) // 2
            y = (sh - h) // 2
            win.geometry(f"{w}x{h}+{x}+{y}")

            # 布局
            frame = ttk.Frame(win, padding=16)
            frame.pack(fill=tk.BOTH, expand=True)

            self._label = ttk.Label(frame, textvariable=self._status_var,
                                    wraplength=380, anchor=tk.W)
            self._label.pack(fill=tk.X, pady=(0, 12))

            self._progress_bar = ttk.Progressbar(
                frame, orient=tk.HORIZONTAL, length=380,
                mode='determinate', maximum=100
            )
            self._progress_bar.pack(pady=(0, 8))
            self._progress_bar['value'] = max(0, min(progress, 100))

            btn_frame = ttk.Frame(frame)
            btn_frame.pack(fill=tk.X, pady=(4, 0))
            self._close_btn = ttk.Button(
                btn_frame, text="关闭", command=self._on_close
            )
            self._close_btn.pack(side=tk.RIGHT)

            self._window = win
            self._window.protocol("WM_DELETE_WINDOW", self._on_close)
            self._window.update()
        except Exception as e:
            write_diagnostic_log(f"ProgressDialog 初始化失败: {e}")
            self._window = None

    def update(self, msg: str, progress: int) -> None:
        """更新状态文本和进度条。

        参数:
            msg: 新状态文本
            progress: 进度值 (0-100)
        """
        if self._window is None:
            return
        try:
            self._status_var.set(msg)
            self._progress_bar['value'] = max(0, min(progress, 100))
            self._window.update()
        except Exception:
            pass

    def finish(self, ok: bool = True,
               fail: str = "",
               extra: str = "") -> None:
        """标记任务完成，更新最终状态。

        参数:
            ok: 是否成功
            fail: 失败时的错误信息
            extra: 额外提示信息
        """
        if self._window is None:
            return
        self._result = ok
        try:
            if ok:
                msg = "✅ 操作完成"
                if extra:
                    msg += f" — {extra}"
            else:
                msg = f"❌ 操作失败"
                if fail:
                    msg += f" — {fail}"
                if extra:
                    msg += f" ({extra})"
            self._status_var.set(msg)
            self._progress_bar['value'] = 100 if ok else 0
            self._close_btn.config(state=tk.NORMAL)
            self._window.update()
        except Exception:
            pass

    def close(self) -> None:
        """关闭对话框并释放资源。"""
        if self._window is None:
            return
        try:
            self._window.grab_release()
            self._window.destroy()
        except Exception:
            pass
        self._window = None
        self._label = None
        self._progress_bar = None

    def _on_close(self) -> None:
        """内部关闭回调。"""
        self.close()
