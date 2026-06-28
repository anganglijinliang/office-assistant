# -*- coding: utf-8 -*-
"""
万能办公助手 - 工具模块 (utils.py)
辅助函数 + 库检测 + 进度对话框 + 错误日志
"""
import sys, os, json, time, hashlib, threading, warnings
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext

warnings.filterwarnings("ignore")

# ============================================================
# 1. 数据目录 + 库可用性检测
# ============================================================
DATA_DIR = Path.home() / ".office_assistant"
DATA_DIR.mkdir(exist_ok=True)
CONFIG_FILE = DATA_DIR / "config.json"

# PIL（图片处理）
_PIL_ERROR = ""
try:
    from PIL import Image, ImageDraw, ImageFont, ImageFilter
    PIL_AVAILABLE = True
except ImportError as e:
    PIL_AVAILABLE = False
    _PIL_ERROR = str(e)

# openpyxl（Excel）
_OPENPYL_ERROR = ""
try:
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    OPENPYL_AVAILABLE = True
except ImportError as e:
    OPENPYL_AVAILABLE = False
    _OPENPYL_ERROR = str(e)

# python-docx（Word）
_DOCX_ERROR = ""
try:
    from docx import Document
    DOCX_AVAILABLE = True
except ImportError as e:
    DOCX_AVAILABLE = False
    _DOCX_ERROR = str(e)

# PyPDF2（PDF基础）
_PDF_ERROR = ""
try:
    import PyPDF2
    _ = PyPDF2.PdfReader; _ = PyPDF2.PdfWriter; _ = PyPDF2.PdfMerger
    PDF_AVAILABLE = True
except Exception as _pdf_err:
    PDF_AVAILABLE = False
    _PDF_ERROR = str(_pdf_err)

# pdfplumber（PDF增强）
try:
    import pdfplumber
    if not PDF_AVAILABLE:
        PDF_AVAILABLE = True
except Exception:
    pass

# reportlab（PDF生成）
_REPORTLAB_ERROR = ""
try:
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table as RLTable
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.lib import colors
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.pdfgen import canvas as rl_canvas
    REPORTLAB_AVAILABLE = True
except ImportError as e:
    REPORTLAB_AVAILABLE = False
    _REPORTLAB_ERROR = str(e)

# pypdfium2（PDF→图片）
_PYPDFIUM_ERROR = ""
try:
    import pypdfium2
    PYPDFIUM_AVAILABLE = True
except ImportError as e:
    PYPDFIUM_AVAILABLE = False
    _PYPDFIUM_ERROR = str(e)


# ============================================================
# 2. 配置管理
# ============================================================
def load_config():
    if CONFIG_FILE.exists():
        try:
            return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}

def save_config(cfg):
    CONFIG_FILE.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")


# ============================================================
# 3. 辅助函数
# ============================================================
def safe_str(val):
    """安全转字符串（None→''）"""
    return str(val) if val is not None else ""

def get_font(size):
    """获取中文字体"""
    fonts_to_try = ["C:/Windows/Fonts/simhei.ttf", "C:/Windows/Fonts/msyh.ttc",
                    "C:/Windows/Fonts/simsun.ttc", "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc"]
    for fp in fonts_to_try:
        if Path(fp).exists():
            try:
                return ImageFont.truetype(fp, size)
            except Exception:
                continue
    return ImageFont.load_default()

def _log_error(context, err):
    """记录错误到crash.log（替代沉默except:pass）"""
    try:
        log_path = DATA_DIR / "crash.log"
        with open(str(log_path), "a", encoding="utf-8") as f:
            f.write(f"\n[{datetime.now()}] [{context}] {err}\n")
    except Exception:
        pass  # 无可救药时沉默

def safe_cond_check(val, cond):
    """安全条件比较（替代eval）"""
    try:
        nv = float(val)
        cd = cond.strip()
        if cd.startswith(">="): return nv >= float(cd[2:].strip())
        elif cd.startswith("<="): return nv <= float(cd[2:].strip())
        elif cd.startswith("!="): return nv != float(cd[2:].strip())
        elif cd.startswith("=="): return nv == float(cd[2:].strip())
        elif cd.startswith(">"): return nv > float(cd[1:].strip())
        elif cd.startswith("<"): return nv < float(cd[1:].strip())
        else: return False
    except (ValueError, TypeError):
        return False

def get_lib_status_text():
    """状态栏显示的可用库文本"""
    libs = []
    if PIL_AVAILABLE: libs.append("PIL")
    if OPENPYL_AVAILABLE: libs.append("Openpyxl")
    if DOCX_AVAILABLE: libs.append("Docx")
    if PDF_AVAILABLE: libs.append("PDF")
    return "+".join(libs)

def write_diagnostic_log():
    """写入启动诊断日志"""
    log_path = DATA_DIR / "load_diagnostic.log"
    try:
        with open(log_path, "w", encoding="utf-8") as f:
            f.write(f"=== 万能办公助手 启动诊断 ===\n")
            f.write(f"时间: {datetime.now()}\n")
            f.write(f"sys.frozen: {getattr(sys, 'frozen', False)}\n")
            f.write(f"sys._MEIPASS: {getattr(sys, '_MEIPASS', 'N/A')}\n")
            f.write(f"PIL: {PIL_AVAILABLE}\n")
            f.write(f"Openpyxl: {OPENPYL_AVAILABLE}\n")
            f.write(f"Docx: {DOCX_AVAILABLE}\n")
            f.write(f"PDF: {PDF_AVAILABLE}\n")
            f.write(f"Reportlab: {REPORTLAB_AVAILABLE}\n")
            f.write(f"Pypdfium2: {PYPDFIUM_AVAILABLE}\n")
    except Exception:
        pass


# ============================================================
# 4. 进度对话框
# ============================================================
class ProgressDialog:
    """统一进度对话框 - 进度条 + 实时日志 + 耗时统计"""
    def __init__(self, parent, title="处理中...", width=560, height=400):
        self.win = tk.Toplevel(parent)
        self.win.title(title)
        self.win.geometry(f"{width}x{height}")
        self.win.minsize(420, 260)
        self.win.transient(parent)
        self.win.grab_set()
        self.win.configure(bg="#FFFFFF")
        self.start_time = time.time()
        self._closed = False
        self._lock = threading.Lock()
        tk.Frame(self.win, bg="#4F46E5", height=3).pack(fill=tk.X)
        bar_frame = tk.Frame(self.win, bg="#FFFFFF")
        bar_frame.pack(fill=tk.X, padx=24, pady=(18, 8))
        self.progress = ttk.Progressbar(bar_frame, length=500, mode="indeterminate")
        self.progress.pack(fill=tk.X)
        self.progress.start(10)
        self.status_var = tk.StringVar(value="正在处理...")
        tk.Label(self.win, textvariable=self.status_var, font=("微软雅黑", 10),
                fg="#1E293B", bg="#FFFFFF", anchor="w").pack(fill=tk.X, padx=24, pady=(0, 8))
        self.log = scrolledtext.ScrolledText(self.win, height=12, font=("Consolas", 9),
                                              bg="#0F172A", fg="#E2E8F0", insertbackground="white")
        self.log.pack(fill=tk.BOTH, expand=True, padx=24, pady=(0, 16))

    def update(self, msg, progress=None):
        self.win.after(0, lambda: self._do_update(msg, progress))

    def _do_update(self, msg, progress):
        if self._closed: return
        self.status_var.set(msg if len(msg) < 80 else msg[:77] + "...")
        self.log.insert(tk.END, msg + "\n")
        self.log.see(tk.END)
        if progress is not None:
            self.progress.stop()
            self.progress.config(mode="determinate", value=progress * 100)

    def finish(self, ok=0, fail=0, extra=""):
        def _do_finish():
            if self._closed: return
            elapsed = time.time() - self.start_time
            summary = "\n" + "=" * 40 + "\n"
            summary += f"耗时: {elapsed:.1f}秒\n"
            summary += f"成功: {ok} | 失败: {fail}\n"
            if extra: summary += extra
            summary += "=" * 40 + "\n"
            self.log.insert(tk.END, summary)
            self.log.see(tk.END)
            self.progress.stop()
            self.progress.config(mode="determinate", value=100)
        self.win.after(0, _do_finish)

    def close(self):
        def _do_close():
            self._closed = True
            try: self.win.destroy()
            except tk.TclError: pass
        self.win.after(0, _do_close)


def show_result(title, ok, fail, extra=""):
    elapsed = extra if extra else ""
    lines = [f"OK: {ok}   FAIL: {fail}"]
    if ok + fail > 0: lines.append(f"Total: {ok + fail}")
    if elapsed: lines.append(f"Time: {elapsed}")
    messagebox.showinfo(title, "\n".join(lines))
