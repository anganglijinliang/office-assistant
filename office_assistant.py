# -*- coding: utf-8 -*-
"""万能办公助手 v6.2 — 模块化架构"""
import sys, os, warnings
warnings.filterwarnings("ignore")

if not getattr(sys, 'frozen', False):
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# DPI感知
try:
    import ctypes
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

import re, json, time, threading, shutil, hashlib
from datetime import datetime, timedelta
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from collections import Counter
import base64, urllib.parse

# 工具模块导入
from utils import (
    DATA_DIR, CONFIG_FILE,
    PIL_AVAILABLE, OPENPYL_AVAILABLE, DOCX_AVAILABLE,
    PDF_AVAILABLE, REPORTLAB_AVAILABLE, PYPDFIUM_AVAILABLE,
    WORD_COM_AVAILABLE,
    load_config, save_config, safe_str, get_font,
    safe_cond_check as _safe_cond_check_util,
    ProgressDialog,
    write_diagnostic_log,
    register_chinese_font, get_chinese_font_name, esc_xml,
    check_ocr_available,
)
from lib_license import LicenseManager, LICENSE_FILE

# 模块化Mixin导入（仅导入实际存在的模块）
from modules.file_tools import FileToolsMixin
from modules.quick_tools import QuickToolsMixin
from modules.clipboard_tools import ClipboardToolsMixin
from modules.search_tools import SearchToolsMixin
from modules.calendar_tools import CalendarToolsMixin
from modules.system_tools import SystemToolsMixin
from modules.convert_tools import ConvertToolsMixin
from modules.pdf_tools import PdfToolsMixin
from modules.excel_tools import ExcelToolsMixin
from modules.doc_tools import DocToolsMixin
from modules.image_tools import ImageToolsMixin
from shared.ui_common import BaseUIMixin


class OfficeAssistant(
    BaseUIMixin,
    FileToolsMixin,
    QuickToolsMixin,
    ClipboardToolsMixin,
    SearchToolsMixin,
    CalendarToolsMixin,
    SystemToolsMixin,
    ConvertToolsMixin,
    PdfToolsMixin,
    ExcelToolsMixin,
    DocToolsMixin,
    ImageToolsMixin,
):
    def __init__(self):
        self.root = tk.Tk()
        self._dpi_scale = self._get_dpi_scale()
        base_w, base_h = 1320, 860
        self.config = load_config()
        saved_geo = self.config.get('window_geometry')
        if saved_geo:
            try: self.root.geometry(saved_geo)
            except: self._set_default_geo(base_w, base_h)
        else: self._set_default_geo(base_w, base_h)
        self.root.title("万能办公助手 v6.2 商业版")
        self.root.minsize(int(1050 * self._dpi_scale), int(720 * self._dpi_scale))
        if self._dpi_scale > 1.0:
            try:
                import tkinter.font as tkfont
                for fname in ("TkDefaultFont","TkTextFont","TkFixedFont"):
                    f = tkfont.nametofont(fname)
                    f.configure(size=int(f.cget("size") * self._dpi_scale))
            except: pass
        self.colors = {
            'primary': '#4F46E5', 'primary_hover': '#4338CA',
            'secondary': '#7C3AED', 'accent': '#06B6D4',
            'success': '#059669', 'warning': '#D97706',
            'danger': '#DC2626', 'dark': '#0F172A',
            'light': '#F8FAFC', 'white': '#FFFFFF',
            'gray': '#64748B', 'gray_light': '#94A3B8',
            'border': '#E2E8F0', 'hover_bg': '#F1F5F9',
            'nav_bg': '#0F172A', 'nav_item': '#334155',
            'nav_active': '#4F46E5', 'nav_text': '#94A3B8',
            'nav_text_active': '#FFFFFF', 'card_bg': '#FFFFFF',
            'success_bg': '#ECFDF5', 'warning_bg': '#FFFBEB',
        }
        self._tray_icon = None
        self._tray_visible = False
        self.clipboard_history = []
        self.clipboard_listbox = None
        self.memo_text = None
        self.todo_listbox = None
        self._todo_data = []
        self.progress_var = tk.StringVar(value="就绪")
        self.license_label = None
        self.license = LicenseManager(self)
        self._setup_ui()
        if not LICENSE_FILE.exists():
            # 首次启动：欢迎 + 激活提示 → 之后检查许可证
            self.root.after(500, lambda: self.license.show_welcome_dialog(
                on_close=lambda: self.root.after(200, self.license.check_and_show_license_warning)
            ))
        else:
            self.root.after(1000, self.license.check_and_show_license_warning)
        self.root.bind('<Control-f>', lambda e: self._focus_search())
        self.root.bind('<Control-q>', lambda e: self._quit_app())
        self.root.bind('<Escape>', lambda e: self._on_esc())

    def _set_default_geo(self, base_w, base_h):
        win_w, win_h = int(base_w * self._dpi_scale), int(base_h * self._dpi_scale)
        self.root.geometry(f"{win_w}x{win_h}")

    def _save_window_geometry(self):
        try:
            geo = self.root.geometry()
            self.config['window_geometry'] = geo
            save_config(self.config)
        except: pass

    def _focus_search(self):
        try:
            if hasattr(self, '_clip_search_var') and self._clip_search_var:
                self._navigate('clipboard')
        except: pass

    def _on_esc(self):
        try:
            tl = self.root.winfo_containing(self.root.winfo_pointerx(), self.root.winfo_pointery())
            while tl:
                if tl != self.root and hasattr(tl, 'destroy'):
                    tl.destroy(); return
                tl = tl.master if hasattr(tl, 'master') else None
        except: pass

    def _setup_ui(self):
        self.root.configure(bg=self.colors['light'])
        top_bar = tk.Frame(self.root, bg=self.colors['dark'], height=56)
        top_bar.pack(fill=tk.X, side=tk.TOP)
        top_bar.pack_propagate(False)
        tk.Label(top_bar, text="🧰  万能办公助手 v6.2", font=("微软雅黑", 15, "bold"),
                fg="white", bg=self.colors['dark']).pack(side=tk.LEFT, padx=20, pady=10)
        tk.Label(top_bar, text="商业版 · 万能办公助手 · 安全 · 高效 · 全能",
                font=("微软雅黑", 9), fg="#94A3B8", bg=self.colors['dark']).pack(side=tk.LEFT, padx=10, pady=10)
        tk.Button(top_bar, text="ℹ️ 关于", command=self.license.about_dialog, cursor="hand2",
                 font=("微软雅黑", 9), bd=0, padx=8, pady=2,
                 bg=self.colors['dark'], fg="#94A3B8",
                 activebackground='#1E293B', activeforeground='white').pack(side=tk.RIGHT, padx=(0, 4))
        tk.Button(top_bar, text="📌 托盘", command=self._minimize_to_tray, cursor="hand2",
                 font=("微软雅黑", 9), bd=0, padx=8, pady=2,
                 bg=self.colors['dark'], fg="#94A3B8",
                 activebackground='#1E293B', activeforeground='white').pack(side=tk.RIGHT, padx=4)
        tk.Button(top_bar, text="⚙ 卸载", command=self.license.uninstall_dialog, cursor="hand2",
                 font=("微软雅黑", 9), bd=0, padx=8, pady=2,
                 bg=self.colors['dark'], fg="#F87171",
                 activebackground='#1E293B', activeforeground='#FCA5A5').pack(side=tk.RIGHT, padx=4)
        nav = tk.Frame(self.root, bg=self.colors['nav_bg'], width=220)
        nav.pack(side=tk.LEFT, fill=tk.Y)
        nav.pack_propagate(False)
        self.nav_buttons = {}
        nav_items = [
            ("📁  文件处理", "file"),
            ("⚡  快捷工具", "quick"),
            ("📋  剪贴板", "clipboard"),
            ("🔍  内容搜索", "search"),
            ("📅  日程管理", "calendar"),
            ("📄  PDF工具", "pdf"),
            ("📊  Excel处理", "excel"),
            ("📝  文档处理", "doc"),
            ("🖼  图片工具", "image"),
            ("🔄  格式互转", "convert"),
        ]
        for text, key in nav_items:
            btn = tk.Button(nav, text=text, font=("微软雅黑", 11), anchor="w", padx=24, pady=13,
                           bg=self.colors['nav_bg'], fg="#CBD5E1", bd=0, cursor="hand2",
                           activebackground=self.colors['nav_active'], activeforeground="white",
                           relief="flat", command=lambda k=key: self._navigate(k))
            btn.pack(fill=tk.X)
            self.nav_buttons[key] = btn
        main_bg = tk.Frame(self.root, bg=self.colors['light'])
        main_bg.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.content_frame = tk.Frame(main_bg, bg=self.colors['light'])
        self.content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        bottom = tk.Frame(self.root, bg=self.colors['dark'], height=32)
        bottom.pack(side=tk.BOTTOM, fill=tk.X)
        bottom.pack_propagate(False)
        tk.Label(bottom, textvariable=self.progress_var, fg="#CBD5E1", bg=self.colors['dark'],
                font=("Consolas", 9), anchor="w", padx=15).pack(side=tk.LEFT)
        libs = []
        if PIL_AVAILABLE: libs.append("PIL")
        if OPENPYL_AVAILABLE: libs.append("Openpyxl")
        if DOCX_AVAILABLE: libs.append("Docx")
        if PDF_AVAILABLE: libs.append("PDF")
        lib_text = "+".join(libs) if libs else ""
        tk.Label(bottom, text=lib_text, fg="#475569", bg=self.colors['dark'],
                font=("Consolas",8), padx=15).pack(side=tk.RIGHT)
        self.license_label = tk.Label(bottom, text="", fg="#10B981", bg=self.colors['dark'],
                font=("Consolas", 9), padx=10)
        self.license_label.pack(side=tk.RIGHT)
        self._navigate("file")
        self.root.after(500, self._update_status_bar_license)
        self.root.after(3000, self._auto_clip_timer)
        self.root.protocol("WM_DELETE_WINDOW", self._quit_app)

    def _navigate(self, key):
        for k, btn in self.nav_buttons.items():
            if k == key: btn.config(bg=self.colors['nav_active'], fg="white")
            else: btn.config(bg=self.colors['nav_bg'], fg="#CBD5E1")
        routes = {
            "file": self._show_file_tools,
            "quick": self._show_quick_tools,
            "clipboard": self._show_clipboard_tools,
            "search": self._show_search_tools,
            "calendar": self._show_calendar_tools,
            "pdf": self._show_pdf_tools,
            "excel": self._show_excel_tools,
            "doc": self._show_doc_tools,
            "image": self._show_image_tools,
            "convert": self._show_convert_tools,
        }
        routes.get(key, self._show_file_tools)()

    def _show_placeholder(self):
        """占位页面 — 模块尚未实现。"""
        self.clear_content()
        self._section_header("功能开发中", "此模块正在开发中，敬请期待")
        self._show_tips("该功能将在后续版本中上线，感谢您的支持！",
                       "如有紧急需求，请使用「文件处理」「快捷工具」等现有功能。")
        # 获取当前导航键名
        current = None
        for k, btn in self.nav_buttons.items():
            if btn.cget('bg') == self.colors['nav_active']:
                current = k
                break
        label_map = {
            "pdf": "📄 PDF工具",
            "excel": "📊 Excel处理",
            "doc": "📝 文档处理",
            "image": "🖼 图片工具",
            "convert": "🔄 格式互转",
        }
        name = label_map.get(current, "此功能")
        tk.Label(
            self.content_frame, text=f"{name} 即将上线",
            font=("微软雅黑", 16), bg=self.colors['light'],
            fg=self.colors['gray'],
        ).pack(pady=60)

    def _update_status_bar_license(self):
        self.license.update_status_bar_license()

    def _safe_cond_check(self, val, cond):
        return _safe_cond_check_util(val, cond)

    def _about_dialog(self):
        self.license.about_dialog()

    def set_status(self, text: str) -> None:
        """更新状态栏文本。"""
        try:
            self.progress_var.set(str(text)[:60])
        except Exception:
            pass


    def run(self) -> None:
        """启动主事件循环。"""
        self.root.mainloop()


if __name__ == "__main__":
    app = OfficeAssistant()
    app.run()
