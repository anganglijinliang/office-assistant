# -*- coding: utf-8 -*-
"""万能办公助手 — ClipboardTools"""
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
from pathlib import Path
import os, sys, json, time, shutil, hashlib, threading
from datetime import datetime
from collections import Counter


class ClipboardToolsMixin:
    """ClipboardTools — 所有方法通过self访问OfficeAssistant的属性"""

    def _show_clipboard_tools(self):
        self.clear_content()
        self._section_header("剪贴板历史", "捕获 · 搜索 · 回溯 · 一键粘贴")
        self._show_tips(
            "点击「捕获」→ 复制内容自动添加到底部 → 双击条目粘贴",
            "搜索框输入关键词可快速筛选历史记录，最多保存50条"
        )
        search_f = tk.Frame(self.content_frame, bg=self.colors['light'])
        search_f.pack(fill=tk.X, padx=20, pady=(5, 0))
        tk.Label(search_f, text="🔍 搜索:", font=("微软雅黑", 9), bg=self.colors['light'],
                fg=self.colors['gray']).pack(side=tk.LEFT, padx=(0, 5))
        clip_search = tk.Entry(search_f, font=("微软雅黑", 9), width=30)
        clip_search.pack(side=tk.LEFT)
        clip_search.bind("<KeyRelease>", lambda e: self._refresh_clip_list())
        btn_row = tk.Frame(self.content_frame, bg=self.colors['light'])
        btn_row.pack(fill=tk.X, padx=20, pady=5)
        tk.Button(btn_row, text="📋 捕获", command=self._capture_clip,
                 cursor="hand2", bg=self.colors['primary'], fg="white",
                 font=("微软雅黑", 10), width=8).pack(side=tk.LEFT, padx=(0, 8))
        tk.Button(btn_row, text="🗑 清空", command=self._clear_clip,
                 cursor="hand2", font=("微软雅黑", 10), width=6).pack(side=tk.LEFT)
        self.clipboard_listbox = tk.Listbox(self.content_frame, font=("Consolas", 10),
                                           bg="white", selectmode=tk.SINGLE)
        self.clipboard_listbox.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)
        self.clipboard_listbox.bind("<Double-Button-1>", lambda e: self._paste_clip())
        if not hasattr(self, 'clipboard_history') or not self.clipboard_history:
            self.clipboard_history = []
        self._refresh_clip_list()
        self._auto_clip_timer()

    def _auto_clip_timer(self):
        try:
            import tkinter as tk
            t = threading.Timer(2.0, self._capture_clip)
            t.daemon = True; t.start()
        except Exception:
            pass

    def _capture_clip(self):
        try:
            import pyperclip
            text = pyperclip.paste()
            if text and text.strip() and (not hasattr(self, '_last_clip') or text != self._last_clip):
                self._last_clip = text
                if not hasattr(self, 'clipboard_history') or self.clipboard_history is None:
                    self.clipboard_history = []
                if text not in self.clipboard_history:
                    self.clipboard_history.insert(0, text)
                    if len(self.clipboard_history) > 50:
                        self.clipboard_history = self.clipboard_history[:50]
                    try:
                        self.root.after(0, self._refresh_clip_list)
                    except Exception:
                        pass
        except Exception:
            pass

    def _clear_clip(self):
        self.clipboard_history = []
        self._refresh_clip_list()

    def _refresh_clip_list(self):
        if not self.clipboard_listbox:
            return
        self.clipboard_listbox.delete(0, tk.END)
        keyword = ""
        try:
            for w in self.content_frame.winfo_children():
                if isinstance(w, tk.Frame):
                    for c in w.winfo_children():
                        if isinstance(c, tk.Entry):
                            keyword = c.get().strip().lower()
                            break
        except Exception:
            pass
        for i, item in enumerate(self.clipboard_history):
            txt = item.replace("\n", " ")[:80]
            if keyword and keyword not in txt.lower():
                continue
            self.clipboard_listbox.insert(tk.END, f"  {txt}")

    def _paste_clip(self):
        sel = self.clipboard_listbox.curselection()
        if not sel:
            return
        try:
            import pyperclip
            text = self.clipboard_history[sel[0]]
            pyperclip.copy(text)
            self.set_status(f"已复制: {text[:40]}...")
        except Exception:
            pass
