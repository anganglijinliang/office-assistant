# -*- coding: utf-8 -*-
"""万能办公助手 — CalendarTools（日程管理 + 待办事项 + 备忘录）"""
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
from pathlib import Path
import os, sys, json, time, shutil, hashlib, threading
from datetime import datetime, timedelta
from collections import Counter

from utils import DATA_DIR


class CalendarToolsMixin:
    """CalendarTools — 所有方法通过self访问OfficeAssistant的属性"""

    def _show_calendar_tools(self):
        self.clear_content()
        self._section_header("日程管理", "待办事项 · 定时提醒 · 备忘录 (自动保存)")
        self._show_tips(
            "左侧：输入事项后按回车添加 → 双击切换完成状态 → 点「⏰」设提醒",
            "右侧：编辑备忘录后点「💾 保存」，关闭软件后自动保留"
        )
        tf = DATA_DIR / "todo.json"; mf = DATA_DIR / "memo.txt"
        if tf.exists():
            try: self._todo_data = json.loads(tf.read_text(encoding="utf-8"))
            except: self._todo_data = []
        else: self._todo_data = []
        # 主容器
        main = tk.Frame(self.content_frame, bg=self.colors['light'])
        main.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)
        # 左栏：待办
        left = tk.Frame(main, bg='white', relief=tk.GROOVE, bd=1)
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 8))
        tk.Label(left, text="📋 待办事项", font=("微软雅黑", 13, "bold"),
                bg='white', fg=self.colors['dark']).pack(pady=10)
        inp_f = tk.Frame(left, bg='white'); inp_f.pack(fill=tk.X, padx=10)
        inp = tk.Entry(inp_f, font=("微软雅黑", 11), relief=tk.GROOVE, bd=1)
        inp.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=4)
        tk.Button(inp_f, text="➕ 添加", command=self._todo_add, cursor="hand2",
                 bg=self.colors['primary'], fg="white", font=("微软雅黑", 9), width=7).pack(side=tk.RIGHT, padx=(3,0))
        bf = tk.Frame(left, bg='white'); bf.pack(fill=tk.X, padx=10, pady=5)
        tk.Button(bf, text="🗑 清除已完成", command=self._todo_clear_done,
                 font=("微软雅黑", 9), cursor="hand2", fg=self.colors['danger'],
                 bg='white', bd=1, relief=tk.GROOVE).pack(side=tk.LEFT)
        self.todo_listbox = tk.Listbox(left, font=("微软雅黑", 10), bg='white',
                                       selectmode=tk.SINGLE, height=14)
        self.todo_listbox.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        self.todo_listbox.bind("<Double-Button-1>", lambda e: self._todo_toggle())
        self.todo_listbox.bind("<Button-3>", lambda e: self._todo_delete())
        # 右栏：备忘录
        right = tk.Frame(main, bg='white', relief=tk.GROOVE, bd=1)
        right.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        tk.Label(right, text="📝 备忘录", font=("微软雅黑", 13, "bold"),
                bg='white', fg=self.colors['dark']).pack(pady=10)
        mf = tk.Frame(right, bg='white')
        mf.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        e = scrolledtext.ScrolledText(mf, font=("微软雅黑", 10), height=16, bd=1, relief=tk.GROOVE)
        e.pack(fill=tk.BOTH, expand=True)
        if (DATA_DIR / "memo.txt").exists():
            e.insert("1.0", (DATA_DIR / "memo.txt").read_text(encoding="utf-8"))
        tk.Button(right, text="💾 保存备忘录", command=self._save_memo, cursor="hand2",
                 bg=self.colors['primary'], fg="white", font=("微软雅黑", 10, "bold"),
                 width=16).pack(pady=8)
        self.memo_text = e
        inp.bind("<Return>", lambda e: self._todo_add())
        self._todo_refresh_list()
        self._start_reminder_check()

    def _save_memo(self):
        try:
            if self.memo_text:
                (DATA_DIR / "memo.txt").write_text(self.memo_text.get("1.0", tk.END).strip(), encoding="utf-8")
                self.set_status("备忘录已保存")
        except Exception as e:
            messagebox.showerror("错误", f"保存失败: {e}")

    def _todo_add(self):
        try:
            inp = None
            for w in self.content_frame.winfo_children():
                for f in (w.winfo_children() if hasattr(w, 'winfo_children') else []):
                    for c in (f.winfo_children() if hasattr(f, 'winfo_children') else []):
                        if isinstance(c, tk.Entry) and c.get().strip():
                            inp = c; break
            if inp is None: return
            t = inp.get().strip()
            if not t: return
            self._todo_data.append({"text": t, "done": False, "remind": None})
            inp.delete(0, tk.END)
            self._todo_refresh_list()
            self._save_todo()
        except Exception:
            pass

    def _todo_toggle(self):
        sel = self.todo_listbox.curselection()
        if not sel or sel[0] >= len(self._todo_data): return
        self._todo_data[sel[0]]["done"] = not self._todo_data[sel[0]]["done"]
        self._todo_refresh_list(); self._save_todo()

    def _todo_delete(self):
        sel = self.todo_listbox.curselection()
        if not sel or sel[0] >= len(self._todo_data): return
        self._todo_data.pop(sel[0])
        self._todo_refresh_list(); self._save_todo()

    def _todo_clear_done(self):
        self._todo_data = [t for t in self._todo_data if not t["done"]]
        self._todo_refresh_list(); self._save_todo()

    def _todo_set_reminder(self):
        sel = self.todo_listbox.curselection()
        if not sel or sel[0] >= len(self._todo_data): return
        idx = sel[0]
        win = tk.Toplevel(self.root)
        win.title("设置提醒"); win.geometry("350x200")
        win.transient(self.root); win.grab_set()
        tk.Label(win, text="设置提醒时间（分钟）:", font=("微软雅黑", 11)).pack(pady=15)
        t_var = tk.StringVar(value="5")
        ttk.Combobox(win, textvariable=t_var, values=["1","5","10","15","30","60"],
                     state="readonly", width=8).pack()
        d_var = tk.StringVar(value=datetime.now().strftime("%H:%M"))
        tk.Label(win, text="或选择具体时间 (HH:MM):", font=("微软雅黑", 10)).pack(pady=8)
        tk.Entry(win, textvariable=d_var, width=10, font=("Consolas", 12)).pack()
        def _set():
            try:
                remind_dt = datetime.now()
                if ":" in d_var.get():
                    h,m = d_var.get().split(":")
                    remind_dt = remind_dt.replace(hour=int(h), minute=int(m), second=0)
                    if remind_dt < datetime.now():
                        remind_dt = remind_dt.replace(day=remind_dt.day+1)
                else:
                    remind_dt = remind_dt.replace(second=0) + timedelta(minutes=float(t_var.get()))
                idx2 = self.todo_listbox.curselection()
                if idx2 and idx2[0] < len(self._todo_data):
                    self._todo_data[idx2[0]]["remind"] = remind_dt.timestamp()
                    self._save_todo()
                    self._todo_refresh_list()
                    self.set_status(f"⏰ 提醒已设置: {remind_dt.strftime('%H:%M')}")
                win.destroy()
            except: pass
        tk.Button(win, text="✅ 确定", command=_set, bg=self.colors['primary'],
                 fg="white", font=("微软雅黑", 11, "bold"), width=12).pack(pady=12)
        dt_f = tk.Frame(win); dt_f.pack()

    def _todo_refresh_list(self):
        if not self.todo_listbox: return
        self.todo_listbox.delete(0, tk.END)
        now = time.time()
        for i, item in enumerate(self._todo_data):
            txt = item["text"][:50]
            prefix = "✅" if item["done"] else "⬜"
            remind = item.get("remind")
            rd = ""
            if remind:
                remaining = int((remind - now) / 60)
                rd = f" ⏰{remaining}分" if remaining > 0 else " ⏰已到"
            self.todo_listbox.insert(tk.END, f" {prefix} {txt}{rd}")

    def _start_reminder_check(self):
        def _check():
            while True:
                try:
                    now = time.time()
                    for item in self._todo_data:
                        remind = item.get("remind")
                        if remind and 0 < remind - now < 60 and not item.get("_reminded"):
                            item["_reminded"] = True
                            diff = int((remind - now) / 60)
                            t = item["text"][:30]
                            self.root.after(0, lambda: messagebox.showinfo("⏰ 提醒",
                                f"事项「{t}」将在 {diff} 分钟后到期！"))
                except Exception:
                    pass
                time.sleep(30)
        t = threading.Thread(target=_check, daemon=True)
        t.start()

    def _save_todo(self):
        try:
            (DATA_DIR / "todo.json").write_text(
                json.dumps(self._todo_data, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass
