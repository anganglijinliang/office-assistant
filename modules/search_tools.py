# -*- coding: utf-8 -*-
"""万能办公助手 — SearchTools（增强版：可双击打开、右键菜单、导出结果）"""
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
from pathlib import Path
import os, sys, json, time, shutil, hashlib, threading
from datetime import datetime
from collections import Counter


class SearchToolsMixin:
    """SearchTools — 所有方法通过self访问OfficeAssistant的属性"""

    def _show_search_tools(self):
        self.clear_content()
        self._section_header("文件内容搜索", "按关键词+类型递归搜索 · 双击打开 · 右键菜单 · 导出结果")
        self._show_tips(
            "① 选目录 → ② 输入关键词 → ③ 选文件类型 → ④ 点「搜索」",
            "📂 搜索结果显示后：双击打开文件 | 右键更多操作 | 工具栏可导出",
            "支持 .txt .py .md .csv .html .json .xml .yaml .log 等文本文件",
            "点击「停止」可中断长时间搜索"
        )
        # 搜索参数区
        sf = tk.Frame(self.content_frame, bg=self.colors['light'])
        sf.pack(fill=tk.X, padx=20, pady=5)
        tk.Label(sf, text="📁 目录:", font=("微软雅黑", 10), bg=self.colors['light']).grid(row=0, col=0, sticky="w")
        d_var = tk.StringVar(value=str(Path.home() / "Desktop"))
        tk.Entry(sf, textvariable=d_var, width=45, font=("微软雅黑", 10)).grid(row=0, col=1, padx=5)
        tk.Button(sf, text="浏览", command=lambda: (d:=filedialog.askdirectory(), d_var.set(d) if d else None),
                 cursor="hand2", font=("微软雅黑", 9)).grid(row=0, col=2)
        tk.Label(sf, text="🔍 关键词:", font=("微软雅黑", 10), bg=self.colors['light']).grid(row=1, col=0, sticky="w", pady=(8,0))
        kw_var = tk.StringVar()
        kw_entry = tk.Entry(sf, textvariable=kw_var, width=45, font=("微软雅黑", 10))
        kw_entry.grid(row=1, col=1, padx=5, pady=(8,0))
        tk.Label(sf, text="📄 类型:", font=("微软雅黑", 10), bg=self.colors['light']).grid(row=2, col=0, sticky="w", pady=(8,0))
        ft_var = tk.StringVar(value="*.txt")
        ttk.Combobox(sf, textvariable=ft_var,
            values=["*.txt","*.py","*.md","*.csv","*.log","*.html","*.json","*.yaml","*.xml","*.*"],
            state="readonly", width=16).grid(row=2, col=1, sticky="w", padx=5, pady=(8,0))

        # 操作工具栏
        toolbar = tk.Frame(self.content_frame, bg=self.colors['light'])
        toolbar.pack(fill=tk.X, padx=20, pady=(8, 2))
        search_btn = tk.Button(toolbar, text="🔍 搜索", font=("微软雅黑", 10),
                 bg=self.colors['primary'], fg="white", cursor="hand2", width=10)
        search_btn.pack(side=tk.LEFT, padx=(0, 8))
        open_btn = tk.Button(toolbar, text="📂 打开文件", font=("微软雅黑", 9),
                 cursor="hand2", state="disabled", width=10)
        open_btn.pack(side=tk.LEFT, padx=4)
        folder_btn = tk.Button(toolbar, text="📁 打开所在文件夹", font=("微软雅黑", 9),
                 cursor="hand2", state="disabled", width=14)
        folder_btn.pack(side=tk.LEFT, padx=4)
        copy_btn = tk.Button(toolbar, text="📋 复制路径", font=("微软雅黑", 9),
                 cursor="hand2", state="disabled", width=10)
        copy_btn.pack(side=tk.LEFT, padx=4)
        export_btn = tk.Button(toolbar, text="💾 导出结果", font=("微软雅黑", 9),
                 cursor="hand2", state="disabled", width=10)
        export_btn.pack(side=tk.LEFT, padx=4)
        clear_btn = tk.Button(toolbar, text="🗑 清空", font=("微软雅黑", 9),
                 cursor="hand2", width=6)
        clear_btn.pack(side=tk.RIGHT, padx=4)

        # 搜索结果区（Listbox）
        res_frame = tk.Frame(self.content_frame, bg=self.colors['light'])
        res_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)
        listbox = tk.Listbox(res_frame, font=("Consolas", 10), bg="white",
                            selectmode=tk.SINGLE, activestyle="none",
                            exportselection=False)
        scrollbar = tk.Scrollbar(res_frame, orient=tk.VERTICAL, command=listbox.yview)
        listbox.configure(yscrollcommand=scrollbar.set)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self._result_map = []

        # 事件绑定
        def _on_double_click(event):
            sel = listbox.curselection()
            if not sel or sel[0] >= len(self._result_map): return
            path, _, _ = self._result_map[sel[0]]
            if os.path.exists(path): os.startfile(path)

        def _on_select(event):
            sel = listbox.curselection()
            has_sel = len(sel) > 0 and sel[0] < len(self._result_map)
            state = "normal" if has_sel else "disabled"
            open_btn.config(state=state); folder_btn.config(state=state); copy_btn.config(state=state)

        def _open_selected():
            sel = listbox.curselection()
            if not sel or sel[0] >= len(self._result_map): return
            path, _, _ = self._result_map[sel[0]]
            if os.path.exists(path): os.startfile(path)

        def _open_folder():
            sel = listbox.curselection()
            if not sel or sel[0] >= len(self._result_map): return
            path, _, _ = self._result_map[sel[0]]
            folder = os.path.dirname(path)
            if os.path.exists(folder): os.startfile(folder)

        def _copy_path():
            sel = listbox.curselection()
            if not sel or sel[0] >= len(self._result_map): return
            path, _, _ = self._result_map[sel[0]]
            self.root.clipboard_clear(); self.root.clipboard_append(path)
            self.set_status("路径已复制到剪贴板")

        def _export_results():
            if not self._result_map: return messagebox.showwarning("提示", "没有搜索结果可导出")
            save_path = filedialog.asksaveasfilename(title="导出搜索结果",
                defaultextension=".csv", filetypes=[("CSV文件","*.csv"),("文本文件","*.txt")])
            if not save_path: return
            try:
                if save_path.endswith('.csv'):
                    import csv
                    with open(save_path, 'w', newline='', encoding='utf-8-sig') as f:
                        w = csv.writer(f)
                        w.writerow(["文件路径","匹配行","行号"])
                        for path, line, lineno in self._result_map:
                            w.writerow([path, line, lineno])
                else:
                    with open(save_path, 'w', encoding='utf-8') as f:
                        for path, line, lineno in self._result_map:
                            f.write(f"{path}\n  L{lineno}: {line}\n\n")
                self.set_status(f"导出成功: {len(self._result_map)} 条")
                messagebox.showinfo("完成", f"导出成功！\n{len(self._result_map)} 条记录\n{save_path}")
            except Exception as e: messagebox.showerror("错误", f"导出失败:\n{e}")

        def _clear_results():
            listbox.delete(0, tk.END); self._result_map.clear()
            open_btn.config(state="disabled"); folder_btn.config(state="disabled")
            copy_btn.config(state="disabled"); export_btn.config(state="disabled")
            self.set_status("就绪")

        # 右键菜单
        right_click_menu = tk.Menu(self.root, tearoff=0, font=("微软雅黑", 9))
        right_click_menu.add_command(label="📂 打开文件", command=_open_selected)
        right_click_menu.add_command(label="📁 打开所在文件夹", command=_open_folder)
        right_click_menu.add_command(label="📋 复制路径", command=_copy_path)
        right_click_menu.add_separator()
        right_click_menu.add_command(label="💾 导出所有结果", command=_export_results)

        def _on_right_click(event):
            idx = listbox.nearest(event.y)
            listbox.selection_clear(0, tk.END); listbox.selection_set(idx); listbox.activate(idx)
            _on_select(None); right_click_menu.tk_popup(event.x_root, event.y_root)

        listbox.bind("<Double-Button-1>", _on_double_click)
        listbox.bind("<<ListboxSelect>>", _on_select)
        listbox.bind("<Button-3>", _on_right_click)
        open_btn.config(command=_open_selected); folder_btn.config(command=_open_folder)
        copy_btn.config(command=_copy_path); export_btn.config(command=_export_results)
        clear_btn.config(command=_clear_results)

        def _on_key(event):
            if event.keysym in ("Return","KP_Enter"): _open_selected()
            elif event.keysym == "Delete": _clear_results()
        listbox.bind("<Return>", _on_key); listbox.bind("<KP_Enter>", _on_key)
        listbox.bind("<Delete>", lambda e: _clear_results())

        # 搜索功能
        self._search_cancel = False

        def _search():
            d = d_var.get(); kw = kw_var.get().strip(); ft = ft_var.get()
            if not d or not kw: return messagebox.showwarning("提示", "请填写目录和关键词")
            if not os.path.isdir(d): return messagebox.showerror("错误", f"目录不存在:\n{d}")
            listbox.delete(0, tk.END); self._result_map.clear()
            export_btn.config(state="disabled"); self.set_status("搜索中…")
            self._search_cancel = False

            def _cancel():
                self._search_cancel = True
                search_btn.config(text="🔍 搜索", command=_search, state="normal")
            search_btn.config(text="⏹ 停止", command=_cancel)

            def _do_search():
                cnt, matches = 0, 0
                max_results = 100
                try:
                    for p in Path(d).rglob(ft):
                        if self._search_cancel: break
                        if not p.is_file() or matches >= max_results: continue
                        if p.stat().st_size > 5 * 1024 * 1024: continue
                        cnt += 1
                        try:
                            content = p.read_text(encoding="utf-8", errors="ignore")
                            if kw.lower() in content.lower():
                                for i, line in enumerate(content.split("\n"), 1):
                                    if kw.lower() in line.lower() and matches < max_results:
                                        matches += 1
                                        display = line.strip()[:120]
                                        self._result_map.append((str(p), display, i))
                                        self.root.after(0, lambda p=str(p), d=display: (
                                            listbox.insert(tk.END, f"  {Path(p).parent.name}/{Path(p).name}"),
                                            listbox.itemconfig(tk.END, fg="#1E293B")
                                        ))
                                        break
                        except Exception: pass
                    status = "已取消" if self._search_cancel else "就绪"
                    if self._result_map:
                        self.root.after(0, lambda: (
                            listbox.insert(tk.END, ""),
                            listbox.insert(tk.END, f"  📊 扫描 {cnt} 个文件，找到 {matches} 个匹配  [{status}]"),
                            listbox.itemconfig(tk.END, fg="#059669", font=("微软雅黑", 9)),
                            export_btn.config(state="normal"),
                            self.set_status(f"搜索完成: {matches} 个匹配 / {cnt} 个文件")
                        ))
                    else:
                        self.root.after(0, lambda: (
                            listbox.insert(tk.END, f"  📭 未找到包含「{kw}」的文件（扫描 {cnt} 个）  [{status}]"),
                            listbox.itemconfig(tk.END, fg="#D97706"),
                            self.set_status("搜索完成: 未找到匹配")
                        ))
                except Exception as e:
                    self.root.after(0, lambda e=e: (
                        listbox.insert(tk.END, f"  ❌ 搜索出错: {e}"),
                        listbox.itemconfig(tk.END, fg="red")
                    ))
                self.root.after(0, lambda: search_btn.config(text="🔍 搜索", command=_search, state="normal"))
                self._search_cancel = False

            threading.Thread(target=_do_search, daemon=True).start()

        search_btn.config(command=_search)
        kw_entry.bind("<Return>", lambda e: _search())
