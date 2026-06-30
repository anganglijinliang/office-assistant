# -*- coding: utf-8 -*-
"""万能办公助手 — FileTools（批量重命名 · 文件分类 · 查重 · 批量打印 · 分割 · 打包 · 快捷方式 · 校验）"""
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
from pathlib import Path
import os, sys, json, time, shutil, hashlib, threading
from datetime import datetime
from collections import Counter


class FileToolsMixin:
    """FileTools — 所有方法通过self访问OfficeAssistant的属性"""

    def _show_file_tools(self):
        self.clear_content()
        self._section_header("文件处理工具", "批量操作 · 智能分类 · 查重去重 · 批量打印")
        self._show_tips(
            "点击下方卡片选择功能 → 选文件 → 设置参数 → 一键完成",
            "✂ 分割文件前，请先备份重要数据"
        )
        row = tk.Frame(self.content_frame, bg=self.colors['light'])
        row.pack(fill=tk.X, padx=10)
        self._create_card(row, " 📝 批量重命名", "前缀/后缀/替换/序号", self._batch_rename_dialog, "开始")
        self._create_card(row, " 📂 文件分类", "递归子目录整理", self._classify_files_dlg, "开始")
        self._create_card(row, " 🔎 查重并清理", "MD5查重+智能删除", self._find_dupes_dlg, "开始")
        row2 = tk.Frame(self.content_frame, bg=self.colors['light'])
        row2.pack(fill=tk.X, padx=10, pady=5)
        self._create_card(row2, " 🖨 批量打印", "批量打印Word/PDF/图片", self._batch_print_dlg, "开始")
        self._create_card(row2, " ✂ 分割文件", "按大小/行数分割", self._split_file_dlg, "开始")
        self._create_card(row2, " 📦 打包压缩", "文件夹 → ZIP/7z", self._archive_dlg, "开始")
        row3 = tk.Frame(self.content_frame, bg=self.colors['light'])
        row3.pack(fill=tk.X, padx=10)
        self._create_card(row3, " 🔗 创建快捷方式", "文件/文件夹快速访问", self._shortcut_dlg, "开始")
        self._create_card(row3, " 🔐 文件校验", "MD5/SHA1/SHA256", self._checksum_dlg, "开始")

    def _batch_rename_dialog(self):
        win = tk.Toplevel(self.root); win.title("批量重命名"); win.geometry("750x580")
        win.transient(self.root); win.grab_set(); win.configure(bg=self.colors['light'])
        tk.Label(win, text="📝 批量重命名", font=("微软雅黑", 14, "bold"), bg=self.colors['light']).pack(pady=8)
        tk.Label(win, text="①添加文件 ②设置规则 ③预览 ④执行", font=("微软雅黑", 9), fg="gray", bg=self.colors['light']).pack()
        files = []
        # 文件选择区
        sel_f = tk.Frame(win, bg=self.colors['light']); sel_f.pack(fill=tk.X, padx=15, pady=5)
        tk.Button(sel_f, text="📂 添加文件", command=lambda: _pick(), cursor="hand2",
                  bg=self.colors['primary'], fg="white", font=("微软雅黑", 9)).pack(side=tk.LEFT, padx=2)
        tk.Button(sel_f, text="📂 添加文件夹", command=lambda: _pick_dir(), cursor="hand2",
                  font=("微软雅黑", 9)).pack(side=tk.LEFT, padx=2)
        tk.Button(sel_f, text="🗑 清空", command=lambda: (files.clear(), _ref()), cursor="hand2",
                  font=("微软雅黑", 9)).pack(side=tk.LEFT, padx=2)
        file_lb = tk.Listbox(win, height=4, font=("Consolas", 9)); file_lb.pack(fill=tk.X, padx=15, pady=3)
        def _pick():
            f = filedialog.askopenfilenames(title="选择文件")
            for p in f:
                if p not in files: files.append(p)
            _ref()
        def _pick_dir():
            d = filedialog.askdirectory(title="选择文件夹")
            if d:
                for p in Path(d).rglob("*"):
                    if p.is_file() and str(p) not in files: files.append(str(p))
            _ref()
        def _ref():
            file_lb.delete(0, tk.END)
            for f in files[-50:]: file_lb.insert(tk.END, f"  {Path(f).name}")
        # 规则设置区
        rule_f = tk.Frame(win, bg=self.colors['light']); rule_f.pack(fill=tk.X, padx=15, pady=5)
        mode_var = tk.StringVar(value="prefix")
        tk.Radiobutton(rule_f, text="加前缀", variable=mode_var, value="prefix", bg=self.colors['light']).grid(row=0, column=0)
        tk.Radiobutton(rule_f, text="加后缀", variable=mode_var, value="suffix", bg=self.colors['light']).grid(row=0, column=1)
        tk.Radiobutton(rule_f, text="替换", variable=mode_var, value="replace", bg=self.colors['light']).grid(row=0, column=2)
        tk.Radiobutton(rule_f, text="序号", variable=mode_var, value="number", bg=self.colors['light']).grid(row=0, column=3)
        tk.Label(rule_f, text="文本:", bg=self.colors['light']).grid(row=1, column=0, pady=5, sticky="e")
        txt_var = tk.StringVar(value="new_")
        tk.Entry(rule_f, textvariable=txt_var, width=18, font=("微软雅黑", 9)).grid(row=1, column=1, padx=3)
        tk.Label(rule_f, text="旧文本(替换时):", bg=self.colors['light']).grid(row=1, column=2, sticky="e")
        old_var = tk.StringVar(value="old")
        tk.Entry(rule_f, textvariable=old_var, width=14, font=("微软雅黑", 9)).grid(row=1, column=3, padx=3)
        log = scrolledtext.ScrolledText(win, height=10, font=("Consolas", 9), bg="white")
        log.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)
        def _preview():
            log.delete("1.0", tk.END)
            if not files: return log.insert(tk.END, "请先添加文件\n")
            mode = mode_var.get(); text = txt_var.get(); old = old_var.get()
            for i, fp in enumerate(files):
                p = Path(fp); name = p.stem; ext = p.suffix
                new_name = name
                if mode == "prefix": new_name = text + name
                elif mode == "suffix": new_name = name + text
                elif mode == "replace": new_name = name.replace(old, text)
                elif mode == "number": new_name = f"{text}{i+1:03d}"
                log.insert(tk.END, f"  {name}{ext} → {new_name}{ext}\n")
        def _execute():
            if not files: return messagebox.showwarning("提示","请先添加文件")
            mode = mode_var.get(); text = txt_var.get(); old = old_var.get()
            ok, fail = 0, 0
            for i, fp in enumerate(files):
                try:
                    p = Path(fp); name = p.stem; ext = p.suffix
                    new_name = name
                    if mode == "prefix": new_name = text + name
                    elif mode == "suffix": new_name = name + text
                    elif mode == "replace": new_name = name.replace(old, text)
                    elif mode == "number": new_name = f"{text}{i+1:03d}"
                    p.rename(p.with_stem(new_name))
                    ok += 1
                except Exception as e: fail += 1; log.insert(tk.END, f"  ❌ {Path(fp).name}: {e}\n")
            log.insert(tk.END, f"\n🎉 成功 {ok} 个, 失败 {fail} 个\n")
            # 更新文件列表以反映新名称
            files.clear()
            _ref()
            self.set_status(f"重命名完成: {ok} 成功")
        btn_f = tk.Frame(win, bg=self.colors['light']); btn_f.pack(pady=5)
        tk.Button(btn_f, text="👁 预览", command=_preview, cursor="hand2",
                  font=("微软雅黑", 10)).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_f, text="🚀 执行重命名", command=_execute, cursor="hand2",
                  bg=self.colors['primary'], fg="white", font=("微软雅黑", 11, "bold"), width=14).pack(side=tk.LEFT, padx=5)

    def _classify_files_dlg(self):
        win = tk.Toplevel(self.root); win.title("文件分类"); win.geometry("600x480")
        win.transient(self.root); win.grab_set(); win.configure(bg=self.colors['light'])
        tk.Label(win, text="📂 文件智能分类", font=("微软雅黑", 14, "bold"), bg=self.colors['light']).pack(pady=8)
        tk.Label(win, text="按文件类型自动归类到子目录", font=("微软雅黑", 9), fg="gray", bg=self.colors['light']).pack()
        opt_f = tk.Frame(win, bg=self.colors['light']); opt_f.pack(pady=10)
        src = tk.StringVar(); tgt = tk.StringVar()
        tk.Label(opt_f, text="源目录:", bg=self.colors['light']).grid(row=0, column=0)
        tk.Entry(opt_f, textvariable=src, width=40).grid(row=0, column=1, padx=5)
        tk.Button(opt_f, text="浏览", command=lambda: src.set(filedialog.askdirectory() or src.get())).grid(row=0, column=2)
        tk.Label(opt_f, text="目标目录:", bg=self.colors['light']).grid(row=1, column=0, pady=8)
        tk.Entry(opt_f, textvariable=tgt, width=40).grid(row=1, column=1, padx=5)
        tk.Button(opt_f, text="浏览", command=lambda: tgt.set(filedialog.askdirectory() or tgt.get())).grid(row=1, column=2)
        recursive = tk.BooleanVar(value=True)
        tk.Checkbutton(opt_f, text="递归子目录", variable=recursive, bg=self.colors['light']).grid(row=2, columnspan=2, pady=5)
        copy_mode = tk.BooleanVar(value=False)
        tk.Checkbutton(opt_f, text="复制模式（保留原文件）", variable=copy_mode, bg=self.colors['light']).grid(row=3, columnspan=2)
        btn_f = tk.Frame(win, bg=self.colors['light']); btn_f.pack(pady=5)
        log = scrolledtext.ScrolledText(win, height=12, font=("Consolas", 9)); log.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)
        cats = {}; stats = {"moved": 0, "skipped": 0}
        def _go():
            try:
                s = Path(src.get()); d = Path(tgt.get()) if tgt.get() else s
                if not s.exists(): return messagebox.showerror("错误", "源目录不存在")
                d.mkdir(exist_ok=True)
                cat_map = {
                    ".jpg": "图片", ".jpeg": "图片", ".png": "图片", ".gif": "图片", ".bmp": "图片", ".webp": "图片",
                    ".doc": "文档", ".docx": "文档", ".pdf": "文档", ".txt": "文档", ".md": "文档",
                    ".xls": "表格", ".xlsx": "表格", ".csv": "表格",
                    ".mp3": "音频", ".wav": "音频", ".flac": "音频",
                    ".mp4": "视频", ".avi": "视频", ".mkv": "视频",
                    ".zip": "压缩包", ".rar": "压缩包", ".7z": "压缩包",
                    ".exe": "程序", ".msi": "程序",
                    ".py": "代码", ".js": "代码", ".html": "代码", ".css": "代码", ".cpp": "代码",
                }
                pattern = "**/*" if recursive.get() else "*"
                for f in s.glob(pattern):
                    if f.is_file():
                        ext = f.suffix.lower()
                        cat = cat_map.get(ext, "其他")
                        dest = d / cat
                        dest.mkdir(exist_ok=True)
                        target = dest / f.name
                        if target.exists():
                            stats["skipped"] += 1
                            log.insert(tk.END, f"  ⏭ 跳过(已存在): {f.name}\n")
                        else:
                            if copy_mode.get(): shutil.copy2(str(f), str(target))
                            else: shutil.move(str(f), str(target))
                            stats["moved"] += 1
                            log.insert(tk.END, f"  ✅ {f.name} → {cat}/\n")
                    log.see(tk.END); win.update()
                log.insert(tk.END, f"\n🎉 完成！移动 {stats['moved']} 个，跳过 {stats['skipped']} 个\n")
            except Exception as e: log.insert(tk.END, f"❌ 错误: {e}\n")
        tk.Button(btn_f, text="🚀 开始分类", command=_go, cursor="hand2",
                 bg=self.colors['primary'], fg="white", font=("微软雅黑", 11, "bold"), width=14).pack()

    def _find_dupes_dlg(self):
        win = tk.Toplevel(self.root); win.title("查重并清理"); win.geometry("650x500")
        win.transient(self.root); win.grab_set(); win.configure(bg=self.colors['light'])
        tk.Label(win, text="🔎 查找重复文件", font=("微软雅黑", 14, "bold"), bg=self.colors['light']).pack(pady=8)
        tk.Label(win, text="按MD5哈希查找并删除重复文件", font=("微软雅黑", 9), fg="gray", bg=self.colors['light']).pack()
        f1 = tk.Frame(win, bg=self.colors['light']); f1.pack(pady=10)
        tk.Label(f1, text="扫描目录:", bg=self.colors['light']).pack(side=tk.LEFT)
        path = tk.StringVar(value=str(Path.home() / "Desktop"))
        tk.Entry(f1, textvariable=path, width=45).pack(side=tk.LEFT, padx=5)
        tk.Button(f1, text="浏览", command=lambda: path.set(filedialog.askdirectory() or path.get())).pack(side=tk.LEFT)
        log = scrolledtext.ScrolledText(win, height=14, font=("Consolas", 9)); log.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)
        v = tk.BooleanVar(value=True)
        tk.Checkbutton(win, text="自动保留第一个文件（删除其余）", variable=v, bg=self.colors['light']).pack()
        dup_map = {}
        def _go():
            try:
                p = Path(path.get()); dup_map.clear(); count = 0
                log.insert(tk.END, f"扫描 {p} ...\n")
                for f in p.rglob("*"):
                    if f.is_file() and f.stat().st_size > 0:
                        h = hashlib.md5()
                        with open(f, 'rb') as fh:
                            while True:
                                chunk = fh.read(65536)
                                if not chunk: break
                                h.update(chunk)
                        hex_digest = h.hexdigest()
                        dup_map.setdefault(hex_digest, []).append(f)
                        count += 1
                        if count % 100 == 0:
                            log.insert(tk.END, f"  已扫描 {count} 个文件...\n"); log.see(tk.END); win.update()
                log.insert(tk.END, f"\n扫描完成: {count} 个文件\n")
                total_dup = 0
                for h, paths in dup_map.items():
                    if len(paths) > 1:
                        total_dup += len(paths) - 1
                        log.insert(tk.END, f"\n重复组 (MD5: {h[:8]}...):\n")
                        for i, fp in enumerate(paths):
                            log.insert(tk.END, f"  {'[保留]' if i==0 else '[删除]'} {fp}\n")
                        if v.get():
                            for fp in paths[1:]:
                                try: fp.unlink(); log.insert(tk.END, f"  🗑 已删除: {fp.name}\n")
                                except Exception as e: log.insert(tk.END, f"  ❌ 删除失败: {e}\n")
                log.insert(tk.END, f"\n🎉 共 {total_dup} 个重复文件\n")
                self.set_status(f"查重完成: {total_dup} 个重复")
            except Exception as e: log.insert(tk.END, f"❌ 错误: {e}\n")
        tk.Button(win, text="🔍 开始扫描", command=_go, cursor="hand2",
                 bg=self.colors['primary'], fg="white", font=("微软雅黑", 11, "bold"), width=14).pack(pady=8)

    def _batch_print_dlg(self):
        win = tk.Toplevel(self.root); win.title("批量打印"); win.geometry("600x450")
        win.transient(self.root); win.grab_set(); win.configure(bg=self.colors['light'])
        tk.Label(win, text="🖨 批量打印", font=("微软雅黑", 14, "bold"), bg=self.colors['light']).pack(pady=8)
        opt_f = tk.Frame(win, bg=self.colors['light']); opt_f.pack(fill=tk.X, padx=20, pady=10)
        list_f = tk.Frame(opt_f, bg=self.colors['light']); list_f.pack(fill=tk.X)
        files = []
        tk.Button(list_f, text="📂 添加文件", command=lambda: _add(), cursor="hand2",
                 bg=self.colors['primary'], fg="white", font=("微软雅黑",9)).pack(side=tk.LEFT, padx=3)
        tk.Button(list_f, text="🗑 清空", command=lambda: (files.clear(), _ref()), cursor="hand2",
                 font=("微软雅黑",9)).pack(side=tk.LEFT, padx=3)
        lb = tk.Listbox(win, height=6, font=("Consolas", 10)); lb.pack(fill=tk.X, padx=20, pady=5)
        def _add():
            f = filedialog.askopenfilenames(title="选择文件", filetypes=[("文档","*.docx *.pdf *.txt *.jpg *.png")])
            for p in f:
                if p not in files: files.append(p)
            _ref()
        def _ref():
            lb.delete(0, tk.END)
            for f in files: lb.insert(tk.END, f"  {Path(f).name}  ({Path(f).parent.name})")
        log = scrolledtext.ScrolledText(win, height=8, font=("Consolas", 9)); log.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)
        def _print():
            if not files: return messagebox.showwarning("提示", "请先添加文件")
            ok = 0; fail = 0
            for fp in files:
                try:
                    os.startfile(fp, "print")
                    ok += 1; log.insert(tk.END, f"  ✅ 已发送: {Path(fp).name}\n")
                except Exception as e:
                    fail += 1; log.insert(tk.END, f"  ❌ 失败: {Path(fp).name}: {e}\n")
            log.insert(tk.END, f"\n完成: {ok} 成功, {fail} 失败\n")
        tk.Button(win, text="🖨 开始打印", command=_print, cursor="hand2",
                 bg=self.colors['primary'], fg="white", font=("微软雅黑", 11, "bold"), width=14).pack(pady=8)

    def _split_file_dlg(self):
        win = tk.Toplevel(self.root); win.title("分割文件"); win.geometry("550x420")
        win.transient(self.root); win.grab_set(); win.configure(bg=self.colors['light'])
        tk.Label(win, text="✂ 文件分割", font=("微软雅黑", 14, "bold"), bg=self.colors['light']).pack(pady=8)
        f1 = tk.Frame(win, bg=self.colors['light']); f1.pack(pady=8)
        tk.Label(f1, text="文件:", bg=self.colors['light']).pack(side=tk.LEFT)
        fp = tk.StringVar(); tk.Entry(f1, textvariable=fp, width=40).pack(side=tk.LEFT, padx=5)
        tk.Button(f1, text="选择", command=lambda: fp.set(filedialog.askopenfilename() or fp.get())).pack(side=tk.LEFT)
        f2 = tk.Frame(win, bg=self.colors['light']); f2.pack(pady=5)
        mode = tk.StringVar(value="lines")
        tk.Radiobutton(f2, text="按行数", variable=mode, value="lines", bg=self.colors['light']).pack(side=tk.LEFT, padx=10)
        tk.Radiobutton(f2, text="按大小", variable=mode, value="size", bg=self.colors['light']).pack(side=tk.LEFT, padx=10)
        per = tk.IntVar(value=1000)
        sv = tk.Spinbox(f2, from_=1, to=1000000, textvariable=per, width=8); sv.pack(side=tk.LEFT, padx=5)
        out_dir = tk.StringVar()
        tk.Frame(win, bg=self.colors['light']).pack()
        tk.Button(win, text="输出目录", command=lambda: out_dir.set(filedialog.askdirectory() or out_dir.get()),
                 cursor="hand2", font=("微软雅黑",9)).pack()
        log = scrolledtext.ScrolledText(win, height=10, font=("Consolas", 9)); log.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)
        def _go():
            try:
                p = Path(fp.get()); out = Path(out_dir.get()) if out_dir.get() else p.parent
                out.mkdir(exist_ok=True); stem = p.stem; ext = p.suffix
                data = p.read_bytes(); size_b = len(data); chunk = int(per.get())
                if mode.get() == "lines":
                    lines = p.read_text(encoding="utf-8", errors="replace").split("\n")
                    for i in range(0, len(lines), chunk):
                        part = out / f"{stem}_part{i//chunk+1}{ext}"
                        part.write_text("\n".join(lines[i:i+chunk]), encoding="utf-8")
                        log.insert(tk.END, f"  ✅ {part.name}\n")
                else:
                    size_kb = chunk * 1024
                    for i in range(0, size_b, size_kb):
                        part = out / f"{stem}_part{i//size_kb+1}{ext}"
                        part.write_bytes(data[i:i+size_kb])
                        log.insert(tk.END, f"  ✅ {part.name}\n")
                log.insert(tk.END, "\n🎉 分割完成！\n")
            except Exception as e: log.insert(tk.END, f"❌ 错误: {e}\n")
        tk.Button(win, text="✂ 开始分割", command=_go, cursor="hand2",
                 bg=self.colors['primary'], fg="white", font=("微软雅黑", 11, "bold"), width=14).pack(pady=8)

    def _archive_dlg(self):
        src = filedialog.askdirectory(title="选择文件夹打包")
        if not src: return
        save = filedialog.asksaveasfilename(title="保存压缩包",
            defaultextension=".zip", filetypes=[("ZIP","*.zip")])
        if not save: return
        def _go():
            import zipfile
            log_text = []
            try:
                with zipfile.ZipFile(save, 'w', zipfile.ZIP_DEFLATED) as zf:
                    for f in Path(src).rglob("*"):
                        if f.is_file():
                            arcname = Path(src).name + "/" + str(f.relative_to(Path(src)))
                            zf.write(str(f), arcname)
                            log_text.append(f"  + {arcname}")
                self.root.after(0, lambda: self._show_success_dialog("完成", f"打包完成\n{save}\n共 {len(log_text)} 个文件"))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("错误", str(e)))
        self._run_thread(_go)

    def _shortcut_dlg(self):
        win = tk.Toplevel(self.root); win.title("创建快捷方式"); win.geometry("550x300")
        win.transient(self.root); win.grab_set(); win.configure(bg=self.colors['light'])
        tk.Label(win, text="🔗 创建快捷方式", font=("微软雅黑", 14, "bold"), bg=self.colors['light']).pack(pady=8)
        f1 = tk.Frame(win, bg=self.colors['light']); f1.pack(pady=10)
        tk.Label(f1, text="目标:", bg=self.colors['light']).grid(row=0, column=0)
        fp = tk.StringVar(); tk.Entry(f1, textvariable=fp, width=45).grid(row=0, column=1, padx=5)
        tk.Button(f1, text="浏览", command=lambda: fp.set(filedialog.askopenfilename() or fp.get())).grid(row=0, column=2)
        tk.Label(f1, text="名称:", bg=self.colors['light']).grid(row=1, column=0, pady=8)
        nv = tk.StringVar(); tk.Entry(f1, textvariable=nv, width=30).grid(row=1, column=1, padx=5, sticky="w")
        lv = tk.Label(f1, text="", font=("微软雅黑",9), fg="green", bg=self.colors['light'])
        lv.grid(row=2, columnspan=3, pady=5)
        def _go():
            try:
                import win32com.client
                shell = win32com.client.Dispatch("WScript.Shell")
                name = nv.get().strip() or Path(fp.get()).stem
                dest_dir = Path(os.environ['USERPROFILE']) / "Desktop"
                sc = dest_dir / f"{name}.lnk"
                shortcut = shell.CreateShortcut(str(sc))
                shortcut.TargetPath = fp.get()
                shortcut.WorkingDirectory = str(Path(fp.get()).parent)
                shortcut.Save()
                lv.config(text=f"✅ 已创建: {sc}")
            except Exception as e: messagebox.showerror("错误", str(e))
        tk.Button(win, text="🔗 创建到桌面", command=_go, cursor="hand2",
                 bg=self.colors['primary'], fg="white", font=("微软雅黑", 11, "bold"), width=16).pack(pady=10)

    def _checksum_dlg(self):
        win = tk.Toplevel(self.root); win.title("文件校验"); win.geometry("600x400")
        win.transient(self.root); win.grab_set(); win.configure(bg=self.colors['light'])
        tk.Label(win, text="🔐 文件校验", font=("微软雅黑", 14, "bold"), bg=self.colors['light']).pack(pady=8)
        f1 = tk.Frame(win, bg=self.colors['light']); f1.pack(pady=10)
        files = []
        tk.Button(f1, text="📂 选择文件", command=lambda: _add(), cursor="hand2").pack(side=tk.LEFT, padx=5)
        lb = tk.Listbox(win, height=4, font=("Consolas", 10)); lb.pack(fill=tk.X, padx=20, pady=5)
        def _add():
            f = filedialog.askopenfilename()
            if f and f not in files: files.append(f); lb.insert(tk.END, f"  {Path(f).name} ({Path(f).parent.name})")
        d = scrolledtext.ScrolledText(win, height=10, font=("Consolas", 10)); d.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)
        def _go():
            try:
                d.delete("1.0", tk.END)
                for p in files:
                    pobj = Path(p)
                    d.insert(tk.END, f"📄 {pobj.name}\n")
                    for alg in ["md5","sha1","sha256"]:
                        h = hashlib.new(alg)
                        with open(str(pobj), 'rb') as fh:
                            while True:
                                chunk = fh.read(65536)
                                if not chunk: break
                                h.update(chunk)
                        d.insert(tk.END, f"  {alg.upper()}: {h.hexdigest()}\n")
                    d.insert(tk.END, "\n")
            except Exception as e: messagebox.showerror("错误", str(e))
        tk.Button(win, text="🔐 计算校验和", command=_go, cursor="hand2",
                 bg=self.colors['primary'], fg="white", font=("微软雅黑", 11, "bold"), width=14).pack(pady=5)
