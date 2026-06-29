# -*- coding: utf-8 -*-
"""万能办公助手 — QuickTools（文本/编码/时间/哈希/二维码/JSON/正则/取色）"""
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
from pathlib import Path
import os, sys, json, time, shutil, hashlib, threading
from datetime import datetime
from collections import Counter
import base64, urllib.parse, re, qrcode


class QuickToolsMixin:
    """QuickTools — 所有方法通过self访问OfficeAssistant的属性"""

    def _show_quick_tools(self):
        self.clear_content()
        self._section_header("快捷工具", "文本 · 编码 · 时间 · 哈希 · 二维码 · JSON · 正则 · 取色")
        self._show_tips("点击卡片直接打开工具 → 输入内容 → 一键出结果", "二维码、取色器等工具无需选文件，直接打开即用")
        row = tk.Frame(self.content_frame, bg=self.colors['light']); row.pack(fill=tk.X, padx=10)
        for title, desc, cb in [
            ("📝 文本处理","大小写/去重/排序/统计", self._text_tools),
            ("🔣 编码转换","Base64/URL/Hex", self._encode_tools),
            ("⏱ 时间工具","时间戳/日期差/实时", self._time_tools),
        ]: self._create_card(row, title, desc, cb)
        row2 = tk.Frame(self.content_frame, bg=self.colors['light']); row2.pack(fill=tk.X, padx=10, pady=5)
        for title, desc, cb in [
            ("🔑 哈希计算","MD5/SHA1/SHA256", self._hash_tools),
            ("🧹 系统清理","临时文件/缓存/回收站", self._clean_dlg),
            ("📂 文件对比","逐行比较两个文件", self._diff_files_dlg),
        ]: self._create_card(row2, title, desc, cb)
        row3 = tk.Frame(self.content_frame, bg=self.colors['light']); row3.pack(fill=tk.X, padx=10)
        for title, desc, cb in [
            ("📱 二维码","文本/网址一键生成", self._qrcode_dlg),
            ("📋 JSON格式化","格式化/压缩/校验", self._json_format_dlg),
            ("🔍 正则测试","实时匹配测试工具", self._regex_tester_dlg),
        ]: self._create_card(row3, title, desc, cb)
        row4 = tk.Frame(self.content_frame, bg=self.colors['light']); row4.pack(fill=tk.X, padx=10, pady=5)
        for title, desc, cb in [
            ("🎨 取色器","获取屏幕颜色值", self._color_picker_dlg),
        ]: self._create_card(row4, title, desc, cb)

    def _text_tools(self):
        win = tk.Toplevel(self.root); win.title("文本处理"); win.geometry("600x450")
        win.transient(self.root); win.grab_set();
        tk.Label(win, text="📝 文本处理", font=("微软雅黑", 14, "bold")).pack(pady=8)
        inp = scrolledtext.ScrolledText(win, height=6, font=("Consolas", 10)); inp.pack(fill=tk.X, padx=15, pady=5)
        bf = tk.Frame(win); bf.pack(pady=5)
        ops = [("大写","upper"),("小写","lower"),("首字母大写","title"),("去重行","unique"),("排序行","sort"),("统计","stats")]
        for txt, key in ops:
            tk.Button(bf, text=txt, font=("微软雅黑",9), cursor="hand2",
                     command=lambda k=key: _op(k)).pack(side=tk.LEFT, padx=3)
        out = scrolledtext.ScrolledText(win, height=8, font=("Consolas", 10)); out.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)
        l = tk.Label(win, text="", font=("微软雅黑",9), fg="gray"); l.pack()
        def _op(k):
            try:
                t = inp.get("1.0", tk.END).rstrip("\n"); lines = t.split("\n"); r = ""
                if k=="upper": r = t.upper()
                elif k=="lower": r = t.lower()
                elif k=="title": r = t.title()
                elif k=="unique": r = "\n".join(dict.fromkeys(lines))
                elif k=="sort": r = "\n".join(sorted(lines))
                elif k=="stats":
                    wc = len(t.split()); lc = len(lines); cc = len(t)
                    r = f"字符: {cc}  单词: {wc}  行: {lc}"
                out.delete("1.0", tk.END); out.insert("1.0", r)
            except Exception as e: messagebox.showerror("错误", str(e))

    def _encode_tools(self):
        win = tk.Toplevel(self.root); win.title("编码转换"); win.geometry("600x400")
        win.transient(self.root); win.grab_set();
        tk.Label(win, text="🔣 编码转换", font=("微软雅黑", 14, "bold")).pack(pady=8)
        inp = scrolledtext.ScrolledText(win, height=5, font=("Consolas", 10)); inp.pack(fill=tk.X, padx=15, pady=5)
        bf = tk.Frame(win); bf.pack(pady=5)
        ops = [("→ Base64","b64e"),("Base64 →","b64d"),("→ URL编码","urle"),("URL →","urld"),("→ Hex","hexe"),("Hex →","hexd")]
        for txt, key in ops:
            tk.Button(bf, text=txt, font=("微软雅黑",9), cursor="hand2",
                     command=lambda k=key: _op(k)).pack(side=tk.LEFT, padx=3)
        out = scrolledtext.ScrolledText(win, height=5, font=("Consolas", 10)); out.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)
        l = tk.Label(win, text="", font=("微软雅黑",9), fg="gray"); l.pack()
        def _op(k):
            try:
                t = inp.get("1.0", tk.END).strip()
                if k=="b64e": r = base64.b64encode(t.encode()).decode()
                elif k=="b64d": r = base64.b64decode(t).decode(errors='replace')
                elif k=="urle": r = urllib.parse.quote(t)
                elif k=="urld": r = urllib.parse.unquote(t)
                elif k=="hexe": r = t.encode().hex()
                elif k=="hexd": r = bytes.fromhex(t).decode(errors='replace')
                else: r = ""
                out.delete("1.0", tk.END); out.insert("1.0", r)
            except Exception as e: messagebox.showerror("错误", str(e))

    def _time_tools(self):
        win = tk.Toplevel(self.root); win.title("时间工具"); win.geometry("500x320")
        win.transient(self.root); win.grab_set();
        tk.Label(win, text="⏱ 时间工具", font=("微软雅黑", 14, "bold")).pack(pady=8)
        f1 = tk.Frame(win); f1.pack(pady=10)
        tk.Label(f1, text="时间戳 → 日期:", font=("微软雅黑",10)).pack(side=tk.LEFT)
        ts = tk.Entry(f1, width=18, font=("Consolas",11))
        ts.insert(0, str(int(time.time()))); ts.pack(side=tk.LEFT, padx=5)
        ts_l = tk.Label(f1, text=datetime.now().strftime("%Y-%m-%d %H:%M:%S"), font=("微软雅黑",10), fg="green"); ts_l.pack(side=tk.LEFT)
        def _ts():
            try: ts_l.config(text=datetime.fromtimestamp(float(ts.get())).strftime("%Y-%m-%d %H:%M:%S"))
            except: ts_l.config(text="格式错误", fg="red")
        ts.bind("<KeyRelease>", lambda e: _ts())
        f2 = tk.Frame(win); f2.pack(pady=10)
        tk.Label(f2, text="日期 → 时间戳:", font=("微软雅黑",10)).pack(side=tk.LEFT)
        d1 = tk.Entry(f2, width=20, font=("Consolas",11))
        d1.insert(0, datetime.now().strftime("%Y-%m-%d %H:%M:%S")); d1.pack(side=tk.LEFT, padx=5)
        d2 = tk.Label(f2, text="", font=("Consolas",10), fg="green"); d2.pack(side=tk.LEFT)
        def _up():
            try:
                d2.config(text=str(int(datetime.strptime(d1.get()[:19], "%Y-%m-%d %H:%M:%S").timestamp())))
            except: d2.config(text="格式错误", fg="red")
        d1.bind("<KeyRelease>", lambda e: _up())
        dl = tk.Frame(win); dl.pack(pady=10)
        tk.Label(dl, text="日期差计算:", font=("微软雅黑",10)).pack(anchor="w")
        nl = tk.Frame(dl); nl.pack()
        tk.Label(nl, text="开始:", font=("微软雅黑",9)).pack(side=tk.LEFT)
        s1 = tk.Entry(nl, width=14, font=("Consolas",10)); s1.insert(0, "2024-01-01"); s1.pack(side=tk.LEFT, padx=3)
        tk.Label(nl, text="结束:", font=("微软雅黑",9)).pack(side=tk.LEFT, padx=(10,0))
        s2 = tk.Entry(nl, width=14, font=("Consolas",10)); s2.insert(0, "2024-12-31"); s2.pack(side=tk.LEFT, padx=3)
        dd = tk.Label(nl, text="", font=("微软雅黑",10), fg="green"); dd.pack(side=tk.LEFT, padx=10)
        def _diff():
            try:
                from datetime import datetime as dt
                a = dt.strptime(s1.get()[:10], "%Y-%m-%d"); b = dt.strptime(s2.get()[:10], "%Y-%m-%d")
                d = abs((b-a).days); dd.config(text=f"相差 {d} 天")
            except: dd.config(text="格式错误", fg="red")
        tk.Button(dl, text="计算", command=_diff, cursor="hand2", font=("微软雅黑",9)).pack(pady=5)

    def _hash_tools(self):
        win = tk.Toplevel(self.root); win.title("哈希计算"); win.geometry("650x450")
        win.transient(self.root); win.grab_set();
        tk.Label(win, text="🔑 哈希计算", font=("微软雅黑", 14, "bold")).pack(pady=8)
        nb = ttk.Notebook(win); nb.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        # Tab1: 文本哈希
        f1 = tk.Frame(nb); nb.add(f1, text="文本哈希")
        te = scrolledtext.ScrolledText(f1, height=4, font=("Consolas",10)); te.pack(fill=tk.X, padx=10, pady=5)
        rl = tk.Frame(f1); rl.pack(pady=5)
        for alg in ["md5","sha1","sha256","sha512","blake2b"]:
            tk.Button(rl, text=alg.upper(), font=("微软雅黑",9), cursor="hand2",
                     command=lambda a=alg: _calc(a)).pack(side=tk.LEFT, padx=3)
        out = scrolledtext.ScrolledText(f1, height=8, font=("Consolas",10)); out.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        def _calc(a):
            try:
                txt = te.get("1.0", tk.END).strip().encode()
                h = hashlib.new(a, txt).hexdigest()
                out.delete("1.0", tk.END); out.insert("1.0", h)
            except Exception as e: messagebox.showerror("错误", str(e))
        # Tab2: 文件哈希
        f2 = tk.Frame(nb); nb.add(f2, text="文件哈希")
        fpv = tk.StringVar()
        tk.Entry(f2, textvariable=fpv, width=50, font=("微软雅黑",9)).pack(pady=10)
        tk.Button(f2, text="选择文件", command=lambda: fpv.set(filedialog.askopenfilename() or fpv.get()),
                 cursor="hand2", font=("微软雅黑",9)).pack()
        h_all = scrolledtext.ScrolledText(f2, height=8, font=("Consolas",10)); h_all.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        tk.Button(f2, text="计算所有哈希", font=("微软雅黑",10),
                 command=lambda: _fh(), bg=self.colors['primary'], fg="white", cursor="hand2").pack(pady=5)
        def _fh():
            try:
                fp = fpv.get()
                if not fp or not os.path.exists(fp): return
                import io
                h_all.delete("1.0", tk.END)
                for alg in ["md5","sha1","sha256","blake2b"]:
                    h = hashlib.new(alg)
                    with open(fp, 'rb') as f:
                        while True:
                            chunk = f.read(65536)
                            if not chunk: break
                            h.update(chunk)
                    h_all.insert(tk.END, f"{alg.upper()}: {h.hexdigest()}\n")
            except Exception as e: messagebox.showerror("错误", str(e))

    def _clean_dlg(self):
        win = tk.Toplevel(self.root); win.title("系统清理"); win.geometry("540x420")
        win.transient(self.root); win.grab_set();
        tk.Label(win, text="🧹 系统清理", font=("微软雅黑", 14, "bold")).pack(pady=8)
        tk.Label(win, text="选择要清理的项目，点击「开始清理」", font=("微软雅黑", 10), fg="gray").pack()
        vars_dict = {}
        items = [
            ("temp", "Windows 临时文件", Path(os.environ.get('TEMP', 'C:\\Windows\\Temp'))),
            ("recent", "最近文档记录", Path(os.environ['USERPROFILE']) / "Recent"),
            ("prefetch", "预缓存文件", Path("C:\\Windows\\Prefetch")),
            ("recycle", "回收站 (清空)", None),
        ]
        c = tk.Frame(win, bg='white'); c.pack(fill=tk.X, padx=20, pady=10)
        for key, desc, _ in items:
            v = tk.BooleanVar(value=True)
            tk.Checkbutton(c, text=desc, variable=v, font=("微软雅黑",10),
                          bg='white', anchor='w').pack(fill=tk.X, padx=10, pady=3)
            vars_dict[key] = v
        log = scrolledtext.ScrolledText(win, height=8, font=("Consolas", 9)); log.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)
        tk.Button(win, text="🗑 开始清理", command=lambda: _do_clean(), cursor="hand2",
                 bg=self.colors['danger'], fg="white", font=("微软雅黑", 11, "bold"), width=14).pack(pady=8)
        def _do_clean():
            try:
                total, count = 0, 0
                for key, desc, p in items:
                    if not vars_dict[key].get(): continue
                    if key == "recycle":
                        import subprocess
                        subprocess.run(["cmd.exe","/c","rd /s /q C:\\$Recycle.bin 2>nul"], capture_output=True)
                        log.insert(tk.END, f"  ✅ 回收站已清空\n")
                        count += 1
                        continue
                    if p and p.exists():
                        sz = sum(f.stat().st_size for f in p.rglob('*') if f.is_file()) / 1048576
                        import shutil
                        for f in p.rglob('*'):
                            try:
                                if f.is_file(): f.unlink()
                                elif f.is_dir(): shutil.rmtree(f, ignore_errors=True)
                            except: pass
                        total += sz; count += 1
                        log.insert(tk.END, f"  ✅ {desc}: 释放 {sz:.1f}MB\n")
                log.insert(tk.END, f"\n🎉 清理完成！释放 {total:.1f}MB，处理 {count} 项\n")
                self.set_status(f"清理完成: {total:.1f}MB")
            except Exception as e:
                log.insert(tk.END, f"❌ 错误: {e}\n")

    def _diff_files_dlg(self):
        win = tk.Toplevel(self.root); win.title("文件对比"); win.geometry("800x500")
        win.transient(self.root); win.grab_set();
        tk.Label(win, text="📂 文件对比", font=("微软雅黑", 14, "bold")).pack(pady=8)
        f1 = tk.Frame(win); f1.pack(fill=tk.X, padx=15)
        files = [tk.StringVar(), tk.StringVar()]
        for i in range(2):
            tk.Label(f1, text=f"文件{i+1}:", font=("微软雅黑",10)).grid(row=0, column=i*2, padx=5)
            tk.Entry(f1, textvariable=files[i], width=35, font=("微软雅黑",9)).grid(row=0, column=i*2+1, padx=5)
            tk.Button(f1, text="浏览", command=lambda i=i: files[i].set(filedialog.askopenfilename() or files[i].get()),
                     cursor="hand2", font=("微软雅黑",9)).grid(row=0, column=i*2+2, padx=3)
        tf = tk.Frame(win); tf.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)
        left = scrolledtext.ScrolledText(tf, height=15, font=("Consolas", 9), width=40)
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 3))
        right = scrolledtext.ScrolledText(tf, height=15, font=("Consolas", 9), width=40)
        right.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(3, 0))
        lc = tk.Label(win, text="", font=("微软雅黑",9), fg="gray"); lc.pack()
        def _go():
            try:
                a = Path(files[0].get()).read_text(encoding="utf-8", errors="replace").split("\n")
                b = Path(files[1].get()).read_text(encoding="utf-8", errors="replace").split("\n")
                left.delete("1.0", tk.END); right.delete("1.0", tk.END)
                for i, l in enumerate(a, 1): left.insert(tk.END, f"{i:4d}| {l}\n")
                for i, l in enumerate(b, 1): right.insert(tk.END, f"{i:4d}| {l}\n")
                diffs = sum(1 for i in range(max(len(a),len(b))) if i>=len(a) or i>=len(b) or a[i]!=b[i])
                lc.config(text=f"差异: {diffs} 行  |  左: {len(a)} 行  右: {len(b)} 行")
            except Exception as e: messagebox.showerror("错误", str(e))
        tk.Button(win, text="🔄 对比", command=_go, cursor="hand2",
                 bg=self.colors['primary'], fg="white", font=("微软雅黑", 10, "bold"), width=12).pack(pady=5)

    def _qrcode_dlg(self):
        win = tk.Toplevel(self.root); win.title("二维码生成"); win.geometry("480x520")
        win.transient(self.root); win.grab_set();
        tk.Label(win, text="📱 二维码生成", font=("微软雅黑", 14, "bold")).pack(pady=8)
        tv = tk.StringVar(value="https://example.com")
        e = tk.Entry(win, textvariable=tv, font=("微软雅黑", 11), width=40)
        e.pack(pady=8); e.focus()
        preview_lbl = tk.Label(win, text="", bg='white'); preview_lbl.pack(pady=5)
        f_size = tk.IntVar(value=4)
        tk.Scale(win, from_=2, to=10, orient=tk.HORIZONTAL, variable=f_size,
                label="尺寸", font=("微软雅黑",8)).pack()
        tk.Button(win, text="🔄 生成", command=lambda: _gen(), cursor="hand2",
                 bg=self.colors['primary'], fg="white", font=("微软雅黑", 11, "bold"), width=10).pack(pady=5)
        save = [None]
        def _gen():
            try:
                from PIL import Image as _PIL_Img
                import io
                txt = tv.get().strip()
                if not txt: return
                qr = qrcode.QRCode(box_size=f_size.get())
                qr.add_data(txt); qr.make(fit=True)
                img = qr.make_image(fill_color="#0F172A", back_color="white")
                pimg = img.resize((200, 200))
                buf = io.BytesIO(); pimg.save(buf, format="PNG")
                from tkinter import PhotoImage
                photo = PhotoImage(data=buf.getvalue())
                preview_lbl.config(image=photo, width=200, height=200)
                preview_lbl.image = photo
                save[0] = img
            except Exception as e: messagebox.showerror("错误", str(e))
        def _save():
            if save[0] is None: return
            fp = filedialog.asksaveasfilename(defaultextension=".png",
                filetypes=[("PNG","*.png")])
            if fp: save[0].save(fp); self.set_status(f"二维码已保存: {fp}")
        tk.Button(win, text="💾 保存", command=_save, cursor="hand2", font=("微软雅黑",9)).pack(pady=5)

    def _json_format_dlg(self):
        win = tk.Toplevel(self.root); win.title("JSON格式化"); win.geometry("600x450")
        win.transient(self.root); win.grab_set();
        tk.Label(win, text="📋 JSON格式化", font=("微软雅黑", 14, "bold")).pack(pady=8)
        btn_f = tk.Frame(win); btn_f.pack(pady=5)
        inp = scrolledtext.ScrolledText(win, height=6, font=("Consolas", 10)); inp.pack(fill=tk.X, padx=15, pady=5)
        cmd = [None]
        def _fmt(compact=False):
            try:
                t = inp.get("1.0", tk.END).strip()
                parsed = json.loads(t)
                out = json.dumps(parsed, ensure_ascii=False, indent=None if compact else 2)
                inp.delete("1.0", tk.END); inp.insert("1.0", out)
            except Exception as e: messagebox.showerror("JSON错误", str(e))
        tk.Button(btn_f, text="格式化", command=lambda: _fmt(False), cursor="hand2",
                 font=("微软雅黑",9)).pack(side=tk.LEFT, padx=3)
        tk.Button(btn_f, text="压缩", command=lambda: _fmt(True), cursor="hand2",
                 font=("微软雅黑",9)).pack(side=tk.LEFT, padx=3)
        tk.Button(btn_f, text="校验", command=lambda: _fmt(False), cursor="hand2",
                 font=("微软雅黑",9)).pack(side=tk.LEFT, padx=3)

    def _regex_tester_dlg(self):
        win = tk.Toplevel(self.root); win.title("正则测试"); win.geometry("680x500")
        win.transient(self.root); win.grab_set();
        tk.Label(win, text="🔍 正则表达式测试", font=("微软雅黑", 14, "bold")).pack(pady=8)
        f1 = tk.Frame(win); f1.pack(fill=tk.X, padx=15, pady=5)
        tk.Label(f1, text="正则:", font=("微软雅黑",10)).pack(side=tk.LEFT)
        re_var = tk.StringVar(value=r"\d+")
        tk.Entry(f1, textvariable=re_var, width=25, font=("Consolas",11)).pack(side=tk.LEFT, padx=5)
        flags_var = tk.StringVar(value="忽略大小写")
        ttk.Combobox(f1, textvariable=flags_var, values=["无","忽略大小写","多行"], state="readonly", width=10).pack(side=tk.LEFT)
        fstr = tk.Frame(win); fstr.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)
        tk.Label(fstr, text="测试文本:", font=("微软雅黑",9)).pack(anchor="w")
        test_txt = scrolledtext.ScrolledText(fstr, height=6, font=("Consolas", 10))
        test_txt.pack(fill=tk.X, pady=3)
        test_txt.insert("1.0", "Hello 123 World 456 Test 789")
        result = scrolledtext.ScrolledText(fstr, height=8, font=("Consolas", 10))
        result.pack(fill=tk.BOTH, expand=True, pady=3)
        def _test():
            try:
                pat = re_var.get(); txt = test_txt.get("1.0", tk.END).strip()
                flags = 0
                if "忽略大小写" in flags_var.get(): flags |= re.IGNORECASE
                if "多行" in flags_var.get(): flags |= re.MULTILINE
                compiled = re.compile(pat, flags)
                result.delete("1.0", tk.END)
                for m in compiled.finditer(txt):
                    start, end = m.span()
                    ctx = txt[max(0,start-10):end+10]
                    result.insert(tk.END, f"匹配: {m.group()!r} 位置: {start}-{end}\n")
                    result.insert(tk.END, f"  上下文: ...{ctx}...\n\n")
                if not compiled.findall(txt):
                    result.insert(tk.END, "❌ 无匹配\n")
                result.insert(tk.END, f"\n--- 共 {len(compiled.findall(txt))} 个匹配 ---")
            except Exception as e:
                result.delete("1.0", tk.END); result.insert(tk.END, f"❌ 正则错误: {e}")
        test_txt.bind("<KeyRelease>", lambda e: _test())
        re_var.trace_add("write", lambda *_: _test())
        _test()

    def _color_picker_dlg(self):
        win = tk.Toplevel(self.root); win.title("取色器"); win.geometry("380x330")
        win.transient(self.root); win.grab_set();
        tk.Label(win, text="🎨 取色器", font=("微软雅黑", 14, "bold")).pack(pady=8)
        tk.Label(win, text="输入颜色值预览，或点击「取色」选取屏幕颜色",
                font=("微软雅黑", 9), fg="gray").pack()
        c = tk.Frame(win, bg='white', width=120, height=60, relief=tk.GROOVE, bd=2)
        c.pack(pady=10); c.pack_propagate(False)
        e = tk.Entry(win, font=("Consolas", 14), width=16, justify="center")
        e.insert(0, "#4F46E5"); e.pack(pady=5)
        info_var = tk.StringVar(value="RGB: 79, 70, 229")
        tk.Label(win, textvariable=info_var, font=("微软雅黑", 9), fg="gray").pack()
        def _pick():
            try:
                from PIL import ImageColor
                hex_s = e.get().strip()
                if not hex_s.startswith("#"): hex_s = "#" + hex_s
                r, g, b = ImageColor.getcolor(hex_s, "RGB")
                c.config(bg=hex_s)
                info_var.set(f"RGB: {r}, {g}, {b}   HEX: {hex_s}")
            except:
                info_var.set("❌ 无效颜色值")
        e.bind("<KeyRelease>", lambda e: _pick()); _pick()
