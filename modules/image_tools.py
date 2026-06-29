# -*- coding: utf-8 -*-
"""万能办公助手 — ImageTools（格式转换·缩放·水印·拼接·压缩·OCR·九宫格）"""
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
from pathlib import Path
import os, sys, json, time, shutil, hashlib, threading
from datetime import datetime
from collections import Counter

from utils import PIL_AVAILABLE, get_font, check_ocr_available, get_ocr_install_guide

from PIL import Image, ImageDraw


class ImageToolsMixin:
    """ImageTools — 所有方法通过self访问OfficeAssistant的属性"""

    def _show_image_tools(self):
        self.clear_content()
        self._section_header("图片工具", "格式转换 · 缩放 · 水印 · 拼接 · 压缩 · OCR · 九宫格")
        self._show_tips(
            "选功能 → 选择图片 → 设置参数 → 保存结果",
            "批量处理时，同名文件会自动编号，不会覆盖原文件"
        )
        if not PIL_AVAILABLE:
            tk.Label(self.content_frame, text="⚠️ Pillow未安装", fg="red",
                    font=("微软雅黑", 12), bg=self.colors['light']).pack(pady=60)
            return
        for rdata in [
            [("🔄 格式转换", "PNG/JPG/WebP互转", self._convert_img_dlg),
             ("📏 缩放", "按比例/尺寸调整", self._resize_img_dlg),
             ("💧 水印", "添加文字水印", self._watermark_dlg)],
            [("🧩 拼接", "多图横向/纵向拼接", self._concat_img_dlg),
             ("📦 压缩", "减小文件体积", self._compress_img_dlg),
             ("📝 OCR", "图片→文字识别", self._ocr_img_dlg)],
            [("🔲 九宫格", "切割为9张", self._nine_grid_dlg)],
        ]:
            row = tk.Frame(self.content_frame, bg=self.colors['light']); row.pack(fill=tk.X, padx=10)
            for title, desc, cb in rdata:
                self._create_card(row, title, desc, cb)

    def _convert_img_dlg(self):
        files = filedialog.askopenfilenames(title="选择图片", filetypes=[("图片","*.png *.jpg *.jpeg *.bmp *.webp")])
        if not files: return
        win = tk.Toplevel(self.root); win.title("格式转换"); win.geometry("350x200")
        win.transient(self.root); win.grab_set()
        tk.Label(win, text="目标格式:", font=("微软雅黑",11)).pack(pady=15)
        fmt_var = tk.StringVar(value="PNG")
        ttk.Combobox(win, textvariable=fmt_var, values=["PNG","JPEG","BMP","WEBP"], state="readonly", width=10).pack()
        def _go():
            try:
                fmt = fmt_var.get().lower(); ok = 0; fail = 0
                for fp in files:
                    try:
                        img = Image.open(fp)
                        out = Path(fp).with_suffix(f".{fmt}")
                        if fmt == "jpeg" and img.mode in ("RGBA","P","LA"):
                            bg = Image.new("RGB", img.size, (255,255,255))
                            bg.paste(img, mask=img.split()[-1] if img.mode=="RGBA" else None)
                            img = bg
                        img.save(str(out)); ok += 1
                    except: fail += 1
                messagebox.showinfo("完成", f"转换完成\n成功: {ok}  失败: {fail}")
            except Exception as e: messagebox.showerror("错误", str(e))
        tk.Button(win, text="开始转换", command=_go, bg=self.colors['primary'], fg="white",
                 font=("微软雅黑",11,"bold"), cursor="hand2", width=12).pack(pady=20)

    def _resize_img_dlg(self):
        files = filedialog.askopenfilenames(title="选择图片", filetypes=[("图片","*.png *.jpg")])
        if not files: return
        win = tk.Toplevel(self.root); win.title("缩放"); win.geometry("400x250")
        win.transient(self.root); win.grab_set()
        tk.Label(win, text="宽度(px):", font=("微软雅黑",10)).pack()
        wv = tk.IntVar(value=800); tk.Spinbox(win, from_=1, to=10000, textvariable=wv, width=8).pack()
        tk.Label(win, text="高度(px, 0=等比):", font=("微软雅黑",10)).pack(pady=5)
        hv = tk.IntVar(value=0); tk.Spinbox(win, from_=0, to=10000, textvariable=hv, width=8).pack()
        keep = tk.BooleanVar(value=True); tk.Checkbutton(win, text="保持比例", variable=keep).pack()
        def _go():
            try:
                nw, nh = wv.get(), hv.get(); ok = 0
                for fp in files:
                    img = Image.open(fp); ow, oh = img.size
                    if keep.get() and nh == 0: nh = int(oh * nw / ow)
                    r = img.resize((nw, nh))
                    out = Path(fp).parent / f"{Path(fp).stem}_resize{Path(fp).suffix}"
                    r.save(str(out)); ok += 1
                messagebox.showinfo("完成", f"缩放完成: {ok} 个文件")
            except Exception as e: messagebox.showerror("错误", str(e))
        tk.Button(win, text="开始缩放", command=_go, bg=self.colors['primary'], fg="white",
                 font=("微软雅黑",11,"bold"), cursor="hand2", width=12).pack(pady=15)

    def _watermark_dlg(self):
        files = filedialog.askopenfilenames(title="选择图片", filetypes=[("图片","*.png *.jpg")])
        if not files: return
        win = tk.Toplevel(self.root); win.title("水印"); win.geometry("400x300")
        win.transient(self.root); win.grab_set()
        tk.Label(win, text="水印文字:", font=("微软雅黑",10)).pack(pady=5)
        tv = tk.StringVar(value="万能办公助手")
        tk.Entry(win, textvariable=tv, width=25, font=("微软雅黑",10)).pack()
        tk.Label(win, text="位置:", font=("微软雅黑",10)).pack(pady=5)
        pv = ttk.Combobox(win, values=["左上","右上","左下","右下","居中"], state="readonly", width=8)
        pv.set("右下"); pv.pack()
        tk.Label(win, text="字号:", font=("微软雅黑",10)).pack(pady=5)
        sz = tk.IntVar(value=36); tk.Spinbox(win, from_=8, to=200, textvariable=sz, width=6).pack()
        def _go():
            try:
                text = tv.get(); size = sz.get(); pos = pv.get(); ok = 0
                for fp in files:
                    img = Image.open(fp).convert("RGBA")
                    overlay = Image.new("RGBA", img.size, (0,0,0,0))
                    draw = ImageDraw.Draw(overlay)
                    try:
                        from PIL import ImageFont
                        font = ImageFont.truetype("msyh.ttc", size)
                    except Exception:
                        font = ImageFont.load_default()
                    bbox = draw.textbbox((0,0), text, font=font)
                    tw, th = bbox[2]-bbox[0], bbox[3]-bbox[1]
                    margin = 20
                    xy = {"左上":(margin,margin), "右上":(img.width-tw-margin,margin),
                          "左下":(margin,img.height-th-margin), "右下":(img.width-tw-margin,img.height-th-margin),
                          "居中":((img.width-tw)//2,(img.height-th)//2)}.get(pos, (margin,margin))
                    draw.text(xy, text, fill=(255,255,255,180), font=font)
                    watermarked = Image.alpha_composite(img, overlay)
                    out = Path(fp).parent / f"{Path(fp).stem}_wm{Path(fp).suffix}"
                    watermarked.convert("RGB").save(str(out)); ok += 1
                messagebox.showinfo("完成", f"水印添加完成: {ok} 个")
            except Exception as e: messagebox.showerror("错误", str(e))
        tk.Button(win, text="添加水印", command=_go, bg=self.colors['primary'], fg="white",
                 font=("微软雅黑",11,"bold"), cursor="hand2", width=12).pack(pady=15)

    def _concat_img_dlg(self):
        files = filedialog.askopenfilenames(title="选择图片(按顺序拼接)", filetypes=[("图片","*.png *.jpg")])
        if not files: return
        win = tk.Toplevel(self.root); win.title("拼接"); win.geometry("350x150")
        win.transient(self.root); win.grab_set()
        dv = tk.StringVar(value="横向"); ttk.Combobox(win, textvariable=dv, values=["横向","纵向"], state="readonly", width=8).pack(pady=15)
        def _go():
            try:
                imgs = [Image.open(fp) for fp in files]
                direction = dv.get(); save = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG","*.png")])
                if not save: return
                if direction == "横向":
                    mw = sum(i.width for i in imgs); mh = max(i.height for i in imgs)
                    dst = Image.new("RGB", (mw, mh)); x = 0
                    for img in imgs: dst.paste(img, (x,0)); x += img.width
                else:
                    mw = max(i.width for i in imgs); mh = sum(i.height for i in imgs)
                    dst = Image.new("RGB", (mw, mh)); y = 0
                    for img in imgs: dst.paste(img, (0,y)); y += img.height
                dst.save(save)
                self.root.after(0, lambda: self._show_success_dialog("完成", f"拼接成功\n{save}"))
            except Exception as e: messagebox.showerror("错误", str(e))
        tk.Button(win, text="拼接并保存", command=_go, bg=self.colors['primary'], fg="white",
                 font=("微软雅黑",11,"bold"), cursor="hand2", width=12).pack(pady=10)

    def _compress_img_dlg(self):
        files = filedialog.askopenfilenames(title="选择图片", filetypes=[("图片","*.png *.jpg *.jpeg")])
        if not files: return
        win = tk.Toplevel(self.root); win.title("压缩"); win.geometry("450x300")
        win.transient(self.root); win.grab_set()
        tk.Label(win, text="JPEG质量(1-100):", font=("微软雅黑",10)).pack(pady=5)
        qv = tk.IntVar(value=60); tk.Scale(win, from_=1, to=100, variable=qv, orient=tk.HORIZONTAL, length=300).pack()
        tk.Label(win, text="最大宽度(px, 0=不限):", font=("微软雅黑",10)).pack(pady=5)
        mw = tk.IntVar(value=0); tk.Spinbox(win, from_=0, to=10000, textvariable=mw, width=8).pack()
        tk.Label(win, text="最大高度(px, 0=不限):", font=("微软雅黑",10)).pack(pady=5)
        mh = tk.IntVar(value=0); tk.Spinbox(win, from_=0, to=10000, textvariable=mh, width=8).pack()
        def _go():
            try:
                ok = 0; saved = 0
                for fp in files:
                    img = Image.open(fp); ow, oh = img.size; q = qv.get()
                    max_w, max_h = mw.get(), mh.get()
                    if max_w > 0 and ow > max_w: img = img.resize((max_w, int(oh*max_w/ow)))
                    if max_h > 0 and oh > max_h: img = img.resize((int(ow*max_h/oh), max_h))
                    orig = os.path.getsize(fp)
                    out = Path(fp).parent / f"{Path(fp).stem}_compressed{Path(fp).suffix}"
                    fmt = Path(fp).suffix.lower().replace(".","")
                    if fmt in ("jpg","jpeg"): img.save(str(out), "JPEG", quality=q)
                    else: img.save(str(out), optimize=True)
                    saved += orig - os.path.getsize(str(out)); ok += 1
                messagebox.showinfo("完成", f"压缩完成: {ok} 个\n节省: {saved/1024:.1f}KB")
            except Exception as e: messagebox.showerror("错误", str(e))
        tk.Button(win, text="开始压缩", command=_go, bg=self.colors['primary'], fg="white",
                 font=("微软雅黑",11,"bold"), cursor="hand2", width=12).pack(pady=10)

    def _ocr_img_dlg(self):
        files = filedialog.askopenfilenames(title="选择图片", filetypes=[("图片","*.png *.jpg *.bmp")])
        if not files: return
        win = tk.Toplevel(self.root); win.title("OCR识别"); win.geometry("600x400")
        win.transient(self.root); win.grab_set()
        log = scrolledtext.ScrolledText(win, height=10, font=("微软雅黑",10)); log.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        tk.Label(win, text="点击识别按钮，识别结果将显示在上方", font=("微软雅黑",9), fg="gray").pack()
        def _do_ocr():
            try:
                import pytesseract
                for fp in files:
                    img = Image.open(fp)
                    text = pytesseract.image_to_string(img, lang="chi_sim+eng")
                    log.insert(tk.END, f"📄 {Path(fp).name}:\n{text}\n{'-'*40}\n")
                    log.see(tk.END)
            except ImportError:
                guide = get_ocr_install_guide()
                messagebox.showinfo("OCR未安装", guide)
            except Exception as e: log.insert(tk.END, f"❌ 错误: {e}\n")
        def _save():
            text = log.get("1.0", tk.END).strip()
            if text:
                save = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("文本","*.txt")])
                if save: Path(save).write_text(text, encoding="utf-8")
        tk.Button(win, text="📝 识别", command=_do_ocr, bg=self.colors['primary'], fg="white",
                 font=("微软雅黑",11,"bold"), cursor="hand2", width=10).pack(side=tk.LEFT, padx=10, pady=8)
        tk.Button(win, text="💾 保存", command=_save, cursor="hand2", font=("微软雅黑",9), width=8).pack(side=tk.LEFT)

    def _nine_grid_dlg(self):
        files = filedialog.askopenfilenames(title="选择图片", filetypes=[("图片","*.png *.jpg")])
        if not files: return
        save_dir = filedialog.askdirectory(title="保存目录")
        if not save_dir: return
        def _go():
            try:
                for fp in files:
                    img = Image.open(fp); w, h = img.size
                    sw, sh = w//3, h//3
                    for i in range(3):
                        for j in range(3):
                            tile = img.crop((i*sw, j*sh, (i+1)*sw, (j+1)*sh))
                            out = Path(save_dir) / f"{Path(fp).stem}_{j*3+i+1}{Path(fp).suffix}"
                            tile.save(str(out))
                messagebox.showinfo("完成", f"九宫格切割完成\n{len(files)*9} 张\n{save_dir}")
            except Exception as e: messagebox.showerror("错误", str(e))
        self._run_thread(_go, done_msg="九宫格完成")
