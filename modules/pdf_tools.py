# -*- coding: utf-8 -*-
"""万能办公助手 — PdfTools（提取文本·合并·拆分·加密·解密·PDF→Word）"""
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
from pathlib import Path
import os, sys, json, time, shutil, hashlib, threading
from datetime import datetime
from collections import Counter

from utils import PDF_AVAILABLE, _PDF_ERROR, WORD_COM_AVAILABLE

import PyPDF2


class PdfToolsMixin:
    """PdfTools — 所有方法通过self访问OfficeAssistant的属性"""

    def _show_pdf_tools(self):
        self.clear_content()
        self._section_header("PDF工具", "提取文本 · 合并 · 拆分 · 加密 · 解密 · PDF→Word")
        self._show_tips(
            "选功能 → 选择PDF文件 → 设置参数 → 自动处理",
            "PDF→Word 适合文字型PDF，扫描件请先用OCRSD后再转换"
        )
        if not PDF_AVAILABLE:
            err_msg = _PDF_ERROR
            tk.Label(self.content_frame, text=f"⚠️ PDF模块加载失败\n\n错误信息:\n{err_msg}\n\n请确保 PyPDF2 已正确安装",
                    fg="red", font=("微软雅黑", 12), bg=self.colors['light'], justify=tk.LEFT).pack(pady=60)
            return
        for rdata in [
            [("📄 PDF→文本", "提取纯文本内容", self._pdf_to_text_dlg),
             ("📑 合并PDF", "多文件合并为一个", self._merge_pdf_dlg),
             ("✂ 拆分PDF", "按页拆分为多个", self._split_pdf_dlg)],
            [("🔒 加密PDF", "设置打开密码", self._encrypt_pdf_dlg),
             ("🔓 解密PDF", "移除密码保护", self._decrypt_pdf_dlg),
             ("📝 PDF→Word", "转可编辑文档", self._pdf_to_word_dlg)],
        ]:
            row = tk.Frame(self.content_frame, bg=self.colors['light']); row.pack(fill=tk.X, padx=10)
            for title, desc, cb in rdata:
                self._create_card(row, title, desc, cb)

    def _pdf_to_text_dlg(self):
        fp = filedialog.askopenfilename(title="选择PDF", filetypes=[("PDF","*.pdf")])
        if not fp: return
        save = filedialog.asksaveasfilename(title="保存文本", defaultextension=".txt", filetypes=[("文本","*.txt")])
        if not save: return
        def _go(pd):
            try:
                pd.update("读取PDF...", progress=0)
                reader = PyPDF2.PdfReader(fp); total = len(reader.pages)
                all_text = []
                for i, page in enumerate(reader.pages):
                    t = page.extract_text()
                    all_text.append(t); all_text.append(f"\n--- 第 {i+1} 页 ---\n")
                    pd.update(f"提取第 {i+1}/{total} 页...", progress=(i+1)/max(total,1))
                Path(save).write_text("\n".join(all_text), encoding="utf-8")
                pd.finish(ok=total, extra=f"输出: {Path(save).name}")
                self.root.after(0, lambda: self._show_success_dialog("完成", f"PDF→文本 成功\n{total} 页\n{save}"))
            except Exception as e:
                import traceback; err = traceback.format_exc(); pd.finish(fail=1, extra=err[-200:])
                self.root.after(0, lambda: messagebox.showerror("错误", f"提取失败:\n{err[-600:]}"))
        self._run_with_progress("PDF→文本", _go)

    def _merge_pdf_dlg(self):
        files = filedialog.askopenfilenames(title="选择PDF文件（按顺序合并）", filetypes=[("PDF","*.pdf")])
        if not files: return
        save = filedialog.asksaveasfilename(title="保存合并结果", defaultextension=".pdf", filetypes=[("PDF","*.pdf")])
        if not save: return
        def _go(pd):
            try:
                merger = PyPDF2.PdfMerger()
                for i, fp in enumerate(files):
                    merger.append(fp)
                    pd.update(f"合并 {i+1}/{len(files)}...", progress=(i+1)/len(files))
                merger.write(save); merger.close()
                pd.finish(ok=len(files), extra=f"输出: {Path(save).name}")
                self.root.after(0, lambda: self._show_success_dialog("完成", f"合并成功\n{len(files)}个文件\n{save}"))
            except Exception as e:
                import traceback; err = traceback.format_exc(); pd.finish(fail=1, extra=err[-200:])
                self.root.after(0, lambda: messagebox.showerror("错误", f"合并失败:\n{err[-600:]}"))
        self._run_with_progress("合并PDF", _go)

    def _split_pdf_dlg(self):
        fp = filedialog.askopenfilename(title="选择PDF", filetypes=[("PDF","*.pdf")])
        if not fp: return
        out_dir = filedialog.askdirectory(title="输出目录")
        if not out_dir: return
        win = tk.Toplevel(self.root); win.title("拆分设置"); win.geometry("400x180")
        win.transient(self.root); win.grab_set()
        tk.Label(win, text="每多少页一个文件:", font=("微软雅黑",10)).pack(pady=10)
        pp = tk.IntVar(value=1)
        tk.Spinbox(win, from_=1, to=100, textvariable=pp, width=8).pack()
        def _work():
            win.destroy()
            try:
                reader = PyPDF2.PdfReader(fp); total = len(reader.pages); per = pp.get(); stem = Path(fp).stem
                for i in range(0, total, per):
                    writer = PyPDF2.PdfWriter()
                    for j in range(i, min(i+per, total)): writer.add_page(reader.pages[j])
                    writer.write(str(Path(out_dir) / f"{stem}_page{i+1}-{min(i+per,total)}.pdf")); writer.close()
                self.root.after(0, lambda: self._show_success_dialog("完成", f"拆分成功\n共 {total} 页\n{out_dir}"))
            except Exception as e: self.root.after(0, lambda: messagebox.showerror("错误", str(e)))
        tk.Button(win, text="开始拆分", command=_work, bg=self.colors['primary'], fg="white",
                 font=("微软雅黑",11,"bold"), cursor="hand2", width=12).pack(pady=15)

    def _encrypt_pdf_dlg(self):
        fp = filedialog.askopenfilename(title="选择PDF", filetypes=[("PDF","*.pdf")])
        if not fp: return
        win = tk.Toplevel(self.root); win.title("加密PDF"); win.geometry("400x200")
        win.transient(self.root); win.grab_set()
        tk.Label(win, text="设置打开密码:", font=("微软雅黑",11)).pack(pady=15)
        pwd = tk.Entry(win, font=("Consolas",12), width=20, show="*"); pwd.pack()
        def _work():
            win.destroy()
            try:
                reader = PyPDF2.PdfReader(fp); writer = PyPDF2.PdfWriter()
                for page in reader.pages: writer.add_page(page)
                writer.encrypt(pwd.get())
                save = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF","*.pdf")])
                if save:
                    with open(save, 'wb') as f: writer.write(f)
                    self.root.after(0, lambda: self._show_success_dialog("完成", f"加密成功\n{save}"))
            except Exception as e: self.root.after(0, lambda: messagebox.showerror("错误", str(e)))
        tk.Button(win, text="加密并保存", command=_work, bg=self.colors['primary'], fg="white",
                 font=("微软雅黑",11,"bold"), cursor="hand2", width=12).pack(pady=15)

    def _decrypt_pdf_dlg(self):
        fp = filedialog.askopenfilename(title="选择加密PDF", filetypes=[("PDF","*.pdf")])
        if not fp: return
        win = tk.Toplevel(self.root); win.title("解密PDF"); win.geometry("400x200")
        win.transient(self.root); win.grab_set()
        tk.Label(win, text="输入密码:", font=("微软雅黑",11)).pack(pady=15)
        pwd = tk.Entry(win, font=("Consolas",12), width=20, show="*"); pwd.pack()
        def _work():
            win.destroy()
            try:
                reader = PyPDF2.PdfReader(fp); reader.decrypt(pwd.get())
                writer = PyPDF2.PdfWriter()
                for page in reader.pages: writer.add_page(page)
                save = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF","*.pdf")])
                if save:
                    with open(save, 'wb') as f: writer.write(f)
                    self.root.after(0, lambda: self._show_success_dialog("完成", f"解密成功\n{save}"))
            except Exception as e: self.root.after(0, lambda: messagebox.showerror("错误", str(e)))
        tk.Button(win, text="解密并保存", command=_work, bg=self.colors['primary'], fg="white",
                 font=("微软雅黑",11,"bold"), cursor="hand2", width=12).pack(pady=15)

    def _pdf_to_word_dlg(self):
        """PDF → Word（优先 Word COM 引擎，备胎 pdf2docx）"""
        fp = filedialog.askopenfilename(title="选择PDF", filetypes=[("PDF","*.pdf")])
        if not fp: return
        save = filedialog.asksaveasfilename(title="保存Word文档", defaultextension=".docx",
                                             filetypes=[("Word","*.docx")])
        if not save: return
        def _go(pd):
            try:
                if WORD_COM_AVAILABLE:
                    # 方案A：Word COM（完美保留布局，需 Word 2013+）
                    pd.update("调用 Word 引擎转换...", progress=0)
                    import win32com.client
                    word = win32com.client.Dispatch("Word.Application")
                    word.Visible = False
                    try:
                        doc = word.Documents.Open(fp)
                        if doc is None:
                            raise RuntimeError("Word无法打开PDF文件（当前Word版本不支持PDF导入，需要Office 2013+）")
                        pd.update("生成Word...", progress=0.6)
                        doc.SaveAs(save, FileFormat=16)  # 16 = wdFormatDocumentDefault
                        doc.Close()
                        pd.finish(ok=1, extra=f"输出: {Path(save).name}")
                        self.root.after(0, lambda: self._show_success_dialog("完成",
                            f"PDF→Word 转换成功 (Word引擎)\n{save}"))
                    finally:
                        try: word.Quit()
                        except: pass
                else:
                    # 方案B：pdf2docx
                    pd.update("正在分析PDF...", progress=0)
                    from pdf2docx import Converter
                    cv = Converter(fp)
                    pd.update("转换中...", progress=0.3)
                    cv.convert(save, start=0, end=None)
                    cv.close()
                    pd.update("保存文件...", progress=0.9)
                    pd.finish(ok=1, extra=f"输出: {Path(save).name}")
                    self.root.after(0, lambda: self._show_success_dialog("完成",
                        f"PDF→Word 转换成功\n{save}"))
            except Exception as e:
                import traceback
                err = traceback.format_exc()
                # Word COM 失败时自动降级到 pdf2docx
                if WORD_COM_AVAILABLE:
                    try:
                        pd.update("Word引擎不可用，切换到pdf2docx...", progress=0)
                        from pdf2docx import Converter
                        cv = Converter(fp)
                        cv.convert(save, start=0, end=None)
                        cv.close()
                        pd.finish(ok=1, extra=f"输出: {Path(save).name} (pdf2docx)")
                        self.root.after(0, lambda: self._show_success_dialog("完成",
                            f"PDF→Word 转换成功 (pdf2docx)\n{save}"))
                        return
                    except Exception as e2:
                        err += f"\n降级也失败: {e2}"
                pd.finish(fail=1, extra=err[-200:])
                self.root.after(0, lambda: messagebox.showerror("错误", f"转换失败:\n{err[-600:]}"))
        self._run_with_progress("PDF→Word", _go)
