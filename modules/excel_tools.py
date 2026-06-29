# -*- coding: utf-8 -*-
"""万能办公助手 — ExcelTools（合并·拆分·筛选·统计·导出·图表·VLOOKUP·条件格式）"""
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
from pathlib import Path
import os, sys, json, time, shutil, hashlib, threading
from datetime import datetime
from collections import Counter

from utils import OPENPYL_AVAILABLE, _OPENPYL_ERROR, safe_str, safe_cond_check

import openpyxl
from openpyxl.styles import PatternFill


class ExcelToolsMixin:
    """ExcelTools — 所有方法通过self访问OfficeAssistant的属性"""

    def _show_excel_tools(self):
        self.clear_content()
        self._section_header("Excel处理工具", "合并 · 拆分 · 筛选 · 统计 · 导出 · 图表 · VLOOKUP · 条件格式")
        self._show_tips(
            "点击下方卡片 → 选择Excel文件(.xlsx) → 自动处理 → 保存结果",
            "支持 .xlsx 格式，旧版 .xls 请先用Excel另存为 .xlsx"
        )
        if not OPENPYL_AVAILABLE:
            tk.Label(self.content_frame, text=f"⚠️ openpyxl未安装\n{_OPENPYL_ERROR}", fg="red",
                    font=("微软雅黑", 12), bg=self.colors['light']).pack(pady=60)
            return
        for rdata in [
            [("📊 合并Excel", "多文件纵向合并", self._merge_excel_dlg),
             ("✂ 拆分Excel", "按行数拆分为多个", self._split_excel_dlg),
             ("🔍 筛选数据", "按条件筛选/导出", self._filter_excel_dlg)],
            [("📈 数据统计", "最值/平均/计数/标准差", self._stats_excel_dlg),
             ("📊 导出CSV", "Excel→CSV (UTF-8)", self._excel_to_csv_dlg),
             ("📉 创建图表", "柱状图/折线图/饼图", self._excel_chart_dlg)],
            [("🔗 VLOOKUP", "跨表格匹配数据", self._vlookup_dlg),
             ("🎨 条件格式", "条件高亮/标记", self._conditional_format_dlg)],
        ]:
            row = tk.Frame(self.content_frame, bg=self.colors['light']); row.pack(fill=tk.X, padx=10)
            for title, desc, cb in rdata:
                self._create_card(row, title, desc, cb)

    def _merge_excel_dlg(self):
        files = filedialog.askopenfilenames(title="选择Excel文件(按顺序合并)", filetypes=[("Excel","*.xlsx")])
        if not files: return
        save = filedialog.asksaveasfilename(title="保存合并结果", defaultextension=".xlsx", filetypes=[("Excel","*.xlsx")])
        if not save: return
        def _go():
            try:
                wb_out = openpyxl.Workbook(); ws_out = wb_out.active; first = True
                for fp in files:
                    wb = openpyxl.load_workbook(fp, read_only=True); ws = wb.active
                    rows = list(ws.iter_rows(values_only=True)); wb.close()
                    if first:
                        ws_out.append([str(c) if c is not None else "" for c in rows[0]])
                        first = False
                    for r in rows[1:]:
                        ws_out.append([str(c) if c is not None else "" for c in r])
                wb_out.save(save); wb_out.close()
                self.root.after(0, lambda: self._show_success_dialog("完成", f"合并成功\n{len(files)}个文件\n{save}"))
            except Exception as e: self.root.after(0, lambda: messagebox.showerror("错误", str(e)))
        self._run_thread(_go, done_msg="合并完成")

    def _split_excel_dlg(self):
        fp = filedialog.askopenfilename(title="选择Excel", filetypes=[("Excel","*.xlsx")])
        if not fp: return
        out_dir = filedialog.askdirectory(title="输出目录")
        if not out_dir: return
        win = tk.Toplevel(self.root); win.title("拆分设置"); win.geometry("350x150")
        win.transient(self.root); win.grab_set()
        tk.Label(win, text="每多少个行一个文件:", font=("微软雅黑",10)).pack(pady=10)
        rows_per = tk.IntVar(value=1000)
        tk.Spinbox(win, from_=1, to=100000, textvariable=rows_per, width=10).pack()
        def _work():
            win.destroy()
            try:
                wb = openpyxl.load_workbook(fp, read_only=True); all_rows = list(wb.active.iter_rows(values_only=True)); wb.close()
                if not all_rows: return
                hdr, src = all_rows[0], all_rows[1:]; stem = Path(fp).stem; i, op = 0, 0
                while i < len(src):
                    chunk = [hdr] + src[i:i+rows_per.get()]; i += rows_per.get(); op += 1
                    wb2 = openpyxl.Workbook(); ws2 = wb2.active
                    for r in chunk: ws2.append([str(c) if c is not None else "" for c in r])
                    wb2.save(str(Path(out_dir) / f"{stem}_part{op}.xlsx")); wb2.close()
                self.root.after(0, lambda: self._show_success_dialog("完成", f"拆分为 {op} 个文件\n{out_dir}"))
            except Exception as e: self.root.after(0, lambda: messagebox.showerror("错误", str(e)))
        tk.Button(win, text="开始拆分", command=_work, bg=self.colors['primary'], fg="white",
                 font=("微软雅黑",11,"bold"), cursor="hand2", width=12).pack(pady=10)

    def _filter_excel_dlg(self):
        fp = filedialog.askopenfilename(title="选择Excel", filetypes=[("Excel","*.xlsx")])
        if not fp: return
        win = tk.Toplevel(self.root); win.title("筛选数据"); win.geometry("600x450")
        win.transient(self.root); win.grab_set()
        wb = openpyxl.load_workbook(fp, read_only=True); all_rows = list(wb.active.iter_rows(values_only=True)); wb.close()
        hdr = [str(c) for c in all_rows[0]] if all_rows else []
        log = scrolledtext.ScrolledText(win, height=12, font=("Consolas",9)); log.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        lb = tk.Listbox(win, height=4, font=("微软雅黑",9)); lb.pack(fill=tk.X, padx=10)
        for h in hdr: lb.insert(tk.END, h)
        e = tk.Entry(win, font=("微软雅黑",10)); e.pack(fill=tk.X, padx=10, pady=5)
        e.insert(0, ">0"); lb.selection_set(0)
        def _go():
            try:
                ci = lb.curselection(); ci = ci[0] if ci else 0
                cond = e.get().strip(); fltd = [all_rows[0]]
                ve = tk.StringVar(); h = hdr[ci]
                for r in all_rows[1:]:
                    v = r[ci]; val = float(v) if v is not None and str(v).replace('.','').replace('-','').isdigit() else 0
                    if safe_cond_check(val, cond): fltd.append(r)
                save = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel","*.xlsx")])
                if save:
                    wb2 = openpyxl.Workbook(); ws2 = wb2.active
                    for r in fltd: ws2.append([str(c) if c is not None else "" for c in r])
                    wb2.save(save); wb2.close()
                log.delete("1.0", tk.END)
                log.insert(tk.END, f"原始: {len(all_rows)-1} 行 | 筛选后: {len(fltd)-1} 行\n")
                for r in fltd[:30]: log.insert(tk.END, f"{[str(c)[:15] for c in r]}\n")
            except Exception as e: messagebox.showerror("错误", str(e))
        tk.Button(win, text="筛选并导出", command=_go, bg=self.colors['primary'], fg="white",
                 font=("微软雅黑",11,"bold"), cursor="hand2").pack(pady=5)

    def _stats_excel_dlg(self):
        fp = filedialog.askopenfilename(title="选择Excel", filetypes=[("Excel","*.xlsx")])
        if not fp: return
        win = tk.Toplevel(self.root); win.title("数据统计"); win.geometry("600x400")
        win.transient(self.root); win.grab_set()
        try:
            wb = openpyxl.load_workbook(fp, read_only=True); all_rows = list(wb.active.iter_rows(values_only=True)); wb.close()
        except Exception as e:
            messagebox.showerror("错误", f"无法读取文件: {e}"); return
        hdr = [str(c) for c in all_rows[0]] if all_rows else []
        log = scrolledtext.ScrolledText(win, height=15, font=("Consolas",9)); log.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        try:
            for ci in range(len(hdr)):
                vals = []; n = 0
                for r in all_rows[1:]:
                    v = r[ci]
                    if v is not None:
                        try: vals.append(float(v)); n += 1
                        except: pass
                if vals:
                    mn, mx = min(vals), max(vals)
                    avg = sum(vals)/len(vals); sd = (sum((x-avg)**2 for x in vals)/len(vals))**0.5
                    log.insert(tk.END, f"【{hdr[ci]}】 n={n}  平均={avg:.2f}  最值={mn}-{mx}  标准差={sd:.2f}\n\n")
            tk.Label(win, text="统计完成", font=("微软雅黑",9), fg="gray").pack()
        except Exception as e:
            log.insert(tk.END, f"❌ 统计出错: {e}\n")

    def _excel_to_csv_dlg(self):
        fp = filedialog.askopenfilename(title="选择Excel", filetypes=[("Excel","*.xlsx")])
        if not fp: return
        save = filedialog.asksaveasfilename(title="保存CSV", defaultextension=".csv", filetypes=[("CSV","*.csv")])
        if not save: return
        def _go():
            try:
                wb = openpyxl.load_workbook(fp, read_only=True); ws = wb.active
                import csv
                with open(save, 'w', newline='', encoding='utf-8-sig') as f:
                    w = csv.writer(f)
                    for row in ws.iter_rows(values_only=True):
                        w.writerow([str(c) if c is not None else "" for c in row])
                wb.close()
                self.root.after(0, lambda: self._show_success_dialog("完成", f"导出CSV成功\n{save}"))
            except Exception as e: self.root.after(0, lambda: messagebox.showerror("错误", str(e)))
        self._run_thread(_go, done_msg="导出完成")

    def _excel_chart_dlg(self):
        fp = filedialog.askopenfilename(title="选择Excel", filetypes=[("Excel","*.xlsx")])
        if not fp: return
        win = tk.Toplevel(self.root); win.title("创建图表"); win.geometry("400x280")
        win.transient(self.root); win.grab_set()
        wb = openpyxl.load_workbook(fp, read_only=True); all_rows = list(wb.active.iter_rows(values_only=True)); wb.close()
        hdr = [str(c) for c in all_rows[0]] if all_rows else []
        tk.Label(win, text="选择图表类型:", font=("微软雅黑",10)).pack(pady=10)
        ct = ttk.Combobox(win, values=["柱状图","折线图","饼图","条形图"], state="readonly", width=12); ct.set("柱状图"); ct.pack()
        tk.Label(win, text="分类列:", font=("微软雅黑",10)).pack(pady=5)
        c = ttk.Combobox(win, values=hdr, state="readonly", width=16)
        if hdr: c.set(hdr[0])
        c.pack()
        tk.Label(win, text="数值列:", font=("微软雅黑",10)).pack(pady=5)
        v = ttk.Combobox(win, values=hdr, state="readonly", width=16)
        if len(hdr)>1: v.set(hdr[1])
        v.pack()
        def _go():
            try:
                ci = hdr.index(c.get()); vi = hdr.index(v.get())
                data = []; cats = []
                for r in all_rows[1:]:
                    if r[ci] is not None and r[vi] is not None:
                        cats.append(str(r[ci])); data.append(float(r[vi]))
                wb2 = openpyxl.Workbook(); ws2 = wb2.active
                ws2.append([hdr[ci], hdr[vi]])
                for i in range(len(cats)): ws2.append([cats[i], data[i]])
                chart_type = ct.get()
                if "柱状" in chart_type:
                    from openpyxl.chart import BarChart as Chart
                    chart = Chart()
                elif "折线" in chart_type:
                    from openpyxl.chart import LineChart as Chart
                    chart = Chart()
                elif "饼图" in chart_type:
                    from openpyxl.chart import PieChart as Chart
                    chart = Chart()
                else:
                    from openpyxl.chart import BarChart as Chart
                    chart = Chart(); chart.type = "bar"
                from openpyxl.chart import Reference
                data_ref = Reference(ws2, min_col=2, min_row=1, max_row=len(cats)+1)
                cats_ref = Reference(ws2, min_col=1, min_row=2, max_row=len(cats)+1)
                chart.add_data(data_ref, titles_from_data=True)
                chart.set_categories(cats_ref); chart.title = f"{hdr[vi]} {chart_type}"
                ws2.add_chart(chart, "E5")
                save = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel","*.xlsx")])
                if save: wb2.save(save)
                wb2.close()
                if save: self.root.after(0, lambda: self._show_success_dialog("完成", f"图表创建成功\n{save}"))
            except Exception as e: messagebox.showerror("错误", str(e))
        tk.Button(win, text="生成图表", command=_go, bg=self.colors['primary'], fg="white",
                 font=("微软雅黑",11,"bold"), cursor="hand2", width=12).pack(pady=12)

    def _vlookup_dlg(self):
        files = filedialog.askopenfilenames(title="选择两个Excel（主表+查询表）", filetypes=[("Excel","*.xlsx")])
        if len(files) != 2: return messagebox.showwarning("提示", "请选择两个文件")
        win = tk.Toplevel(self.root); win.title("VLOOKUP"); win.geometry("500x300")
        win.transient(self.root); win.grab_set()
        wb1 = openpyxl.load_workbook(files[0], read_only=True); rows1 = list(wb1.active.iter_rows(values_only=True)); wb1.close()
        wb2 = openpyxl.load_workbook(files[1], read_only=True); rows2 = list(wb2.active.iter_rows(values_only=True)); wb2.close()
        hdr1 = [str(c) for c in rows1[0]]; hdr2 = [str(c) for c in rows2[0]]
        tk.Label(win, text="主表匹配列:", font=("微软雅黑",10)).pack()
        key = ttk.Combobox(win, values=hdr1, state="readonly", width=16); key.set(hdr1[0] if hdr1 else ""); key.pack()
        tk.Label(win, text="查询表匹配列:", font=("微软雅黑",10)).pack(pady=5)
        ex = ttk.Combobox(win, values=hdr2, state="readonly", width=16); ex.set(hdr2[0] if hdr2 else ""); ex.pack()
        def _go():
            try:
                ki = hdr1.index(key.get()); ei = hdr2.index(ex.get())
                key_map = {str(r[ei]): [str(c) for c in r] for r in rows2[1:] if r[ei] is not None}
                wb_out = openpyxl.Workbook(); ws_out = wb_out.active
                ws_out.append(hdr1 + [f"{Path(files[1]).stem}_{c}" for c in hdr2])
                for r in rows1[1:]:
                    match_val = str(r[ki]) if r[ki] is not None else ""
                    match_row = key_map.get(match_val, [""]*len(hdr2))
                    ws_out.append([str(c) if c is not None else "" for c in r] + match_row)
                save = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel","*.xlsx")])
                if save: wb_out.save(save)
                wb_out.close()
                if save: self.root.after(0, lambda: self._show_success_dialog("完成", f"VLOOKUP完成\n{save}"))
            except Exception as e: messagebox.showerror("错误", str(e))
        tk.Button(win, text="执行VLOOKUP", command=_go, bg=self.colors['primary'], fg="white",
                 font=("微软雅黑",11,"bold"), cursor="hand2", width=14).pack(pady=12)

    def _conditional_format_dlg(self):
        fp = filedialog.askopenfilename(title="选择Excel", filetypes=[("Excel","*.xlsx")])
        if not fp: return
        win = tk.Toplevel(self.root); win.title("条件格式"); win.geometry("500x350")
        win.transient(self.root); win.grab_set()
        wb = openpyxl.load_workbook(fp, read_only=True); all_rows = list(wb.active.iter_rows(values_only=True)); wb.close()
        hdr = [str(c) for c in all_rows[0]] if all_rows else []
        tk.Label(win, text="选择要格式化的列:", font=("微软雅黑",10)).pack(pady=8)
        lb = tk.Listbox(win, height=5, font=("微软雅黑",9)); lb.pack(fill=tk.X, padx=20)
        for h in hdr: lb.insert(tk.END, h); lb.selection_set(0)
        tk.Label(win, text="条件 (>0, <100, ==0, !=0):", font=("微软雅黑",10)).pack(pady=5)
        e = tk.Entry(win, font=("微软雅黑",10), width=20); e.insert(0, "<60"); e.pack()
        def _go():
            try:
                sel = lb.curselection(); ci = sel[0] if sel else 0
                cond = e.get().strip(); cond_v = cond[2:].strip() if cond[:2] in (">=","<=","==","!=") else cond[1:].strip()
                wb = openpyxl.load_workbook(fp); ws = wb.active
                fill = PatternFill(start_color='FF6B6B', end_color='FF6B6B', fill_type='solid')
                for row in ws.iter_rows(min_row=2, max_col=ws.max_column):
                    val = row[ci].value
                    if val is not None:
                        try:
                            nv = float(val)
                            ex = cond
                            if ex.startswith(">="): ok = nv >= float(ex[2:])
                            elif ex.startswith("<="): ok = nv <= float(ex[2:])
                            elif ex.startswith("=="): ok = nv == float(ex[2:])
                            elif ex.startswith("!="): ok = nv != float(ex[2:])
                            elif ex.startswith(">"): ok = nv > float(ex[1:])
                            elif ex.startswith("<"): ok = nv < float(ex[1:])
                            else: ok = False
                            if ok: row[ci].fill = fill
                        except: pass
                save = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel","*.xlsx")])
                if save: wb.save(save)
                wb.close()
                if save: self.root.after(0, lambda: self._show_success_dialog("完成", f"条件格式完成\n{save}"))
            except Exception as e: messagebox.showerror("错误", str(e))
        tk.Button(win, text="应用", command=_go, bg=self.colors['primary'], fg="white",
                 font=("微软雅黑",11,"bold"), cursor="hand2", width=12).pack(pady=12)
