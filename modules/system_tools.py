# -*- coding: utf-8 -*-
"""万能办公助手 — SystemTools（系统托盘 + DPI）"""
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
from pathlib import Path
import os, sys, json, time, shutil, hashlib, threading
from datetime import datetime
from collections import Counter


class SystemToolsMixin:
    """SystemTools — 所有方法通过self访问OfficeAssistant的属性"""

    def _minimize_to_tray(self):
        """最小化到系统托盘"""
        try:
            import pystray
            from PIL import Image as _PIL_Img
            from io import BytesIO
            if self._tray_icon and self._tray_visible:
                return
            img = _PIL_Img.new('RGBA', (64, 64), (79, 70, 229, 255))
            from PIL import ImageDraw, ImageFont
            draw = ImageDraw.Draw(img)
            try:
                font = ImageFont.truetype("msyh.ttc", 36)
            except Exception:
                font = ImageFont.load_default()
            draw.text((8, 4), "办", fill=(255, 255, 255, 255), font=font)
            def _on_show(icon, item):
                self.root.after(0, self._restore_from_tray)
            def _on_show_click(icon, item):
                self.root.after(0, self._restore_from_tray)
            def _on_exit(icon, item):
                self.root.after(0, self._quit_app)
            menu = pystray.Menu(
                pystray.MenuItem("显示窗口", _on_show, default=True),
                pystray.MenuItem("退出", _on_exit),
            )
            self._tray_icon = pystray.Icon("office_assistant", img, "万能办公助手 v6.2", menu)
            self._tray_visible = True
            self.root.withdraw()
            threading.Thread(target=self._tray_icon.run, daemon=True).start()
        except Exception as e:
            messagebox.showerror("托盘错误", f"无法最小化到托盘:\n{e}")

    def _restore_from_tray(self):
        """从系统托盘恢复窗口"""
        try:
            if self._tray_icon:
                self._tray_icon.stop()
                self._tray_icon = None
            self._tray_visible = False
            self.root.deiconify()
            self.root.lift()
            self.root.focus_force()
        except Exception:
            pass

    def _quit_app(self):
        """退出程序（安全清理后销毁窗口）"""
        # 1) 保存窗口几何
        try:
            self._save_window_geometry()
        except Exception:
            pass
        # 2) 停止托盘图标（不在 GUI 线程则只标记，避免阻塞）
        if self._tray_icon is not None:
            try:
                import threading
                if threading.current_thread() is threading.main_thread():
                    self._tray_icon.stop()
                else:
                    # 从非主线程调用 -> 用 after 调度
                    self.root.after(0, self._tray_icon.stop)
            except Exception:
                pass
            self._tray_icon = None
        self._tray_visible = False
        # 3) 销毁窗口
        try:
            self.root.destroy()
        except Exception:
            import os
            os._exit(0)

    def _get_dpi_scale(self):
        """获取DPI缩放因子"""
        try:
            import ctypes
            hdc = ctypes.windll.user32.GetDC(0)
            dpi = ctypes.windll.gdi32.GetDeviceCaps(hdc, 88)
            ctypes.windll.user32.ReleaseDC(0, hdc)
            return max(1.0, min(dpi / 96.0, 2.5))
        except Exception:
            return 1.0
