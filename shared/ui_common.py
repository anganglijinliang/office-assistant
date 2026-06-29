# -*- coding: utf-8 -*-
"""万能办公助手 — BaseUIMixin（基础 UI 方法，供所有 Mixin 和主入口继承）"""
import tkinter as tk
from tkinter import ttk, messagebox
import threading
from typing import Optional, Callable, Any

from utils import ProgressDialog


class BaseUIMixin:
    """BaseUIMixin — 基础 UI 方法集。

    所有方法通过 self 访问 OfficeAssistant 的属性：
      - self.content_frame   : tk.Frame  主内容区域
      - self.colors          : dict      颜色字典（primary/dark/light/gray/danger）
      - self.root            : tk.Tk     根窗口
      - self.set_status(str) : Callable  状态栏更新
    """

    # ── 内容区清理 ──────────────────────────────────────────────────────

    def clear_content(self) -> None:
        """清空主内容区 (self.content_frame) 中的所有子控件。"""
        try:
            for w in self.content_frame.winfo_children():
                try:
                    # 尝试销毁所有子控件
                    if hasattr(w, 'winfo_children'):
                        for child in w.winfo_children():
                            child.destroy()
                    w.destroy()
                except Exception:
                    pass
        except Exception:
            pass

    # ── 卡片创建 ────────────────────────────────────────────────────────

    def _create_card(self, parent: tk.Frame, title: str,
                     desc: str, callback: Callable[[], Any],
                     btn_text: str = "") -> None:
        """在 parent 容器中创建一个带标题+描述的可点击卡片。

        参数:
            parent:   父容器 Frame
            title:    卡片标题（可含 emoji）
            desc:     卡片描述文字
            callback: 点击回调函数
            btn_text: 可选按钮文字（如 "开始"），提供则卡片底部显示按钮
        """
        card = tk.Frame(
            parent, bg="white", relief=tk.GROOVE, bd=1,
            cursor="hand2", highlightthickness=0,
        )
        card.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=6, pady=4, ipadx=4, ipady=4)

        # 标题
        title_lbl = tk.Label(
            card, text=title, font=("微软雅黑", 12, "bold"),
            bg="white", fg=self.colors.get('dark', '#0F172A'),
            anchor="w", padx=8, pady=10,
        )
        title_lbl.pack(fill=tk.X)

        # 描述
        desc_lbl = tk.Label(
            card, text=desc, font=("微软雅黑", 9),
            bg="white", fg=self.colors.get('gray', '#64748B'),
            anchor="w", padx=8, pady=10,
        )
        desc_lbl.pack(fill=tk.X)

        # 可选按钮
        btn = None
        if btn_text:
            btn = tk.Button(
                card, text=btn_text, cursor="hand2",
                font=("微软雅黑", 9), bd=0, padx=8, pady=2,
                bg=self.colors.get('primary', '#4F46E5'),
                fg="white", command=callback,
                activebackground=self.colors.get('primary_hover', '#4338CA'),
                activeforeground="white",
            )
            btn.pack(anchor="e", padx=(0, 8), pady=(0, 6))

        # 绑定点击事件到整个卡片（有按钮时不覆盖按钮点击区域）
        click_widgets = (card, title_lbl, desc_lbl)
        if btn:
            click_widgets = (card, title_lbl, desc_lbl)
        for widget in click_widgets:
            widget.bind("<Button-1>", lambda e, cb=callback: cb())
            widget.bind("<Enter>", lambda e: (
                card.configure(bg="#F1F5F9"),
                title_lbl.configure(bg="#F1F5F9"),
                desc_lbl.configure(bg="#F1F5F9"),
            ))
            widget.bind("<Leave>", lambda e: (
                card.configure(bg="white"),
                title_lbl.configure(bg="white"),
                desc_lbl.configure(bg="white"),
            ))

    # ── 段落标题 ────────────────────────────────────────────────────────

    def _section_header(self, title: str, subtitle: str = "") -> None:
        """在内容区顶部添加段落标题+副标题。

        参数:
            title:    主标题
            subtitle: 副标题（灰色，可选）
        """
        header_frame = tk.Frame(self.content_frame, bg=self.colors.get('light', '#F8FAFC'))
        header_frame.pack(fill=tk.X, padx=20, pady=(10, 2))

        tk.Label(
            header_frame, text=title, font=("微软雅黑", 16, "bold"),
            bg=self.colors.get('light', '#F8FAFC'),
            fg=self.colors.get('dark', '#0F172A'),
        ).pack(anchor="w")

        if subtitle:
            tk.Label(
                header_frame, text=subtitle, font=("微软雅黑", 9),
                bg=self.colors.get('light', '#F8FAFC'),
                fg=self.colors.get('gray', '#64748B'),
            ).pack(anchor="w", pady=(0, 4))

        # 分隔线
        sep = tk.Frame(
            self.content_frame, height=2,
            bg=self.colors.get('primary', '#4F46E5'),
        )
        sep.pack(fill=tk.X, padx=20, pady=(0, 8))

    # ── 线程执行 ────────────────────────────────────────────────────────

    def _run_thread(self, target: Callable, *,
                    args: tuple = (),
                    callback: Optional[Callable[[Any], None]] = None,
                    daemon: bool = True,
                    **kwargs) -> threading.Thread:
        """在后台线程中执行 target，可选完成回调。

        参数:
            target:   目标函数
            args:     传给 target 的位置参数
            callback: 线程完成后的回调（接收 target 返回值）
            daemon:   是否设为守护线程

        返回:
            threading.Thread 实例（已 start）
        """
        result: list = [None]

        def _wrapper():
            try:
                result[0] = target(*args)
            except Exception as e:
                result[0] = e
            finally:
                if callback is not None:
                    try:
                        self.root.after(0, lambda: callback(result[0]))
                    except Exception:
                        pass

        t = threading.Thread(target=_wrapper, daemon=daemon)
        t.start()
        return t

    # ── 带进度条的执行 ──────────────────────────────────────────────────

    def _run_with_progress(self, target: Callable, *,
                           title: str = "处理中",
                           msg: str = "请稍候...",
                           args: tuple = (),
                           on_done: Optional[Callable[[bool, Any], None]] = None) -> None:
        """弹出进度对话框，在后台线程中执行 target。

        参数:
            target:  耗时函数（可返回结果或 raise 异常）
            title:   对话框标题
            msg:     初始状态文本
            args:    target 的位置参数
            on_done: 完成回调，签名 on_done(success: bool, result_or_error: Any)
        """
        dlg = ProgressDialog(self.root, title=title, msg=msg, progress=0)

        def _work():
            try:
                dlg.update("正在执行...", 30)
                result = target(dlg, *args)
                dlg.update("即将完成...", 80)
                # 短暂延迟让用户看到进度变化
                import time
                time.sleep(0.3)
                dlg.finish(ok=True, extra="已完成")
                if on_done:
                    self.root.after(0, lambda: on_done(True, result))
            except Exception as e:
                dlg.finish(ok=False, fail=str(e))
                if on_done:
                    self.root.after(0, lambda: on_done(False, e))
            finally:
                self.root.after(500, dlg.close)

        t = threading.Thread(target=_work, daemon=True)
        t.start()

    # ── 成功对话框 ──────────────────────────────────────────────────────

    def _show_success_dialog(self, title: str = "成功",
                             message: str = "操作已完成！") -> None:
        """在根窗口上方显示成功信息对话框。"""
        try:
            messagebox.showinfo(title, message, parent=self.root)
        except Exception:
            pass

    # ── 提示信息 ────────────────────────────────────────────────────────

    def _show_tips(self, *tips: str) -> None:
        """在内容区顶部显示淡蓝色提示信息。

        参数:
            *tips: 提示文本（每条一行，带 emoji 更佳）
        """
        if not tips:
            return

        tip_frame = tk.Frame(
            self.content_frame,
            bg="#EFF6FF",  # 淡蓝底色
            relief=tk.FLAT, bd=0,
            highlightthickness=0,
        )
        tip_frame.pack(fill=tk.X, padx=20, pady=(4, 8))

        # 左侧图标
        tk.Label(
            tip_frame, text="💡", font=("微软雅黑", 12),
            bg="#EFF6FF", fg="#3B82F6",
        ).pack(side=tk.LEFT, padx=(12, 4), pady=8)

        # 文本内容
        text_frame = tk.Frame(tip_frame, bg="#EFF6FF")
        text_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, pady=8)

        for tip in tips:
            tk.Label(
                text_frame, text=tip,
                font=("微软雅黑", 9),
                bg="#EFF6FF",
                fg="#1E40AF",  # 深蓝文字
                anchor="w", wraplength=700,
            ).pack(fill=tk.X, padx=(0, 12))
