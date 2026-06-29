# -*- coding: utf-8 -*-
"""
万能办公助手 v6.2 — LicenseManager（许可证管理 + 加卸载安全防护）
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
职责：
  · 激活码验证（HMAC‑SHA256 自校验 16 位十六进制码）
  · ¥29.9/年 激活制（无试用）
  · 激活界面和状态栏交互
  · 卸载安全防护（防误删桌面/项目目录）
"""
from __future__ import annotations

import hashlib
import hmac
import json
import os
import re
import shutil
import subprocess
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path
from tkinter import messagebox, simpledialog, ttk
from typing import Callable, Optional, Tuple

import tkinter as tk

# ── 路径常量 ──────────────────────────────────────────────
# 许可证文件存在 %APPDATA%/OfficeAssistant_v6.2/license.dat
DATA_DIR = Path(os.environ.get("APPDATA", "")) / "OfficeAssistant_v6.2"
LICENSE_FILE = DATA_DIR / "license.dat"

# ── 许可证常量 ──────────────────────────────────────────
TRIAL_DAYS = 0  # 已停用试用模式，改为 ¥29.9/年直接激活
_ACTIVATION_KEY: bytes = b"YTQJ2025_OFFICE_ASSISTANT_PRO"


# ══════════════════════════════════════════════════════════
# 辅助函数
# ══════════════════════════════════════════════════════════
def _generate_activation_code() -> str:
    """
    生成 16 位十六进制自校验激活码。

    格式：xxxxxxxx-xxxxxxxx（连字符仅显示用，存储无连字符）
    前 8 位 = 4 字节随机数据 hex，后 8 位 = HMAC‑SHA256(前8位, _ACTIVATION_KEY)[:8]
    """
    raw = os.urandom(4).hex()  # 8 hex chars = 4 bytes
    sig = hmac.new(_ACTIVATION_KEY, raw.encode("ascii"), hashlib.sha256).hexdigest()[:8]
    return raw + sig


def verify_activation_code(code: str) -> bool:
    """
    HMAC‑SHA256 验证 16 位激活码。

    参数
    ----
    code : str  连续 16 位十六进制字符（无连字符）

    返回
    ----
    bool  True 表示有效激活码
    """
    if not re.fullmatch(r"^[0-9a-fA-F]{16}$", code):
        return False
    code = code.lower()
    data = code[:8]
    expected_sig = hmac.new(_ACTIVATION_KEY, data.encode("ascii"), hashlib.sha256).hexdigest()[:8]
    return expected_sig == code[8:]


# ══════════════════════════════════════════════════════════
# LicenseManager
# ══════════════════════════════════════════════════════════
class LicenseManager:
    """许可证管理器 —— 试用／激活／卸载全生命周期管理。"""

    # ── 公开接口 ──────────────────────────────────────

    def __init__(self, parent) -> None:
        """
        参数
        ----
        parent : tk.Tk / 主窗口对象
            需要具备以下属性：
              - .root          (tk.Tk)
              - .status_bar    (tk.Label / ttk.Label)
              - .colors        (dict, 至少包含 'primary')

            也可直接传入 tk.Tk；此时状态栏功能静默降级。
        """
        self.parent = parent
        self.root: tk.Tk = getattr(parent, "root", parent)

        # 尝试从 parent 上获取颜色 / 状态栏引用
        self.colors: dict = getattr(parent, "colors", {"primary": "#4F46E5"})
        self._status_bar = getattr(parent, "status_bar", None)

        # 运行时缓存
        self._license_cache: Optional[dict] = None
        self._overlay_window: Optional[tk.Toplevel] = None

        # 确保 DATA_DIR 存在
        DATA_DIR.mkdir(parents=True, exist_ok=True)

    # ── 核心许可证检查 ────────────────────────────────

    def check_license(self) -> dict:
        """
        检查许可证状态。不再自动创建试用；无许可文件时返回"请激活"。

        返回 dict（始终包含 'status' 和 'detail'）：
          {'status': 'active',  'days_left': -1,   'detail': '已激活（¥29.9/年）'}
          {'status': 'expired', 'days_left': 0,    'detail': '请激活许可证'}
        """
        if not LICENSE_FILE.exists():
            return {"status": "expired", "days_left": 0, "detail": "请激活许可证"}

        data = self._read_license()
        if data is None:
            return {"status": "expired", "days_left": 0, "detail": "许可证文件损坏"}

        # 已激活 → 返回激活态
        if data.get("license_type") == "activated":
            return {"status": "active", "days_left": -1, "detail": "已激活（¥29.9/年）"}

        return {"status": "expired", "days_left": 0, "detail": "请激活许可证"}

    # ── 激活码验证 ────────────────────────────────────

    def verify_and_activate(self, code: str) -> bool:
        """
        验证激活码并写入 license.dat。

        参数
        ----
        code : str  16 位十六进制激活码（可含连字符，会被自动去除）

        返回
        ----
        bool  True 表示激活成功
        """
        # 允许带连字符的格式 xxxxxxxx-xxxxxxxx
        clean = code.replace("-", "").strip()
        if not verify_activation_code(clean):
            return False

        # 写入激活状态
        data = {
            "license_type": "activated",
            "activation_code": clean,
            "activated_at": datetime.now().isoformat(),
        }
        try:
            LICENSE_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
            self._license_cache = None
            return True
        except OSError:
            return False

    # ── 激活对话框 ────────────────────────────────────

    def activation_dialog(self, parent: Optional[tk.Toplevel] = None) -> bool:
        """
        弹出激活对话框（simpledialog）。

        参数
        ----
        parent : 可选，指定对话框的父窗口（默认 self.root）

        返回
        ----
        bool  True = 激活成功
        """
        p = parent if parent else self.root
        code = simpledialog.askstring(
            "激活许可证",
            "万能办公助手 ¥29.9/年\n\n请输入 16 位激活码（格式：xxxxxxxx-xxxxxxxx）：\n\n（如需购买请前往面包多）",
            parent=p,
        )
        if code is None:
            return False

        code = code.strip()
        if not code:
            messagebox.showwarning("输入错误", "激活码不能为空！", parent=self.root)
            return False

        if self.verify_and_activate(code):
            messagebox.showinfo("激活成功", "🎉 许可证已激活，感谢支持！", parent=self.root)
            # 更新状态栏
            self.update_status_bar_license()
            return True
        else:
            messagebox.showerror("激活失败", "❌ 激活码无效，请检查后重试。", parent=self.root)
            return False

    # ── 过期遮罩层 ────────────────────────────────────

    def expired_overlay(self) -> None:
        """
        在根窗口上创建一个半透明遮罩，提示用户激活。
        避免用户继续使用已过期的试用版本。
        """
        self._destroy_overlay()

        overlay = tk.Toplevel(self.root)
        overlay.title("许可证已过期")
        overlay.configure(bg="#1E1E1E")
        overlay.attributes("-topmost", True)
        # 拦截关闭按钮 → 强制退出（不让用户关掉遮罩继续用）
        overlay.protocol("WM_DELETE_WINDOW", lambda: self._on_overlay_exit(overlay))

        # 覆盖根窗口
        self.root.update_idletasks()
        x = self.root.winfo_x()
        y = self.root.winfo_y()
        w = max(self.root.winfo_width(), 800)
        h = max(self.root.winfo_height(), 600)
        overlay.geometry(f"{w}x{h}+{x}+{y}")
        overlay.resizable(False, False)
        overlay.grab_set()  # 确保遮罩是模态的，阻止点击下层窗口

        # 内容
        frame = tk.Frame(overlay, bg="#2D2D2D", bd=0)
        frame.place(relx=0.5, rely=0.45, anchor="center")

        tk.Label(frame, text="🔑 需激活许可证", font=("微软雅黑", 20, "bold"),
                 bg="#2D2D2D", fg="#FF9800").pack(pady=(30, 10))

        tk.Label(frame, text="¥29.9/年 · 请激活后使用", font=("微软雅黑", 11),
                 bg="#2D2D2D", fg="#CCCCCC").pack(pady=(0, 25))

        btn_frame = tk.Frame(frame, bg="#2D2D2D")
        btn_frame.pack(pady=(0, 30))

        activate_btn = tk.Button(
            btn_frame, text="🔑 激活许可证", font=("微软雅黑", 12, "bold"),
            bg=self.colors.get("primary", "#4F46E5"), fg="white",
            cursor="hand2", width=16, height=1,
            command=lambda: self._on_overlay_activate(overlay),
        )
        activate_btn.pack(side=tk.LEFT, padx=8)

        exit_btn = tk.Button(
            btn_frame, text="退出程序", font=("微软雅黑", 12),
            bg="#555555", fg="white", cursor="hand2", width=10, height=1,
            command=lambda: self._on_overlay_exit(overlay),
        )
        exit_btn.pack(side=tk.LEFT, padx=8)

        self._overlay_window = overlay

    def _on_overlay_activate(self, overlay: tk.Toplevel) -> None:
        """遮罩层『激活』按钮回调。"""
        if self.activation_dialog(parent=overlay):
            self._destroy_overlay()
            self.root.deiconify()
            self.root.lift()

    def _on_overlay_exit(self, overlay: tk.Toplevel) -> None:
        """遮罩层『退出』按钮回调。"""
        try:
            overlay.destroy()
        except Exception:
            pass
        try:
            self.root.destroy()
        except Exception:
            pass
        # 兜底强制退出
        try:
            os._exit(0)
        except Exception:
            pass

    def _destroy_overlay(self) -> None:
        if self._overlay_window is not None:
            try:
                self._overlay_window.destroy()
            except Exception:
                pass
            self._overlay_window = None

    # ── 状态栏许可证指示器 ────────────────────────────

    def update_status_bar_license(self) -> None:
        """更新状态栏上的许可证指示文字。"""
        if self._status_bar is None:
            return
        info = self.check_license()
        if info["status"] == "active":
            text = "✅ 已激活"
            fg = "#4CAF50"  # 绿
        else:
            text = "🔑 需激活"
            fg = "#FF9800"  # 橙
        try:
            self._status_bar.configure(text=text, fg=fg)
        except Exception:
            pass

    # ── 许可证警告 ────────────────────────────────────

    def check_and_show_license_warning(self) -> bool:
        """
        检查许可证状态。未激活则弹出遮罩阻塞界面。

        返回
        ----
        bool  True = 可以继续使用（已激活）
              False = 需要激活许可证
        """
        info = self.check_license()

        if info["status"] == "active":
            return True

        # 未激活 → 弹出遮罩
        if not LICENSE_FILE.exists():
            messagebox.showinfo(
                "欢迎使用",
                "🎉 万能办公助手 v6.2\n\n"
                "售价：¥29.9/年\n"
                "请激活后使用。",
                parent=self.root,
            )
        else:
            messagebox.showerror(
                "许可证无效",
                "⛔ 许可证无效或已损坏，请重新激活。",
                parent=self.root,
            )

        self.expired_overlay()
        return False

    # ── 欢迎对话框 ────────────────────────────────────

    def show_welcome_dialog(self, on_close: Optional[Callable] = None) -> None:
        """启动时显示欢迎/许可证信息对话框。"""
        info = self.check_license()

        dialog = tk.Toplevel(self.root)
        dialog.title("欢迎使用 万能办公助手")
        dialog.configure(bg="#F5F5F5")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        # 拦截关闭按钮 → 触发 on_close（进激活流程）
        if on_close:
            dialog.protocol("WM_DELETE_WINDOW", lambda: [dialog.destroy(), on_close()])

        # 窗口尺寸与居中
        win_w, win_h = 480, 380
        scr_w = dialog.winfo_screenwidth()
        scr_h = dialog.winfo_screenheight()
        x = (scr_w - win_w) // 2
        y = (scr_h - win_h) // 2
        dialog.geometry(f"{win_w}x{win_h}+{x}+{y}")

        # 标题
        tk.Label(dialog, text="✨ 万能办公助手 v6.2", font=("微软雅黑", 16, "bold"),
                 bg="#F5F5F5", fg="#333333").pack(pady=(30, 5))

        if info["status"] == "active":
            tk.Label(dialog, text="✅ 已激活 · 完整功能已解锁", font=("微软雅黑", 11),
                     bg="#F5F5F5", fg="#4CAF50").pack(pady=(5, 20))
        else:
            tk.Label(dialog, text="🔑 需激活 · ¥29.9/年", font=("微软雅黑", 11),
                     bg="#F5F5F5", fg="#FF9800").pack(pady=(5, 10))
            tk.Button(dialog, text="🔑 立即激活", font=("微软雅黑", 10),
                      bg=self.colors.get("primary", "#4F46E5"), fg="white",
                      cursor="hand2", command=lambda: self._welcome_activate(dialog),
                      ).pack(pady=(5, 20))

        # 功能介绍
        features_frame = tk.Frame(dialog, bg="#FFFFFF", bd=1, relief=tk.GROOVE)
        features_frame.pack(fill=tk.X, padx=30, pady=10)
        features = [
            "📋 剪贴板历史管理与回溯",
            "🔍 智能搜索与批量处理",
            "🛠 系统工具与托盘快捷键",
            "🔄 数据同步与备份",
        ]
        for feat in features:
            tk.Label(features_frame, text=feat, font=("微软雅黑", 9),
                     bg="#FFFFFF", fg="#555555", anchor="w").pack(fill=tk.X, padx=15, pady=3)

        btn = tk.Button(dialog, text="开始使用", font=("微软雅黑", 11, "bold"),
                        bg=self.colors.get("primary", "#4F46E5"), fg="white",
                        cursor="hand2", width=12,
                        command=lambda: [dialog.destroy(), on_close() if on_close else None])
        btn.pack(pady=(15, 25))

        dialog.wait_window()

    def _welcome_activate(self, dialog: tk.Toplevel) -> None:
        """欢迎对话框『激活』按钮。"""
        if self.activation_dialog():
            dialog.destroy()
            self.show_welcome_dialog()

    # ── 关于对话框 ────────────────────────────────────

    def about_dialog(self) -> None:
        """显示『关于』对话框。"""
        info = self.check_license()

        lines = [
            "万能办公助手 v6.2",
            "",
            "📌 智能办公 · 效率倍增",
            "帮助您在日常办公中提升效率",
            "",
            "━━━ 许可证 ━━━",
        ]
        if info["status"] == "active":
            lines.append("✅ 已激活（完整版 · ¥29.9/年）")
        else:
            lines.append("🔑 需激活 · ¥29.9/年")

        lines += [
            "",
            "━━━ 安装信息 ━━━",
            f"📂 安装路径: {Path(sys.executable).resolve().parent}",
            f"📦 大小: 约 153 MB",
            "",
            "💡 如需移动到 Program Files，请运行",
            "   安装万能办公助手.bat（管理员权限）",
        ]

        messagebox.showinfo(
            "关于 万能办公助手",
            "\n".join(lines),
            parent=self.root,
        )

    # ── 卸载流程（带安全防护） ──────────────────────

    def uninstall_dialog(self) -> None:
        """
        卸载确认与执行对话框。

        【安全防护】
          1) 仅在打包 (frozen) 的 exe 中运行
          2) install_dir 不得是桌面路径或项目路径
        """
        # ── 安全防护 1：仅打包 exe ──
        if not getattr(sys, "frozen", False):
            messagebox.showwarning(
                "卸载失败",
                "请在打包后的 exe 中运行卸载。\n\n"
                "当前为源代码/开发环境运行，请使用打包后的应用程序执行卸载。",
                parent=self.root,
            )
            return

        # ── 确定安装目录 ──
        install_dir = Path(sys.executable).resolve().parent

        # ── 安全防护 2：非桌面/项目路径 ──
        desktop = Path.home() / "Desktop"
        project_root = Path(r"C:\Users\Administrator\Desktop\OfficeAssistant_v6.2")

        if install_dir == desktop:
            messagebox.showwarning(
                "卸载失败",
                "检测到安装目录为「桌面路径」，请从正确的程序安装位置运行卸载。",
                parent=self.root,
            )
            return

        if install_dir == project_root:
            messagebox.showwarning(
                "卸载失败",
                "检测到安装目录为「项目源码目录」，请从正确的程序安装位置运行卸载。",
                parent=self.root,
            )
            return

        # 防止 install_dir 是桌面子目录或项目子目录
        try:
            if desktop in install_dir.parents:
                messagebox.showwarning(
                    "卸载失败",
                    "安装目录位于桌面路径下，请确认正确的程序安装位置。",
                    parent=self.root,
                )
                return
        except ValueError:
            pass

        try:
            if project_root in install_dir.parents:
                messagebox.showwarning(
                    "卸载失败",
                    "安装目录位于项目源码目录下，请确认正确的程序安装位置。",
                    parent=self.root,
                )
                return
        except ValueError:
            pass

        # ── 展示确认对话框 ──
        confirm = messagebox.askyesno(
            "确认卸载",
            f"⚠️ 确定要卸载「万能办公助手 v6.2」吗？\n\n"
            f"安装路径：{install_dir}\n"
            f"将删除：\n"
            f"  · 程序目录（{install_dir.name}/）\n"
            f"  · 许可证与配置 (%APPDATA%/OfficeAssistant_v6.2/)\n\n"
            f"此操作不可撤销！",
            parent=self.root,
        )
        if not confirm:
            return

        # ── 执行卸载 ──
        self.do_uninstall(install_dir)

    def do_uninstall(self, install_dir: Optional[Path] = None) -> bool:
        """
        实际执行卸载操作。

        参数
        ----
        install_dir : Path | None
            程序安装目录。若为 None，自动从 sys.executable 推断。

        返回
        ----
        bool  True = 卸载成功
        """
        if install_dir is None:
            if getattr(sys, "frozen", False):
                install_dir = Path(sys.executable).resolve().parent
            else:
                messagebox.showwarning(
                    "卸载失败",
                    "请在打包后的 exe 中运行卸载。",
                    parent=self.root,
                )
                return False

        install_dir = Path(install_dir).resolve()

        # ── 二次安全防护 ──
        if not getattr(sys, "frozen", False):
            messagebox.showwarning(
                "卸载失败",
                "请在打包后的 exe 中运行卸载。",
                parent=self.root,
            )
            return False

        desktop = Path.home() / "Desktop"
        project_root = Path(r"C:\Users\Administrator\Desktop\OfficeAssistant_v6.2").resolve()

        if install_dir == desktop:
            messagebox.showwarning("卸载失败", "安装目录为桌面路径，拒绝卸载。", parent=self.root)
            return False
        if install_dir == project_root:
            messagebox.showwarning("卸载失败", "安装目录为项目源码目录，拒绝卸载。", parent=self.root)
            return False

        # ── 执行删除 ──
        errors: list[str] = []

        # 1) 删除 APPDATA 下的许可证和配置
        try:
            if DATA_DIR.exists():
                shutil.rmtree(DATA_DIR, ignore_errors=True)
        except Exception as e:
            errors.append(f"配置目录删除失败: {e}")

        # 2) 删除程序目录
        try:
            if install_dir.exists():
                shutil.rmtree(install_dir, ignore_errors=True)
        except Exception as e:
            errors.append(f"程序目录删除失败: {e}")

        # 3) 尝试删除开始菜单快捷方式 (Windows)
        try:
            start_menu = Path(os.environ.get("APPDATA", "")) / "Microsoft" / "Windows" / "Start Menu" / "Programs"
            for lnk in start_menu.rglob("*万能办公助手*.lnk"):
                lnk.unlink(missing_ok=True)
            for lnk in start_menu.rglob("*OfficeAssistant*.lnk"):
                lnk.unlink(missing_ok=True)
        except Exception:
            pass

        # 4) 尝试删除桌面快捷方式
        try:
            for lnk in Path.home().glob("Desktop/*万能办公助手*.lnk"):
                lnk.unlink(missing_ok=True)
            for lnk in Path.home().glob("Desktop/*OfficeAssistant*.lnk"):
                lnk.unlink(missing_ok=True)
        except Exception:
            pass

        if errors:
            messagebox.showwarning(
                "卸载报告",
                "部分内容未能完全删除：\n" + "\n".join(errors),
                parent=self.root,
            )
            return False

        messagebox.showinfo("卸载完成", "✅ 万能办公助手 v6.2 已成功卸载。\n感谢您的使用！", parent=self.root)

        # ── 退出程序 ──
        try:
            self.root.destroy()
        except Exception:
            pass

        return True

    # ══════════════════════════════════════════════════
    # 内部方法
    # ══════════════════════════════════════════════════

    def _create_trial(self) -> None:
        """创建 7 天试用许可证文件。"""
        data = {
            "license_type": "trial",
            "first_run": datetime.now().isoformat(),
            "created_at": datetime.now().isoformat(),
        }
        try:
            DATA_DIR.mkdir(parents=True, exist_ok=True)
            LICENSE_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
            self._license_cache = None
        except OSError:
            pass  # 静默失败，check_license 会兜底

    def _read_license(self) -> Optional[dict]:
        """读取 license.dat，返回 dict 或 None。"""
        if self._license_cache is not None:
            return self._license_cache

        try:
            if not LICENSE_FILE.exists():
                return None
            raw = LICENSE_FILE.read_text(encoding="utf-8").strip()
            if not raw:
                return None
            data = json.loads(raw)
            self._license_cache = data
            return data
        except (json.JSONDecodeError, OSError, ValueError):
            return None


# ══════════════════════════════════════════════════════════
# 模块级便捷入口
# ══════════════════════════════════════════════════════════
def generate_activation_code() -> str:
    """生成一个 16 位十六进制自校验激活码（含连字符方便阅读）。"""
    raw = _generate_activation_code()
    return f"{raw[:8]}-{raw[8:]}"


def batch_generate_codes(count: int = 10) -> list[str]:
    """批量生成激活码。"""
    return [generate_activation_code() for _ in range(count)]


# ── 自测 ──
if __name__ == "__main__":
    # 验证码生成与验证自检
    print("=" * 50)
    print("  LicenseManager — 自检")
    print("=" * 50)

    # 生成 & 验证
    for i in range(5):
        code = generate_activation_code()
        clean = code.replace("-", "")
        valid = verify_activation_code(clean)
        print(f"  [{i+1}] {code}  →  {'✅ 有效' if valid else '❌ 无效'}")

    # 无效码测试
    print()
    print("  无效码测试：")
    tests = [
        ("12345678-12345678", False),   # 随机，非 HMAC 校验
        ("abcd1234",         False),    # 长度不足
        ("zzzzzzzz-xxxxxxxx", False),   # 非法 hex
    ]
    for code, expected in tests:
        clean = code.replace("-", "").strip() if "-" in code else code
        result = verify_activation_code(clean)
        status = "✅" if result == expected else "❌"
        print(f"  {status} verify({code}) = {result} (期望={expected})")

    print()
    print("  DATA_DIR:", DATA_DIR)
    print("  LICENSE_FILE:", LICENSE_FILE)
    print()
    print("  自检完成。")
