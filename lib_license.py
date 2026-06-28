# -*- coding: utf-8 -*-
"""
万能办公助手 - 许可证管理模块 (lib_license.py)
试用 · 激活 · 到期遮罩 · 状态栏 · 关于对话框
"""
import sys, os, json, time, hashlib, threading, urllib
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from utils import DATA_DIR, safe_str, _log_error


class LicenseManager:
    """许可证管理器 — 封装所有授权相关逻辑"""

    LICENSE_FILE = DATA_DIR / "license.dat"
    USED_CODES_FILE = DATA_DIR / "used_codes.json"
    TRIAL_DAYS = 7
    _ACTIVATION_KEY = b"YTQJ2025_OFFICE_ASSISTANT_PRO"

    def __init__(self, parent):
        """parent: OfficeAssistant 实例"""
        self.parent = parent
        self.root = parent.root
        self.colors = parent.colors
        self.license_label = parent.license_label

    # ===================== 许可证检查 =====================

    def check_license(self):
        """检查许可证状态，返回 (是否有效, 剩余天数, 错误信息)"""
        try:
            if not self.LICENSE_FILE.exists():
                install_time = time.time()
                install_date = datetime.now().strftime("%Y-%m-%d")
                import uuid
                machine_id = hashlib.md5(
                    (str(uuid.getnode()) + str(uuid.getnode()) + "OfficeAssistant_v5").encode()
                ).hexdigest()[:16].upper()
                data = json.dumps({
                    "install_date": install_date,
                    "install_timestamp": install_time,
                    "expire_timestamp": install_time + self.TRIAL_DAYS * 86400,
                    "machine_id": machine_id,
                    "activated": False,
                    "activation_code": ""
                }, ensure_ascii=False)
                self.LICENSE_FILE.write_text(data, encoding="utf-8")
                return True, self.TRIAL_DAYS, ""
            data = json.loads(self.LICENSE_FILE.read_text(encoding="utf-8"))
            expire_ts = data.get("expire_timestamp", 0)
            remaining = int((expire_ts - time.time()) / 86400)
            activated = data.get("activated", False)
            if activated:
                return True, remaining, ""
            if remaining < 0:
                return False, remaining, "授权已过期"
            if remaining == 0:
                return False, 0, "今天到期，请尽快续费"
            return True, remaining, ""
        except Exception as e:
            return True, self.TRIAL_DAYS, f"许可证检查异常({e})，视为有效"

    def verify_activation_code(self, code: str) -> dict:
        """验证HMAC通用激活码"""
        raw = code.strip().upper().replace("-", "").replace(" ", "")
        if not raw or len(raw) != 16:
            return {"valid": False, "msg": "格式错误，请输入16位激活码"}
        used = self._load_used_codes()
        if raw in used:
            return {"valid": False, "msg": "该激活码已被使用"}
        try:
            import hmac as _hm
            for cid in range(1, 2001):
                expected = _hm.new(self._ACTIVATION_KEY, str(cid).encode(),
                                   hashlib.sha256).hexdigest()[:16].upper()
                if raw == expected:
                    return {"valid": True, "msg": f"激活码有效 (序号{cid})", "code": raw}
            return {"valid": False, "msg": "激活码无效"}
        except Exception as e:
            return {"valid": False, "msg": f"验证异常: {e}"}

    def _load_used_codes(self):
        try:
            if self.USED_CODES_FILE.exists():
                return set(json.loads(self.USED_CODES_FILE.read_text(encoding="utf-8")))
        except Exception:
            pass
        return set()

    def _save_used_code(self, code):
        used = self._load_used_codes()
        used.add(code)
        self.USED_CODES_FILE.write_text(
            json.dumps(sorted(used), ensure_ascii=False, indent=2), encoding="utf-8")

    # ===================== UI 对话框 =====================

    def activation_dialog(self, error_msg=""):
        """激活码输入对话框"""
        win = tk.Toplevel(self.root)
        win.title("激活万能办公助手")
        win.geometry("520x400")
        win.transient(self.root)
        win.grab_set()
        win.configure(bg=self.colors['light'])
        tk.Label(win, text="🔐 激活授权", font=("微软雅黑", 18, "bold"),
                bg=self.colors['light'], fg=self.colors['dark']).pack(pady=15)
        tk.Label(win, text="永久授权 · 终身免费更新 · 专属客服",
                font=("微软雅黑", 10), fg=self.colors['gray'], bg=self.colors['light']).pack()
        tk.Label(win, text="输入您的激活码:", font=("微软雅黑", 11),
                bg=self.colors['light']).pack(pady=(20,5))
        code_var = tk.StringVar()
        code_entry = tk.Entry(win, textvariable=code_var, font=("Consolas", 14),
                             width=22, justify="center")
        code_entry.pack(pady=10)
        tk.Label(win, text="格式: XXXX-XXXX-XXXX-XXXX 或 16位激活码",
                font=("微软雅黑", 9), fg=self.colors['gray'], bg=self.colors['light']).pack()
        if error_msg:
            tk.Label(win, text=f"❌ {error_msg}", fg="red",
                    font=("微软雅黑", 10), bg=self.colors['light']).pack()
        tk.Label(win, text="提示: 购买后联系客服获取激活码",
                font=("微软雅黑", 9), fg=self.colors['gray'], bg=self.colors['light']).pack(pady=5)
        btn_row = tk.Frame(win, bg=self.colors['light'])
        btn_row.pack(pady=15)

        def _activate():
            code = code_var.get().strip()
            result = self.verify_activation_code(code)
            if result["valid"]:
                try:
                    lic = json.loads(self.LICENSE_FILE.read_text(encoding="utf-8"))
                    lic["activated"] = True
                    lic["activation_code"] = result["code"]
                    lic["expire_timestamp"] = time.time() + 365 * 86400
                    self.LICENSE_FILE.write_text(json.dumps(lic, ensure_ascii=False, indent=2), encoding="utf-8")
                    self._save_used_code(result["code"])
                    win.destroy()
                    messagebox.showinfo("✅ 激活成功", "🎉 激活成功！有效期1年\n\n感谢您的支持！")
                except Exception as e:
                    messagebox.showerror("错误", f"激活失败: {e}")
            else:
                for w in btn_row.winfo_children():
                    w.destroy()
                tk.Label(btn_row, text=f"❌ {result['msg']}", fg="red",
                        font=("微软雅黑", 11), bg=self.colors['light']).pack()
                tk.Button(btn_row, text="重试", command=lambda: self.activation_dialog(""),
                         cursor="hand2", bg=self.colors['primary'], fg="white",
                         font=("微软雅黑", 10), width=10).pack(pady=10)

        tk.Button(btn_row, text="🚀 立即激活", command=_activate, cursor="hand2",
                 bg=self.colors['primary'], fg="white", font=("微软雅黑", 12, "bold"),
                 width=16).pack(pady=5)
        valid, remaining, _ = self.check_license()
        tk.Button(btn_row, text=f"试用 (剩余{remaining}天)", command=win.destroy, cursor="hand2",
                 font=("微软雅黑", 10), width=16).pack(pady=5)
        win.bind("<Return>", lambda e: _activate())
        code_entry.focus()

    def expired_overlay(self, remaining):
        """许可证到期全屏遮罩"""
        overlay = tk.Frame(self.root, bg="#1E293B")
        overlay.place(relx=0, rely=0, relwidth=1, relheight=1)
        card = tk.Frame(overlay, bg="white", padx=40, pady=30)
        card.place(relx=0.5, rely=0.5, anchor="center")
        tk.Label(card, text="⏰ 授权已到期",
                font=("微软雅黑", 22, "bold"), fg="#EF4444", bg="white").pack(pady=(0,10))
        tk.Label(card, text="您的万能办公助手已过期",
                font=("微软雅黑", 12), fg="#374151", bg="white").pack()
        tk.Label(card, text=f"到期时间: {datetime.now().strftime('%Y-%m-%d')}",
                font=("微软雅黑", 10), fg="#9CA3AF", bg="white").pack(pady=5)
        tk.Label(card, text="续费价格: ¥99/年 (首年优惠)",
                font=("微软雅黑", 12, "bold"), fg="#10B981", bg="white").pack(pady=15)
        tk.Label(card, text="✅ 全部功能 · ✅ 永久更新 · ✅ 技术支持",
                font=("微软雅黑", 10), fg="#64748B", bg="white").pack()
        tk.Label(card, text='购买激活码: 面包多 (mbd.pub) 搜索「万能办公助手」',
                font=("微软雅黑", 9), fg="#9CA3AF", bg="white").pack(pady=10)
        tk.Button(card, text="💳 续费激活",
                 command=lambda: [overlay.destroy(), self.activation_dialog("")],
                 cursor="hand2", bg=self.colors['primary'], fg="white",
                 font=("微软雅黑", 13, "bold"), width=18, pady=8).pack(pady=15)

        def on_closing():
            pass
        self.root.protocol("WM_DELETE_WINDOW", on_closing)

    def update_status_bar_license(self):
        """更新状态栏的许可证信息"""
        valid, remaining, err = self.check_license()
        if not self.license_label:
            return
        if err and remaining < 0:
            self.root.after(500, self.expired_overlay, remaining)
        elif remaining <= 30 and remaining > 0:
            self.license_label.config(text=f"⚠️ 剩余 {remaining} 天", fg="#F59E0B")
        elif remaining <= 7 and remaining > 0:
            self.license_label.config(text=f"🔥 仅剩 {remaining} 天!", fg="#EF4444")
        else:
            self.license_label.config(text=f"✅ {remaining}天", fg="#10B981")

    def check_and_show_license_warning(self):
        """启动时检查授权状态"""
        valid, remaining, err = self.check_license()
        if not valid:
            self.expired_overlay(remaining)
        elif remaining <= 30:
            if self.license_label:
                self.license_label.config(text=f"剩余 {remaining} 天", fg="#F59E0B")
            messagebox.showwarning("授权即将到期",
                f"您的万能办公助手将在 {remaining} 天后到期\n\n请及时续费以免影响使用\n\n续费请联系微信: 在线客服")
        else:
            if self.license_label:
                self.license_label.config(text=f"已授权 {remaining}天", fg="#10B981")

    def show_welcome_dialog(self):
        """首次启动欢迎对话框"""
        win = tk.Toplevel(self.root)
        win.title("欢迎使用万能办公助手")
        win.geometry("520x440")
        win.transient(self.root)
        win.grab_set()
        win.configure(bg='#FFFFFF')
        tk.Frame(win, bg='#4F46E5', height=4).pack(fill=tk.X)
        tk.Label(win, text="🧰", font=("微软雅黑", 36), bg='#FFFFFF').pack(pady=(20, 5))
        tk.Label(win, text="欢迎使用万能办公助手！", font=("微软雅黑", 18, "bold"),
                fg='#0F172A', bg='#FFFFFF').pack()
        tk.Label(win, text="v6.1 商业版", font=("微软雅黑", 10),
                fg='#64748B', bg='#FFFFFF').pack()
        tk.Frame(win, bg='#E2E8F0', height=1).pack(fill=tk.X, padx=40, pady=15)
        tk.Label(win, text="📄 格式互转  📊 Excel  📝 文档  📄 PDF\n🖼 图片  ⚡ 快捷  📋 剪贴板  🔍 搜索  📅 日程",
                font=("微软雅黑", 10), fg='#475569', bg='#FFFFFF', justify=tk.CENTER).pack(padx=30)
        tk.Frame(win, bg='#E2E8F0', height=1).pack(fill=tk.X, padx=40, pady=15)
        tk.Label(win, text="您有 7 天免费试用期，到期后需激活方可继续使用",
                font=("微软雅黑", 10), fg='#D97706', bg='#FFFFFF').pack()
        tk.Label(win, text="激活码 ¥99/年 · 面包多 (mbd.pub) 购买",
                font=("微软雅黑", 9), fg='#94A3B8', bg='#FFFFFF').pack(pady=(4, 12))
        btn_f = tk.Frame(win, bg='#FFFFFF')
        btn_f.pack(pady=8)
        tk.Button(btn_f, text="🔓 免费试用 7 天",
                 command=lambda: [win.destroy(), self.root.after(500, self.check_and_show_license_warning)],
                 cursor="hand2", bg='#E2E8F0', fg='#0F172A',
                 font=("微软雅黑", 11), width=18, bd=0, pady=6).pack(side=tk.LEFT, padx=8)
        tk.Button(btn_f, text="📥 输入激活码",
                 command=lambda: [win.destroy(), self.root.after(200, lambda: self.activation_dialog(""))],
                 cursor="hand2", bg='#4F46E5', fg='white',
                 font=("微软雅黑", 11, "bold"), width=18, bd=0, pady=6).pack(side=tk.LEFT, padx=8)

    def check_for_update(self):
        """检查更新（后台线程）"""
        def _check():
            try:
                import urllib.request
                url = "https://raw.githubusercontent.com/your-username/office-assistant/main/version.json"
                req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
                try:
                    with urllib.request.urlopen(req, timeout=5) as resp:
                        data = json.loads(resp.read().decode('utf-8'))
                    remote = data.get('version', '')
                    if remote and remote > "6.1":
                        self.root.after(0, lambda: messagebox.showinfo("🔄 有更新",
                            f"新版本 {remote} 已发布！\n{data.get('note', '')}\n\n请前往面包多购买最新版。"))
                    else:
                        self.root.after(0, lambda: messagebox.showinfo("✅ 已最新",
                            "您正在使用最新版本 v6.1"))
                except Exception:
                    self.root.after(0, lambda: messagebox.showinfo("检查更新",
                        "无法连接更新服务器\n\n请前往面包多 (mbd.pub) 查看最新版本"))
            except Exception:
                self.root.after(0, lambda: messagebox.showinfo("检查更新",
                    "无法连接更新服务器\n\n请前往面包多 (mbd.pub) 查看最新版本"))
        threading.Thread(target=_check, daemon=True).start()

    def about_dialog(self):
        """专业关于对话框"""
        win = tk.Toplevel(self.root)
        win.title("关于 万能办公助手")
        win.geometry("460x420")
        win.transient(self.root)
        win.grab_set()
        win.configure(bg='#FFFFFF')
        tk.Frame(win, bg='#4F46E5', height=4).pack(fill=tk.X)
        tk.Label(win, text="🧰", font=("微软雅黑", 40), bg='#FFFFFF').pack(pady=(20, 5))
        tk.Label(win, text="万能办公助手", font=("微软雅黑", 20, "bold"),
                fg='#0F172A', bg='#FFFFFF').pack()
        tk.Label(win, text="v6.1 商业版", font=("微软雅黑", 11),
                fg='#64748B', bg='#FFFFFF').pack()
        tk.Frame(win, bg='#E2E8F0', height=1).pack(fill=tk.X, padx=40, pady=15)
        info_frame = tk.Frame(win, bg='#FFFFFF')
        info_frame.pack(padx=40)
        items = [
            ("📋", "功能", "格式互转 · PDF工具 · 文档处理 · 图片处理 · 更多"),
            ("💳", "价格", "¥99/年（一年授权，到期续费）"),
            ("🛒", "购买", "面包多 (mbd.pub) 搜索「万能办公助手」"),
            ("🔧", "技术", "Python + tkinter · PyInstaller 单文件封装"),
        ]
        for icon, label, value in items:
            row = tk.Frame(info_frame, bg='#FFFFFF')
            row.pack(fill=tk.X, pady=4)
            tk.Label(row, text=f"{icon} {label}:", font=("微软雅黑", 10, "bold"),
                    fg='#0F172A', bg='#FFFFFF', width=6, anchor='w').pack(side=tk.LEFT)
            tk.Label(row, text=value, font=("微软雅黑", 10),
                    fg='#475569', bg='#FFFFFF', anchor='w').pack(side=tk.LEFT, padx=(8, 0))
        tk.Frame(win, bg='#E2E8F0', height=1).pack(fill=tk.X, padx=40, pady=15)
        tk.Label(win, text="© 2026 万能办公助手 保留所有权利",
                font=("微软雅黑", 9), fg='#94A3B8', bg='#FFFFFF').pack()
        btn_f = tk.Frame(win, bg='#FFFFFF')
        btn_f.pack(pady=(10, 20))
        tk.Button(btn_f, text="🔄 检查更新", command=self.check_for_update, cursor="hand2",
                 bg='#E2E8F0', fg='#0F172A', font=("微软雅黑", 10),
                 width=14, bd=0).pack(side=tk.LEFT, padx=6)
        tk.Button(btn_f, text="关闭", command=win.destroy, cursor="hand2",
                 bg='#4F46E5', fg='white', font=("微软雅黑", 10),
                 width=12, bd=0).pack(side=tk.LEFT, padx=6)
