# -*- coding: utf-8 -*-
"""
短信群发系统 —— v2 科技感界面（发送 + 接收监听）
- 开启 Per-Monitor DPI 感知，解决高DPI屏幕模糊问题
- 霓虹青/电紫配色 + 圆角发光组件
- 实时统计仪表盘 + 呼吸状态灯 + 自绘进度条
- 发送：Excel 名单批量发送，自动重试，导出成功/失败 CSV
- 接收：轮询手机网关收件箱，去重后追加到 回复记录.csv
"""
import ctypes

# ===== 必须在创建任何 Tk 窗口之前开启 DPI 感知 =====
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

import sys
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import requests
import pandas as pd
import datetime
import threading


def app_dir():
    """兼容 PyInstaller onefile：返回 exe 所在目录。"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


# ===================== 业务配置 =====================
DEFAULT_IP = "255.255.255.0"
DEFAULT_PORT = "8080"
DEFAULT_USER = "username"
DEFAULT_PASS = "12345678"
WAIT_BETWEEN = 3
RETRY_TIMES = 3
RETRY_WAIT = 3
POLL_INTERVAL = 5
HTTP_TIMEOUT = 5

MSG_TPL = ("【电子协会】恭喜{major}专业的{name}同学"
           "被协会硬件设计部正式录取,学号{sid},期待你的到来!")

# ===================== 配色 =====================
BG      = "#05070f"
PANEL   = "#0a1020"
CARD    = "#0d1526"
CARD_HI = "#141f38"
BORDER  = "#1e2c4e"
ACCENT  = "#00e5ff"
ACCENT2 = "#8b5cff"
SUCCESS = "#22e5a0"
DANGER  = "#ff4d6d"
WARN    = "#ffb547"
TEXT    = "#e8f1ff"
MUTED   = "#62749a"

# ===================== 字体 =====================
F_UI    = ("Microsoft YaHei", 10)
F_UI_B  = ("Microsoft YaHei", 10, "bold")
F_H1    = ("Microsoft YaHei", 15, "bold")
F_LABEL = ("Microsoft YaHei", 9)
F_STAT_N= ("Segoe UI", 22, "bold")
F_STAT_L= ("Microsoft YaHei", 8)
F_LOG2  = ("Consolas", 9)


# ===================== 工具函数 =====================
def round_rect(cv, x1, y1, x2, y2, r, **kw):
    pts = [
        x1 + r, y1, x2 - r, y1,
        x2, y1, x2, y1 + r,
        x2, y2 - r, x2, y2,
        x2 - r, y2, x1 + r, y2,
        x1, y2, x1, y2 - r,
        x1, y1 + r, x1, y1,
    ]
    return cv.create_polygon(pts, smooth=True, **kw)


def lighten(hex_color, amount=30):
    h = hex_color.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    r = min(255, r + amount)
    g = min(255, g + amount)
    b = min(255, b + amount)
    return f"#{r:02x}{g:02x}{b:02x}"


# ===================== 组件 =====================
class Card(tk.Frame):
    def __init__(self, parent, title=None, accent=ACCENT, padding=16, **kw):
        super().__init__(parent, bg=CARD, highlightthickness=1,
                         highlightbackground=BORDER, bd=0, **kw)
        self.body = None
        if title:
            head = tk.Frame(self, bg=CARD)
            head.pack(fill="x", padx=padding, pady=(12, 6))
            tk.Frame(head, bg=accent, width=3, height=14).pack(side="left", padx=(0, 9))
            tk.Label(head, text=title, bg=CARD, fg=TEXT,
                     font=F_UI_B).pack(side="left")
            self.body = tk.Frame(self, bg=CARD)
            self.body.pack(fill="both", expand=True, padx=padding, pady=(0, 12))
        else:
            self.body = tk.Frame(self, bg=CARD)
            self.body.pack(fill="both", expand=True, padx=padding, pady=padding)


class NeonButton(tk.Canvas):
    def __init__(self, parent, text, command, color=ACCENT,
                 w=200, h=44, fg="#04121a", font=None):
        super().__init__(parent, width=w, height=h, bg=BG,
                         highlightthickness=0, bd=0)
        self.command = command
        self.color = color
        self.fg = fg
        self.text = text
        self.w, self.h = w, h
        self.font = font or F_UI_B
        self.hover = False
        self.enabled = True
        self._draw()
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<Button-1>", self._on_click)

    def _draw(self):
        self.delete("all")
        c = self.color if self.enabled else "#3a4a6e"
        if self.hover and self.enabled:
            c = lighten(c, 25)
        round_rect(self, 1.5, 1.5, self.w - 1.5, self.h - 1.5, 10,
                   fill=c, outline="")
        self.create_rectangle(14, 6, self.w - 14, 7,
                              fill="#ffffff", outline="")
        self.create_text(self.w / 2, self.h / 2 + 1, text=self.text,
                         fill=self.fg if self.enabled else "#9aa8c7",
                         font=self.font)

    def _on_enter(self, e):
        self.hover = True
        self._draw()

    def _on_leave(self, e):
        self.hover = False
        self._draw()

    def _on_click(self, e):
        if self.enabled and self.command:
            self.command()

    def set_text(self, t):
        self.text = t
        self._draw()

    def set_color(self, c):
        self.color = c
        self._draw()


class ProgressBar(tk.Canvas):
    def __init__(self, parent, w=320, h=10, bg=CARD):
        super().__init__(parent, width=w, height=h, bg=bg,
                         highlightthickness=0, bd=0)
        self.w, self.h = w, h
        self.value = 0
        self._draw()

    def _draw(self):
        self.delete("all")
        round_rect(self, 0, 0, self.w, self.h, 5, fill=PANEL, outline="")
        fw = max(6, int(self.w * self.value / 100))
        if fw > 6:
            round_rect(self, 0, 0, fw, self.h, 5, fill=ACCENT, outline="")
            self.create_oval(fw - 6, 1, fw + 4, self.h - 1,
                             fill="#aef6ff", outline="")

    def set(self, v):
        self.value = max(0.0, min(100.0, v))
        self._draw()


class StatusDot(tk.Canvas):
    def __init__(self, parent, color=MUTED, size=12):
        super().__init__(parent, width=size, height=size, bg=BG,
                         highlightthickness=0, bd=0)
        self.size = size
        self.color = color
        self._on = True
        self._draw()

    def _draw(self):
        self.delete("all")
        s = self.size
        self.create_oval(0, 0, s, s, fill="", outline=self.color, width=1)
        r = 2 if self._on else 3
        self.create_oval(r, r, s - r, s - r, fill=self.color, outline="")

    def set_color(self, c):
        self.color = c
        self._draw()

    def blink(self, on):
        self._on = on
        self._draw()


# ===================== 主应用 =====================
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("SMS Gateway Pro")
        W, H = 1160, 1000
        self.root.minsize(1020, 880)
        self.root.configure(bg=BG)
        x = (self.root.winfo_screenwidth() - W) // 2
        y = (self.root.winfo_screenheight() - H) // 2
        self.root.geometry(f"{W}x{H}+{x}+{y}")

        self.excel_path = tk.StringVar()
        self.gw_ip = tk.StringVar(value=DEFAULT_IP)
        self.gw_port = tk.StringVar(value=DEFAULT_PORT)
        self.gw_user = tk.StringVar(value=DEFAULT_USER)
        self.gw_pass = tk.StringVar(value=DEFAULT_PASS)

        self.sending = False
        self._stop_evt = threading.Event()
        self.ok_n = self.fail_n = self.total_n = 0

        self.listening = False
        self._recv_evt = threading.Event()
        self._listen_gen = 0
        self.reply_n = 0
        self.seen_ids = set()
        self.recv_csv = os.path.join(app_dir(), "回复记录.csv")

        self._setup_style()
        self._build_ui()
        self._load_seen_async()

    # ---------- ttk 深色主题 ----------
    def _setup_style(self):
        s = ttk.Style()
        try:
            s.theme_use("clam")
        except Exception:
            pass
        s.configure("TNotebook", background=CARD, borderwidth=0,
                   tabmargins=(8, 6, 8, 0))
        s.configure("TNotebook.Tab", background=PANEL, foreground=MUTED,
                    padding=(14, 7), font=F_UI_B, borderwidth=0)
        s.map("TNotebook.Tab",
              background=[("selected", CARD_HI)],
              foreground=[("selected", ACCENT)])

    # ---------- UI 构建 ----------
    def _build_ui(self):
        tk.Frame(self.root, bg=ACCENT, height=2).pack(fill="x")

        bar = tk.Frame(self.root, bg=BG, height=56)
        bar.pack(fill="x")
        bar.pack_propagate(False)
        inner = tk.Frame(bar, bg=BG)
        inner.pack(fill="both", expand=True, padx=22)
        logo = tk.Canvas(inner, width=30, height=30, bg=BG,
                         highlightthickness=0)
        logo.pack(side="left", padx=(0, 12))
        round_rect(logo, 0, 0, 30, 30, 7, fill=ACCENT, outline="")
        logo.create_text(15, 16, text="⚡", fill="#04121a",
                         font=("Segoe UI", 13, "bold"))
        tk.Label(inner, text="SMS GATEWAY PRO", bg=BG, fg=TEXT,
                 font=F_H1).pack(side="left")
        tk.Label(inner, text="  短信群发控制台", bg=BG, fg=MUTED,
                 font=F_LABEL).pack(side="left", padx=(6, 0), pady=(6, 0))

        right = tk.Frame(inner, bg=BG)
        right.pack(side="right")
        self.dot = StatusDot(right, color=MUTED)
        self.dot.pack(side="left", padx=(0, 8))
        self.dot_label = tk.Label(right, text="就绪", bg=BG, fg=MUTED,
                                  font=F_LABEL)
        self.dot_label.pack(side="left", padx=(0, 16))
        tk.Label(right, text="v2.1", bg=BG, fg=MUTED,
                 font=F_LABEL).pack(side="left")

        body = tk.Frame(self.root, bg=BG)
        body.pack(fill="both", expand=True, padx=20, pady=14)
        body.columnconfigure(0, weight=0, minsize=370)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)

        # ===== 左列 =====
        left = tk.Frame(body, bg=BG)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 14))

        c1 = Card(left, "网关配置", accent=ACCENT)
        c1.pack(fill="x", pady=(0, 12))
        self._build_gw(c1.body)

        c2 = Card(left, "名单文件", accent=ACCENT2)
        c2.pack(fill="x", pady=(0, 12))
        self._build_file(c2.body)

        c3 = Card(left, "发送控制", accent=SUCCESS)
        c3.pack(fill="x", pady=(0, 12))
        self._build_action(c3.body)

        c4 = Card(left, "接收监听", accent=ACCENT2)
        c4.pack(fill="both", expand=True)
        self._build_recv(c4.body)

        # ===== 右列 =====
        right_col = tk.Frame(body, bg=BG)
        right_col.grid(row=0, column=1, sticky="nsew")
        right_col.rowconfigure(1, weight=1)
        right_col.columnconfigure(0, weight=1)

        c5 = Card(right_col, "实时统计", accent=ACCENT)
        c5.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        self._build_stats(c5.body)

        c6 = Card(right_col, "日志 / 收件箱", accent=ACCENT2)
        c6.grid(row=1, column=0, sticky="nsew")
        self._build_notebook(c6.body)

        foot = tk.Frame(self.root, bg=PANEL, height=26)
        foot.pack(fill="x", side="bottom")
        foot.pack_propagate(False)
        self.status = tk.Label(foot, text="● 就绪", bg=PANEL, fg=MUTED,
                               font=F_LABEL, anchor="w")
        self.status.pack(side="left", padx=16)
        self.hint = tk.Label(foot, text="手机网关 · HTTP API", bg=PANEL,
                             fg=MUTED, font=F_LABEL)
        self.hint.pack(side="right", padx=16)

    def _input(self, parent, label, var, show=None):
        row = tk.Frame(parent, bg=CARD)
        row.pack(fill="x", pady=(0, 10))
        tk.Label(row, text=label, bg=CARD, fg=MUTED,
                 font=F_LABEL).pack(anchor="w")
        tk.Entry(row, textvariable=var, show=show,
                 bg=PANEL, fg=TEXT, insertbackground=ACCENT,
                 relief="flat", font=F_UI,
                 highlightthickness=1, highlightbackground=BORDER,
                 highlightcolor=ACCENT).pack(fill="x", ipady=5, pady=(4, 0))

    def _build_gw(self, parent):
        self._input(parent, "网关 IP", self.gw_ip)
        row = tk.Frame(parent, bg=CARD)
        row.pack(fill="x", pady=(0, 10))
        row.columnconfigure(0, weight=1)
        row.columnconfigure(1, weight=2)
        tk.Label(row, text="端口", bg=CARD, fg=MUTED,
                 font=F_LABEL).grid(row=0, column=0, sticky="w", padx=(0, 10))
        tk.Label(row, text="用户名", bg=CARD, fg=MUTED,
                 font=F_LABEL).grid(row=0, column=1, sticky="w")
        tk.Entry(row, textvariable=self.gw_port, bg=PANEL, fg=TEXT,
                 insertbackground=ACCENT, relief="flat", font=F_UI,
                 highlightthickness=1, highlightbackground=BORDER,
                 highlightcolor=ACCENT).grid(row=1, column=0, sticky="ew",
                                             padx=(0, 10), pady=(4, 0))
        tk.Entry(row, textvariable=self.gw_user, bg=PANEL, fg=TEXT,
                 insertbackground=ACCENT, relief="flat", font=F_UI,
                 highlightthickness=1, highlightbackground=BORDER,
                 highlightcolor=ACCENT).grid(row=1, column=1, sticky="ew",
                                             pady=(4, 0))
        self._input(parent, "密码", self.gw_pass, show="•")

    def _build_file(self, parent):
        row = tk.Frame(parent, bg=CARD)
        row.pack(fill="x")
        row.columnconfigure(0, weight=1)
        tk.Entry(row, textvariable=self.excel_path, bg=PANEL, fg=MUTED,
                 relief="flat", font=F_UI,
                 highlightthickness=1, highlightbackground=BORDER,
                 highlightcolor=ACCENT).grid(row=0, column=0, sticky="ew",
                                             ipady=5, padx=(0, 8))
        NeonButton(row, "浏览", self._browse, color=ACCENT2,
                   w=78, h=36, fg="#ffffff").grid(row=0, column=1)

    def _build_action(self, parent):
        self.btn = NeonButton(parent, "▶  开始发送", self._toggle,
                              color=ACCENT, w=320, h=46,
                              font=("Microsoft YaHei", 12, "bold"))
        self.btn.pack(fill="x", pady=(0, 8))
        NeonButton(parent, "打开失败记录", self._open_fail,
                   color=CARD_HI, w=320, h=32, fg=TEXT,
                   font=F_UI).pack(fill="x", pady=(0, 12))

        tk.Label(parent, text="发送进度", bg=CARD, fg=MUTED,
                 font=F_LABEL).pack(anchor="w")
        self.pb = ProgressBar(parent, w=320, h=10, bg=CARD)
        self.pb.pack(fill="x", pady=(6, 4))
        self.pb_text = tk.Label(parent, text="0%", bg=CARD, fg=ACCENT,
                                font=F_LABEL)
        self.pb_text.pack(anchor="e")

    def _build_recv(self, parent):
        self.recv_btn = NeonButton(parent, "▶  开始监听回复",
                                    self._toggle_listen,
                                    color=ACCENT2, w=320, h=40, fg="#ffffff",
                                    font=F_UI_B)
        self.recv_btn.pack(fill="x", pady=(0, 8))
        row = tk.Frame(parent, bg=CARD)
        row.pack(fill="x")
        NeonButton(row, "打开回复记录", self._open_reply,
                   color=CARD_HI, w=150, h=32, fg=TEXT,
                   font=F_UI).pack(side="left")
        self.recv_count = tk.Label(row, text="已收 0 条", bg=CARD,
                                   fg=ACCENT2, font=F_LABEL)
        self.recv_count.pack(side="right", pady=(8, 0))

    def _build_stats(self, parent):
        row = tk.Frame(parent, bg=CARD)
        row.pack(fill="x")
        for c in range(4):
            row.columnconfigure(c, weight=1)
        self.stat_total = self._mk_stat(row, "总人数", ACCENT2, 0)
        self.stat_ok = self._mk_stat(row, "成功", SUCCESS, 1)
        self.stat_fail = self._mk_stat(row, "失败", DANGER, 2)
        self.stat_cur = self._mk_stat(row, "进行中", ACCENT, 3)

    def _mk_stat(self, parent, label, color, col):
        box = tk.Frame(parent, bg=CARD_HI, highlightthickness=1,
                       highlightbackground=BORDER)
        box.grid(row=0, column=col, sticky="nsew", padx=(0, 10))
        tk.Label(box, text=label, bg=CARD_HI, fg=MUTED,
                 font=F_STAT_L).pack(anchor="w", padx=14, pady=(10, 0))
        lbl = tk.Label(box, text="0", bg=CARD_HI, fg=color, font=F_STAT_N)
        lbl.pack(anchor="w", padx=14, pady=(0, 8))
        return lbl

    def _build_notebook(self, parent):
        nb = ttk.Notebook(parent)
        nb.pack(fill="both", expand=True)

        f1 = tk.Frame(nb, bg=CARD)
        nb.add(f1, text=" 发送日志 ")
        self.log = tk.Text(f1, height=10, font=F_LOG2, wrap="word",
                           relief="flat", bg=PANEL, fg=TEXT,
                           padx=12, pady=10, insertbackground=ACCENT, bd=0)
        self.log.pack(fill="both", expand=True)
        self._mk_tags(self.log)
        self.log.configure(state="disabled")

        f2 = tk.Frame(nb, bg=CARD)
        nb.add(f2, text=" 回复收件箱 ")
        self.recv_text = tk.Text(f2, height=10, font=F_LOG2, wrap="word",
                                 relief="flat", bg=PANEL, fg=TEXT,
                                 padx=12, pady=10, insertbackground=ACCENT,
                                 bd=0)
        self.recv_text.pack(fill="both", expand=True)
        self._mk_tags(self.recv_text)
        self.recv_text.configure(state="disabled")

    def _mk_tags(self, w):
        w.tag_config("ok", foreground=SUCCESS)
        w.tag_config("fail", foreground=DANGER)
        w.tag_config("info", foreground=ACCENT)
        w.tag_config("warn", foreground=WARN)
        w.tag_config("dim", foreground=MUTED)

    # ---------- 发送 ----------
    def _browse(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if p:
            self.excel_path.set(p)
            self._log(f"已选择: {os.path.basename(p)}", "info")

    def _log(self, m, t="info"):
        self.log.configure(state="normal")
        self.log.insert("end", m + "\n", t)
        self.log.see("end")
        self.log.configure(state="disabled")

    def _set_status(self, text, color=MUTED, dot=None, label=None):
        self.status.config(text=text, fg=color)
        if dot:
            self.dot.set_color(dot)
        if label:
            self.dot_label.config(text=label, fg=color)

    def _toggle(self):
        if self.sending:
            self._stop_evt.set()
            return
        if not self.excel_path.get():
            messagebox.showwarning("提示", "请先选择 Excel 名单文件")
            return
        self.sending = True
        self._stop_evt.clear()
        self.ok_n = self.fail_n = self.total_n = 0
        self.stat_ok.config(text="0")
        self.stat_fail.config(text="0")
        self.stat_cur.config(text="-")
        self.stat_total.config(text="0")
        self.btn.set_text("⏸  停止发送")
        self.btn.set_color(DANGER)
        self.pb.set(0)
        self.pb_text.config(text="0%")
        self._log("──── 任务开始 ────", "dim")
        threading.Thread(target=self._send, daemon=True).start()

    def _wait(self, sec):
        self._stop_evt.wait(sec)

    def _send(self):
        ef = self.excel_path.get()
        url = f"http://{self.gw_ip.get().strip()}:{self.gw_port.get().strip()}/message"
        u = self.gw_user.get().strip()
        p = self.gw_pass.get().strip()
        try:
            df = pd.read_excel(ef, dtype=str)
            b = df.iloc[:, :4].copy()
            b.columns = ["专业", "姓名", "学号", "电话号"]
            b = b.dropna(subset=["电话号"])
            b["电话号"] = b["电话号"].astype(str).str.replace(r"\D", "", regex=True)
            studs = b.to_dict("records")
        except Exception as e:
            self.root.after(0, self._log, f"读取失败: {e}", "fail")
            self.root.after(0, self._done, "读取失败", DANGER, DANGER, "错误")
            return

        total = len(studs)
        self.total_n = total
        self.root.after(0, self.stat_total.config, {"text": str(total)})
        self.root.after(0, self._log, f"共 {total} 人待发送", "info")
        self.root.after(0, self._set_status, "● 发送中…", ACCENT, ACCENT, "发送中")

        ok, fail = [], []
        for i, row in enumerate(studs, 1):
            if self._stop_evt.is_set():
                self.root.after(0, self._log, "■ 已手动停止", "warn")
                break
            name = row["姓名"]
            phone = row["电话号"]
            self.root.after(0, self.stat_cur.config, {"text": f"{i}/{total}"})
            self.root.after(0, self._log, f"[{i}/{total}] {name}  {phone}", "info")

            success, err = False, ""
            for a in range(1, RETRY_TIMES + 1):
                if self._stop_evt.is_set():
                    break
                try:
                    ph = "+86" + phone
                    msg = MSG_TPL.format(name=row["姓名"],
                                         major=row["专业"], sid=row["学号"])
                    r = requests.post(url, auth=(u, p),
                                      json={"textMessage": {"text": msg},
                                            "phoneNumbers": [ph]},
                                      timeout=10)
                    if r.status_code in (200, 202):
                        success = True
                        break
                    err = f"HTTP {r.status_code}"
                except Exception as e:
                    err = str(e)[:60]
                if a < RETRY_TIMES:
                    self.root.after(0, self._log, f"  ↳ 第{a}次失败,重试…", "warn")
                    self._wait(RETRY_WAIT)

            rec = {"专业": row["专业"], "姓名": row["姓名"],
                   "学号": row["学号"], "电话号": phone,
                   "时间": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
            if success:
                ok.append(rec)
                self.ok_n += 1
                self.root.after(0, self.stat_ok.config, {"text": str(self.ok_n)})
                self.root.after(0, self._log, "  ✓ 发送成功", "ok")
            else:
                rec["错误"] = err
                fail.append(rec)
                self.fail_n += 1
                self.root.after(0, self.stat_fail.config, {"text": str(self.fail_n)})
                self.root.after(0, self._log, f"  ✗ {err}", "fail")

            pct = (i - 1) / total * 100 if total else 0
            self.root.after(0, self._progress, pct)
            if i < total and not self._stop_evt.is_set():
                self._wait(WAIT_BETWEEN)

        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        folder = os.path.dirname(ef)
        try:
            if ok:
                pd.DataFrame(ok).to_csv(
                    os.path.join(folder, f"发送成功_{ts}.csv"),
                    index=False, encoding="utf-8-sig")
            if fail:
                pd.DataFrame(fail).to_csv(
                    os.path.join(folder, f"发送失败_{ts}.csv"),
                    index=False, encoding="utf-8-sig")
        except Exception as e:
            self.root.after(0, self._log, f"导出失败: {e}", "warn")

        self.root.after(0, self._progress, 100)
        self.root.after(0, self.stat_cur.config,
                        {"text": f"{self.ok_n + self.fail_n}/{total}"})
        self.root.after(0, self._log,
                       f"──── 完成: 成功 {len(ok)}  失败 {len(fail)} ────", "info")
        self.root.after(0, self._done,
                        f"● 完成  成功 {len(ok)} / 失败 {len(fail)}",
                        SUCCESS, SUCCESS, "完成")

    def _progress(self, pct):
        self.pb.set(pct)
        self.pb_text.config(text=f"{pct:.0f}%")

    def _done(self, text, color, dot, label):
        self.sending = False
        self._stop_evt.clear()
        self.btn.set_text("▶  开始发送")
        self.btn.set_color(ACCENT)
        self._set_status(text, color, dot, label)

    def _open_fail(self):
        f = filedialog.askdirectory()
        if not f:
            return
        fs = [x for x in os.listdir(f)
              if x.startswith("发送失败") and x.endswith(".csv")]
        if fs:
            os.startfile(os.path.join(f, sorted(fs)[-1]))
        else:
            messagebox.showinfo("提示", "该目录下没有失败记录")

    # ---------- 接收 ----------
    def _load_seen_async(self):
        def work():
            try:
                if os.path.exists(self.recv_csv):
                    df = pd.read_csv(self.recv_csv, dtype=str)
                    if "消息ID" in df.columns:
                        self.seen_ids.update(df["消息ID"].dropna().tolist())
                    self.reply_n = len(df)
            except Exception:
                pass
            self.root.after(0, self._refresh_recv_count)
        threading.Thread(target=work, daemon=True).start()

    def _refresh_recv_count(self):
        self.recv_count.config(text=f"已收 {self.reply_n} 条")

    def _toggle_listen(self):
        if self.listening:
            # —— 停止：立即响应 UI，不等线程退出 ——
            self.listening = False
            self._recv_evt.set()
            self.recv_btn.set_text("▶  开始监听回复")
            self.recv_btn.set_color(ACCENT2)
            self._log("■ 正在停止监听…", "warn")
            if not self.sending:
                self._set_status("● 正在停止…", MUTED, MUTED, "停止中")
            return
        # —— 开始：新代际，防止旧线程干扰 ——
        self._listen_gen += 1
        gen = self._listen_gen
        self.listening = True
        self._recv_evt.clear()
        self.recv_btn.set_text("⏸  停止监听")
        self.recv_btn.set_color(DANGER)
        self._log("──── 开始监听回复 ────", "dim")
        threading.Thread(target=self._recv_loop, args=(gen,), daemon=True).start()

    def _recv_loop(self, gen):
        base = f"http://{self.gw_ip.get().strip()}:{self.gw_port.get().strip()}"
        u = self.gw_user.get().strip()
        p = self.gw_pass.get().strip()

        data = self._fetch_messages(base, u, p, gen)
        if self._is_stale(gen):
            return
        if data is None:
            self.root.after(0, self._log,
                           "✗ 连不上网关，无法拉取回复", "fail")
            self.root.after(0, self._recv_stopped, "接收失败", DANGER)
            return
        self.root.after(0, self._log,
                        f"✓ 已连接，每 {POLL_INTERVAL} 秒轮询…", "ok")

        while not self._recv_evt.is_set():
            data = self._fetch_messages(base, u, p, gen)
            if self._is_stale(gen):
                return
            if data:
                for msg in self._parse_messages(data):
                    if self._recv_evt.is_set():
                        break
                    mid = str(msg.get("id") or msg.get("messageId")
                              or msg.get("_id") or hash(str(msg)))
                    if mid in self.seen_ids:
                        continue
                    self.seen_ids.add(mid)
                    sender = (msg.get("phoneNumber") or msg.get("sender")
                              or msg.get("from") or msg.get("address") or "未知")
                    text = (msg.get("text") or msg.get("body")
                            or msg.get("content") or "")
                    ts = (msg.get("receivedAt") or msg.get("timestamp")
                          or msg.get("date")
                          or datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                    row = {"消息ID": mid, "发送人": sender, "内容": text,
                           "时间": str(ts),
                           "收到时间": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
                    self._append_reply_csv(row)
                    self.reply_n += 1
                    self.root.after(0, self._on_reply, row)
            # 分段等待，1 秒粒度响应停止
            waited = 0
            while waited < POLL_INTERVAL and not self._recv_evt.is_set():
                self._recv_evt.wait(1)
                waited += 1

        self.root.after(0, self._log, "■ 已停止监听", "warn")
        self.root.after(0, self._recv_stopped, "监听已停止", MUTED)

    def _is_stale(self, gen):
        return gen != self._listen_gen

    def _fetch_messages(self, base, u, p, gen):
        for ep in ("/messages", "/api/messages",
                   "/message/inbox", "/api/messages/inbox"):
            if self._recv_evt.is_set() or self._is_stale(gen):
                return None
            try:
                r = requests.get(f"{base}{ep}", auth=(u, p), timeout=HTTP_TIMEOUT)
                if r.status_code == 200:
                    return r.json()
            except Exception:
                continue
        return None

    def _parse_messages(self, data):
        if isinstance(data, list):
            return data
        if isinstance(data, dict):
            for key in ("messages", "data", "items", "results"):
                if key in data and isinstance(data[key], list):
                    return data[key]
        return []

    def _append_reply_csv(self, row):
        try:
            new_row = pd.DataFrame([row])
            if os.path.exists(self.recv_csv):
                df = pd.read_csv(self.recv_csv, dtype=str)
                df = pd.concat([df, new_row], ignore_index=True)
            else:
                df = new_row
            df.to_csv(self.recv_csv, index=False, encoding="utf-8-sig")
        except Exception:
            pass

    def _on_reply(self, row):
        self.recv_count.config(text=f"已收 {self.reply_n} 条")
        self.recv_text.configure(state="normal")
        self.recv_text.insert("end",
                              f"[{row['收到时间']}]  {row['发送人']}\n"
                              f"  {row['内容']}\n\n", "ok")
        self.recv_text.see("end")
        self.recv_text.configure(state="disabled")
        preview = row["内容"][:24].replace("\n", " ")
        self._log(f"← 回复 [{row['发送人']}] {preview}", "ok")

    def _recv_stopped(self, label="就绪", dot_color=MUTED):
        # 线程退出后的收尾：只清 event，不改按钮（按钮已由 toggle 即时更新）
        self._recv_evt.clear()
        if not self.listening and not self.sending:
            self._set_status(f"● {label}", dot_color, dot_color, label)

    def _open_reply(self):
        if os.path.exists(self.recv_csv):
            os.startfile(self.recv_csv)
        else:
            messagebox.showinfo("提示", "还没有回复记录")

    # ---------- 呼吸灯 ----------
    def _pulse(self):
        if self.sending:
            self.dot.blink(not self.dot._on)
            self.root.after(500, self._pulse)
        else:
            self.dot._on = True
            self.dot._draw()


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
