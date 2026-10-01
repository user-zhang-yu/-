# -*- coding: utf-8 -*-
"""
SMS Gateway Pro —— 模板版（单人发送 / Excel 批量发送 / 接收监听 / 联系人记忆）
- Per-Monitor DPI 感知，高分屏不模糊
- 单人：手机号 + 消息内容直发
- 批量：Excel 名单 + {表头} 占位符模板
- 接收：轮询收件箱，保存到自选目录
- 记忆：SQLite 自动记录所有联系人，可搜索/回填
"""
import ctypes

# ===== DPI 感知（必须在创建窗口之前）=====
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

import sys
import os
import json
import sqlite3
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import requests
import pandas as pd
import datetime
import threading


def app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


# ===================== 默认配置 =====================
CFG_PATH = os.path.join(app_dir(), "gateway_config.json")

DEFAULT_CFG = {
    "ip": "10.241.110.111",
    "port": "8080",
    "user": "HappyStart",
    "pass": "12345678",
    "recv_dir": app_dir(),
}

POLL_INTERVAL = 5
HTTP_TIMEOUT = 5
WAIT_BETWEEN = 1.5
RETRY_TIMES = 2
RETRY_WAIT = 2

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

F_UI    = ("Microsoft YaHei", 10)
F_UI_B  = ("Microsoft YaHei", 10, "bold")
F_H1    = ("Microsoft YaHei", 15, "bold")
F_LABEL = ("Microsoft YaHei", 9)
F_STAT_N= ("Segoe UI", 20, "bold")
F_STAT_L= ("Microsoft YaHei", 8)
F_LOG   = ("Consolas", 9)


def round_rect(cv, x1, y1, x2, y2, r, **kw):
    pts = [x1+r, y1, x2-r, y1, x2, y1, x2, y1+r,
           x2, y2-r, x2, y2, x2-r, y2, x1+r, y2,
           x1, y2, x1, y2-r, x1, y1+r, x1, y1]
    return cv.create_polygon(pts, smooth=True, **kw)


def lighten(hex_color, amount=30):
    h = hex_color.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return f"#{min(255,r+amount):02x}{min(255,g+amount):02x}{min(255,b+amount):02x}"


# ===================== 组件 =====================
class Card(tk.Frame):
    def __init__(self, parent, title=None, accent=ACCENT, padding=14, **kw):
        super().__init__(parent, bg=CARD, highlightthickness=1,
                         highlightbackground=BORDER, bd=0, **kw)
        self.body = None
        if title:
            head = tk.Frame(self, bg=CARD)
            head.pack(fill="x", padx=padding, pady=(10, 6))
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
                 w=200, h=38, fg="#04121a", font=None):
        super().__init__(parent, width=w, height=h, bg=BG,
                         highlightthickness=0, bd=0)
        self.command = command
        self.color = color
        self.fg = fg
        self.text = text
        self.w, self.h = w, h
        self.font = font or F_UI_B
        self.hover = False
        self._draw()
        self.bind("<Enter>", lambda e: (setattr(self, "hover", True), self._draw()))
        self.bind("<Leave>", lambda e: (setattr(self, "hover", False), self._draw()))
        self.bind("<Button-1>", lambda e: self.command() if self.command else None)

    def _draw(self):
        self.delete("all")
        c = lighten(self.color, 25) if self.hover else self.color
        round_rect(self, 1.5, 1.5, self.w-1.5, self.h-1.5, 9, fill=c, outline="")
        self.create_rectangle(12, 5, self.w-12, 6, fill="#ffffff", outline="")
        self.create_text(self.w/2, self.h/2+1, text=self.text,
                         fill=self.fg, font=self.font)

    def set_text(self, t):
        self.text = t; self._draw()

    def set_color(self, c):
        self.color = c; self._draw()


class StatusDot(tk.Canvas):
    def __init__(self, parent, color=MUTED, size=12):
        super().__init__(parent, width=size, height=size, bg=BG,
                         highlightthickness=0, bd=0)
        self.size = size; self.color = color
        self._draw()

    def _draw(self):
        self.delete("all")
        s = self.size
        self.create_oval(0, 0, s, s, fill="", outline=self.color, width=1)
        self.create_oval(2, 2, s-2, s-2, fill=self.color, outline="")

    def set_color(self, c):
        self.color = c; self._draw()


# ===================== 联系人数据库 =====================
class ContactDB:
    def __init__(self, path):
        self.conn = sqlite3.connect(path, check_same_thread=False)
        self.conn.execute("""
            CREATE TABLE IF NOT EXISTS contacts (
                phone TEXT PRIMARY KEY,
                name TEXT DEFAULT '',
                send_count INTEGER DEFAULT 0,
                recv_count INTEGER DEFAULT 0,
                last_contact TEXT DEFAULT ''
            )
        """)
        self.conn.commit()
        self.lock = threading.Lock()

    def upsert_sent(self, phone, name=""):
        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with self.lock:
            self.conn.execute("""
                INSERT INTO contacts (phone, name, send_count, last_contact)
                VALUES (?, ?, 1, ?)
                ON CONFLICT(phone) DO UPDATE SET
                    send_count = send_count + 1,
                    last_contact = excluded.last_contact,
                    name = CASE WHEN excluded.name <> '' THEN excluded.name ELSE contacts.name END
            """, (phone, name, now))
            self.conn.commit()

    def upsert_recv(self, phone, text=""):
        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with self.lock:
            self.conn.execute("""
                INSERT INTO contacts (phone, recv_count, last_contact)
                VALUES (?, 1, ?)
                ON CONFLICT(phone) DO UPDATE SET
                    recv_count = recv_count + 1,
                    last_contact = excluded.last_contact
            """, (phone, now))
            self.conn.commit()

    def search(self, kw=""):
        with self.lock:
            if kw:
                q = f"%{kw}%"
                rows = self.conn.execute("""
                    SELECT phone, name, send_count, recv_count, last_contact
                    FROM contacts WHERE phone LIKE ? OR name LIKE ?
                    ORDER BY last_contact DESC LIMIT 500
                """, (q, q)).fetchall()
            else:
                rows = self.conn.execute("""
                    SELECT phone, name, send_count, recv_count, last_contact
                    FROM contacts ORDER BY last_contact DESC LIMIT 500
                """).fetchall()
        return rows

    def suggest_phones(self, kw=""):
        with self.lock:
            q = f"%{kw}%"
            rows = self.conn.execute("""
                SELECT phone, name FROM contacts
                WHERE phone LIKE ? OR name LIKE ?
                ORDER BY last_contact DESC LIMIT 20
            """, (q, q)).fetchall()
        return [{"phone": r[0], "name": r[1]} for r in rows]


# ===================== 主应用 =====================
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("SMS Gateway Pro — 模板版")
        W, H = 1180, 820
        self.root.minsize(980, 700)
        self.root.configure(bg=BG)
        x = (self.root.winfo_screenwidth() - W) // 2
        y = (self.root.winfo_screenheight() - H) // 2
        self.root.geometry(f"{W}x{H}+{x}+{y}")

        # 配置
        self.cfg = dict(DEFAULT_CFG)
        self._load_cfg()

        # 变量
        self.gw_ip = tk.StringVar(value=self.cfg["ip"])
        self.gw_port = tk.StringVar(value=self.cfg["port"])
        self.gw_user = tk.StringVar(value=self.cfg["user"])
        self.gw_pass = tk.StringVar(value=self.cfg["pass"])
        self.recv_dir = tk.StringVar(value=self.cfg["recv_dir"])

        self.excel_path = tk.StringVar()
        self.single_phone = tk.StringVar()
        self.single_phone_combo = None

        # 状态
        self.connection_ok = False
        self.test_btn = None
        self.sending = False
        self.batch_sending = False
        self._stop_evt = threading.Event()
        self.listening = False
        self._recv_evt = threading.Event()
        self._listen_gen = 0
        self.reply_n = 0
        self.seen_ids = set()
        self.ok_n = self.fail_n = self.total_n = 0

        # 数据库
        self.db = ContactDB(os.path.join(app_dir(), "contacts.db"))

        self._setup_style()
        self._build_ui()
        self._load_seen_async()

    # ---------- 配置持久化 ----------
    def _load_cfg(self):
        try:
            if os.path.exists(CFG_PATH):
                with open(CFG_PATH, "r", encoding="utf-8") as f:
                    self.cfg.update(json.load(f))
        except Exception:
            pass

    def _save_cfg(self):
        try:
            self.cfg["ip"] = self.gw_ip.get().strip()
            self.cfg["port"] = self.gw_port.get().strip()
            self.cfg["user"] = self.gw_user.get().strip()
            self.cfg["pass"] = self.gw_pass.get().strip()
            self.cfg["recv_dir"] = self.recv_dir.get().strip()
            with open(CFG_PATH, "w", encoding="utf-8") as f:
                json.dump(self.cfg, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    # ---------- ttk 样式 ----------
    def _setup_style(self):
        s = ttk.Style()
        try:
            s.theme_use("clam")
        except Exception:
            pass
        s.configure("TNotebook", background=BG, borderwidth=0, tabmargins=(10, 8, 10, 0))
        s.configure("TNotebook.Tab", background=PANEL, foreground=MUTED,
                    padding=(18, 8), font=F_UI_B, borderwidth=0)
        s.map("TNotebook.Tab",
              background=[("selected", CARD_HI)],
              foreground=[("selected", ACCENT)])
        s.configure("Treeview", background=CARD, fieldbackground=CARD,
                    foreground=TEXT, borderwidth=0, rowheight=26, font=F_UI)
        s.configure("Treeview.Heading", background=PANEL, foreground=ACCENT,
                    font=F_UI_B, borderwidth=0, relief="flat")
        s.map("Treeview", background=[("selected", ACCENT2)],
              foreground=[("selected", "#ffffff")])
        s.configure("TCombobox", fieldbackground=PANEL, background=PANEL,
                    foreground=TEXT, arrowcolor=ACCENT)

    # ---------- UI ----------
    def _build_ui(self):
        tk.Frame(self.root, bg=ACCENT, height=2).pack(fill="x")

        # 标题栏
        bar = tk.Frame(self.root, bg=BG, height=54)
        bar.pack(fill="x"); bar.pack_propagate(False)
        inner = tk.Frame(bar, bg=BG)
        inner.pack(fill="both", expand=True, padx=20)
        logo = tk.Canvas(inner, width=28, height=28, bg=BG, highlightthickness=0)
        logo.pack(side="left", padx=(0, 10))
        round_rect(logo, 0, 0, 28, 28, 6, fill=ACCENT, outline="")
        logo.create_text(14, 15, text="⚡", fill="#04121a", font=("Segoe UI", 12, "bold"))
        tk.Label(inner, text="SMS GATEWAY PRO", bg=BG, fg=TEXT,
                 font=F_H1).pack(side="left")
        tk.Label(inner, text="  模板版 · 单人/批量/接收/记忆", bg=BG, fg=MUTED,
                 font=F_LABEL).pack(side="left", padx=(6, 0), pady=(6, 0))
        self.dot = StatusDot(inner, color=MUTED)
        self.dot.pack(side="right", padx=(0, 8))
        self.dot_label = tk.Label(inner, text="就绪", bg=BG, fg=MUTED, font=F_LABEL)
        self.dot_label.pack(side="right")

        # 网关配置条
        cfgbar = tk.Frame(self.root, bg=CARD, highlightthickness=1,
                           highlightbackground=BORDER)
        cfgbar.pack(fill="x", padx=16, pady=(10, 8))
        cfgin = tk.Frame(cfgbar, bg=CARD)
        cfgin.pack(fill="x", padx=14, pady=8)
        for i in range(8):
            cfgin.columnconfigure(i, weight=0)
        tk.Label(cfgin, text="IP", bg=CARD, fg=MUTED, font=F_LABEL).grid(row=0, column=0, padx=(0, 6))
        tk.Entry(cfgin, textvariable=self.gw_ip, width=14, bg=PANEL, fg=TEXT,
                 insertbackground=ACCENT, relief="flat", font=F_UI,
                 highlightthickness=1, highlightbackground=BORDER).grid(row=0, column=1, padx=(0, 12), ipady=3)
        tk.Label(cfgin, text="端口", bg=CARD, fg=MUTED, font=F_LABEL).grid(row=0, column=2, padx=(0, 6))
        tk.Entry(cfgin, textvariable=self.gw_port, width=7, bg=PANEL, fg=TEXT,
                 insertbackground=ACCENT, relief="flat", font=F_UI,
                 highlightthickness=1, highlightbackground=BORDER).grid(row=0, column=3, padx=(0, 12), ipady=3)
        tk.Label(cfgin, text="账号", bg=CARD, fg=MUTED, font=F_LABEL).grid(row=0, column=4, padx=(0, 6))
        tk.Entry(cfgin, textvariable=self.gw_user, width=10, bg=PANEL, fg=TEXT,
                 insertbackground=ACCENT, relief="flat", font=F_UI,
                 highlightthickness=1, highlightbackground=BORDER).grid(row=0, column=5, padx=(0, 12), ipady=3)
        tk.Label(cfgin, text="密码", bg=CARD, fg=MUTED, font=F_LABEL).grid(row=0, column=6, padx=(0, 6))
        tk.Entry(cfgin, textvariable=self.gw_pass, width=10, show="•", bg=PANEL, fg=TEXT,
                 insertbackground=ACCENT, relief="flat", font=F_UI,
                 highlightthickness=1, highlightbackground=BORDER).grid(row=0, column=7, padx=(0, 0), ipady=3)
        self.test_btn = NeonButton(cfgin, "测试连接", self._quick_test, color=ACCENT2,
                                   w=110, h=32, fg="#ffffff")
        self.test_btn.grid(row=0, column=8, padx=(14, 0))
        # 配置变更 → 连接失效
        for v in (self.gw_ip, self.gw_port, self.gw_user, self.gw_pass):
            v.trace_add("write", lambda *a: self._invalidate_conn())

        # 主 Notebook
        nb = ttk.Notebook(self.root)
        nb.pack(fill="both", expand=True, padx=16, pady=(0, 8))

        # 标签1：单人发送
        tab1 = tk.Frame(nb, bg=BG)
        nb.add(tab1, text=" 单人发送 ")
        self._build_single(tab1)

        # 标签2：批量发送
        tab2 = tk.Frame(nb, bg=BG)
        nb.add(tab2, text=" Excel 批量发送 ")
        self._build_batch(tab2)

        # 标签3：接收监听
        tab3 = tk.Frame(nb, bg=BG)
        nb.add(tab3, text=" 接收监听 ")
        self._build_recv(tab3)

        # 标签4：联系人
        tab4 = tk.Frame(nb, bg=BG)
        nb.add(tab4, text=" 联系人记忆 ")
        self._build_contacts(tab4)

        # 标签5：连接检测
        tab5 = tk.Frame(nb, bg=BG)
        nb.add(tab5, text=" 连接检测 ")
        self._build_diag(tab5)

        # 底部状态栏
        foot = tk.Frame(self.root, bg=PANEL, height=26)
        foot.pack(fill="x", side="bottom"); foot.pack_propagate(False)
        self.status = tk.Label(foot, text="● 就绪", bg=PANEL, fg=MUTED,
                               font=F_LABEL, anchor="w")
        self.status.pack(side="left", padx=16)
        self.hint = tk.Label(foot, text="", bg=PANEL, fg=MUTED, font=F_LABEL)
        self.hint.pack(side="right", padx=16)

    # ---------- 单人发送 ----------
    def _build_single(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)

        c1 = Card(parent, "发送对象", accent=ACCENT)
        c1.grid(row=0, column=0, sticky="ew", padx=4, pady=(8, 8))
        row = tk.Frame(c1.body, bg=CARD)
        row.pack(fill="x")
        row.columnconfigure(1, weight=1)
        tk.Label(row, text="手机号", bg=CARD, fg=MUTED, font=F_LABEL).grid(row=0, column=0, padx=(0, 8), pady=4)
        self.single_phone_combo = ttk.Combobox(row, textvariable=self.single_phone,
                                                font=F_UI, width=24)
        self.single_phone_combo.grid(row=0, column=1, sticky="w", pady=4)
        self.single_phone_combo.bind("<KeyRelease>", self._on_phone_type)
        self.single_phone_combo.bind("<<ComboboxSelected>>", self._on_phone_pick)
        NeonButton(row, "从联系人选", self._pick_contact, color=ACCENT2,
                   w=130, h=32, fg="#ffffff").grid(row=0, column=2, padx=(10, 0))

        c2 = Card(parent, "消息内容", accent=ACCENT2)
        c2.grid(row=1, column=0, sticky="nsew", padx=4, pady=(0, 8))
        self.single_msg = tk.Text(c2.body, height=8, font=F_UI, wrap="word",
                                  relief="flat", bg=PANEL, fg=TEXT,
                                  padx=10, pady=8, insertbackground=ACCENT, bd=0)
        self.single_msg.pack(fill="both", expand=True)

        # 发送按钮 + 日志
        c3 = Card(parent, "发送结果", accent=SUCCESS)
        c3.grid(row=2, column=0, sticky="ew", padx=4, pady=(0, 8))
        btns = tk.Frame(c3.body, bg=CARD)
        btns.pack(fill="x", pady=(0, 8))
        self.single_btn = NeonButton(btns, "▶  立即发送", self._send_single,
                                     color=ACCENT, w=200, h=42,
                                     font=("Microsoft YaHei", 11, "bold"))
        self.single_btn.pack(side="left")
        self.single_result = tk.Label(btns, text="", bg=CARD, fg=MUTED, font=F_UI_B)
        self.single_result.pack(side="left", padx=16)

    def _on_phone_type(self, e):
        kw = self.single_phone.get()
        if len(kw) >= 1:
            sug = self.db.suggest_phones(kw)
            vals = [f"{x['phone']}  {x['name']}" if x['name'] else x['phone'] for x in sug]
            self.single_phone_combo["values"] = vals
            # 不自动弹出，避免打扰

    def _on_phone_pick(self, e):
        v = self.single_phone_combo.get()
        phone = v.split("  ")[0].strip()
        self.single_phone.set(phone)

    def _pick_contact(self):
        # 跳到联系人标签页
        nb = self._find_notebook(self.root)
        if nb:
            nb.select(3)
        self._refresh_contacts()

    def _find_notebook(self, w):
        for child in w.winfo_children():
            if isinstance(child, ttk.Notebook):
                return child
            r = self._find_notebook(child)
            if r:
                return r
        return None

    # ---------- 批量发送 ----------
    def _build_batch(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(2, weight=1)

        # 文件选择
        c1 = Card(parent, "Excel 名单文件", accent=ACCENT)
        c1.grid(row=0, column=0, sticky="ew", padx=4, pady=(8, 8))
        row = tk.Frame(c1.body, bg=CARD)
        row.pack(fill="x")
        row.columnconfigure(0, weight=1)
        tk.Entry(row, textvariable=self.excel_path, bg=PANEL, fg=MUTED,
                 relief="flat", font=F_UI,
                 highlightthickness=1, highlightbackground=BORDER).grid(row=0, column=0, sticky="ew", ipady=4, padx=(0, 8))
        NeonButton(row, "浏览", self._browse_excel, color=ACCENT2,
                   w=78, h=34, fg="#ffffff").grid(row=0, column=1)
        self.batch_cols_label = tk.Label(c1.body, text="表头：未加载", bg=CARD,
                                         fg=MUTED, font=F_LABEL, anchor="w")
        self.batch_cols_label.pack(fill="x", pady=(8, 6))

        # 电话列选择
        ph_row = tk.Frame(c1.body, bg=CARD)
        ph_row.pack(fill="x")
        tk.Label(ph_row, text="电话列：", bg=CARD, fg=MUTED,
                 font=F_LABEL).pack(side="left", padx=(0, 6))
        self.phone_col_var = tk.StringVar()
        self.phone_combo = ttk.Combobox(ph_row, textvariable=self.phone_col_var,
                                        font=F_UI, width=20, state="readonly")
        self.phone_combo.pack(side="left")
        tk.Label(ph_row, text="（选哪一列作为接收短信的手机号）",
                 bg=CARD, fg=MUTED, font=F_LABEL).pack(side="left", padx=(8, 0))

        # 模板
        c2 = Card(parent, "消息模板（用 {表头列名} 占位，如 {姓名}）", accent=ACCENT2)
        c2.grid(row=1, column=0, sticky="ew", padx=4, pady=(0, 8))
        self.batch_tpl = tk.Text(c2.body, height=5, font=F_UI, wrap="word",
                                 relief="flat", bg=PANEL, fg=TEXT,
                                 padx=10, pady=8, insertbackground=ACCENT, bd=0)
        self.batch_tpl.pack(fill="x")
        self.batch_tpl.insert("1.0", "【电子协会】恭喜{姓名}同学被正式录取，学号{学号}，期待你的到来！")

        # 预览 + 发送
        c3 = Card(parent, "预览与发送", accent=SUCCESS)
        c3.grid(row=2, column=0, sticky="nsew", padx=4, pady=(0, 8))
        btns = tk.Frame(c3.body, bg=CARD)
        btns.pack(fill="x", pady=(0, 8))
        NeonButton(btns, "👁 预览第一条", self._preview_batch, color=ACCENT2,
                   w=150, h=36, fg="#ffffff").pack(side="left", padx=(0, 10))
        self.batch_btn = NeonButton(btns, "▶  开始批量发送", self._toggle_batch,
                                    color=ACCENT, w=180, h=36,
                                    font=F_UI_B)
        self.batch_btn.pack(side="left", padx=(0, 10))
        self.batch_count = tk.Label(btns, text="", bg=CARD, fg=MUTED, font=F_UI_B)
        self.batch_count.pack(side="left")

        self.batch_preview = tk.Label(c3.body, text="", bg=PANEL, fg=TEXT,
                                      font=F_UI, justify="left", anchor="nw",
                                      wraplength=800)
        self.batch_preview.pack(fill="both", expand=True, pady=(4, 0))
        self.batch_preview.configure(bg=PANEL)

        # 进度
        self.batch_progress = tk.Canvas(c3.body, height=8, bg=CARD,
                                        highlightthickness=0, bd=0)
        self.batch_progress.pack(fill="x", pady=(8, 0))

    def _browse_excel(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if p:
            self.excel_path.set(p)
            self._load_excel_cols()

    def _load_excel_cols(self):
        try:
            df = pd.read_excel(self.excel_path.get(), dtype=str, nrows=1)
            cols = [str(c) for c in df.columns]
            self.batch_cols_label.config(text="表头（可用于 {列名}）：" + "、".join(cols),
                                         fg=ACCENT)
            self.phone_combo["values"] = cols
            # 自动预选含"电话/手机"的列
            guess = next((c for c in cols if "电话" in c or "手机" in c), None)
            if guess:
                self.phone_col_var.set(guess)
            elif cols:
                self.phone_col_var.set(cols[-1])
        except Exception as e:
            self.batch_cols_label.config(text=f"读取失败：{e}", fg=DANGER)

    def _preview_batch(self):
        try:
            df = pd.read_excel(self.excel_path.get(), dtype=str)
            if len(df) == 0:
                self.batch_preview.config(text="（空表）", fg=WARN)
                return
            row = df.iloc[0].to_dict()
            tpl = self.batch_tpl.get("1.0", "end").strip()
            text = self._render_tpl(tpl, row)
            self.batch_preview.config(text=text, fg=TEXT)
        except Exception as e:
            self.batch_preview.config(text=f"预览失败：{e}", fg=DANGER)

    def _render_tpl(self, tpl, row):
        try:
            return tpl.format(**{k: (v if pd.notna(v) else "") for k, v in row.items()})
        except KeyError as e:
            return f"[模板占位符 {e} 在表中不存在]"
        except Exception as e:
            return f"[模板渲染错误：{e}]"

    # ---------- 接收监听 ----------
    def _build_recv(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(2, weight=1)

        c1 = Card(parent, "保存位置", accent=ACCENT)
        c1.grid(row=0, column=0, sticky="ew", padx=4, pady=(8, 8))
        row = tk.Frame(c1.body, bg=CARD)
        row.pack(fill="x")
        row.columnconfigure(0, weight=1)
        tk.Entry(row, textvariable=self.recv_dir, bg=PANEL, fg=TEXT,
                 relief="flat", font=F_UI,
                 highlightthickness=1, highlightbackground=BORDER).grid(row=0, column=0, sticky="ew", ipady=4, padx=(0, 8))
        NeonButton(row, "选择", self._pick_recv_dir, color=ACCENT2,
                   w=78, h=34, fg="#ffffff").grid(row=0, column=1)

        c2 = Card(parent, "监听控制", accent=ACCENT2)
        c2.grid(row=1, column=0, sticky="ew", padx=4, pady=(0, 8))
        btns = tk.Frame(c2.body, bg=CARD)
        btns.pack(fill="x")
        self.recv_btn = NeonButton(btns, "▶  开始监听", self._toggle_listen,
                                   color=ACCENT, w=180, h=40,
                                   font=("Microsoft YaHei", 11, "bold"))
        self.recv_btn.pack(side="left", padx=(0, 12))
        self.recv_count = tk.Label(btns, text="本次已收 0 条", bg=CARD,
                                   fg=ACCENT, font=F_UI_B)
        self.recv_count.pack(side="left")

        c3 = Card(parent, "接收记录", accent=SUCCESS)
        c3.grid(row=2, column=0, sticky="nsew", padx=4, pady=(0, 8))
        self.recv_text = tk.Text(c3.body, font=F_LOG, wrap="word",
                                 relief="flat", bg=PANEL, fg=TEXT,
                                 padx=10, pady=8, insertbackground=ACCENT, bd=0)
        self.recv_text.pack(fill="both", expand=True)
        self.recv_text.tag_config("ok", foreground=SUCCESS)
        self.recv_text.tag_config("info", foreground=ACCENT)
        self.recv_text.tag_config("dim", foreground=MUTED)

    def _pick_recv_dir(self):
        d = filedialog.askdirectory()
        if d:
            self.recv_dir.set(d)
            self._save_cfg()

    # ---------- 联系人 ----------
    def _build_contacts(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)

        bar = tk.Frame(parent, bg=BG)
        bar.grid(row=0, column=0, sticky="ew", padx=4, pady=(8, 4))
        bar.columnconfigure(1, weight=1)
        tk.Label(bar, text="搜索：", bg=BG, fg=MUTED, font=F_UI).grid(row=0, column=0, padx=(0, 6))
        self.contact_search = tk.StringVar()
        ent = tk.Entry(bar, textvariable=self.contact_search, bg=PANEL, fg=TEXT,
                       relief="flat", font=F_UI,
                       highlightthickness=1, highlightbackground=BORDER)
        ent.grid(row=0, column=1, sticky="ew", ipady=4)
        ent.bind("<KeyRelease>", lambda e: self._refresh_contacts())
        NeonButton(bar, "刷新", self._refresh_contacts, color=ACCENT2,
                   w=70, h=32, fg="#ffffff").grid(row=0, column=2, padx=(8, 0))

        cols = ("phone", "name", "send", "recv", "last")
        self.tree = ttk.Treeview(parent, columns=cols, show="headings")
        self.tree.grid(row=1, column=0, sticky="nsew", padx=4, pady=(0, 8))
        self.tree.heading("phone", text="手机号")
        self.tree.heading("name", text="姓名")
        self.tree.heading("send", text="发送")
        self.tree.heading("recv", text="接收")
        self.tree.heading("last", text="最近联系")
        self.tree.column("phone", width=130, anchor="w")
        self.tree.column("name", width=120, anchor="w")
        self.tree.column("send", width=60, anchor="center")
        self.tree.column("recv", width=60, anchor="center")
        self.tree.column("last", width=160, anchor="w")
        self.tree.bind("<Double-1>", self._contact_use)

        tip = tk.Label(parent, text="双击某行 → 回填到「单人发送」", bg=BG,
                       fg=MUTED, font=F_LABEL)
        tip.grid(row=2, column=0, sticky="w", padx=4, pady=(0, 4))

    def _refresh_contacts(self):
        kw = self.contact_search.get().strip()
        rows = self.db.search(kw)
        for item in self.tree.get_children():
            self.tree.delete(item)
        for r in rows:
            self.tree.insert("", "end", values=r)

    def _contact_use(self, e):
        sel = self.tree.selection()
        if not sel:
            return
        vals = self.tree.item(sel[0], "values")
        phone = vals[0]
        self.single_phone.set(phone)
        nb = self._find_notebook(self.root)
        if nb:
            nb.select(0)
        self._set_status(f"● 已回填 {phone}", ACCENT, ACCENT, "就绪")

    # ---------- 快速连接测试 ----------
    def _invalidate_conn(self):
        if self.connection_ok:
            self.connection_ok = False
            if self.test_btn:
                self.test_btn.set_text("测试连接"); self.test_btn.set_color(ACCENT2)

    def _quick_test(self):
        self.test_btn.set_text("测试中…"); self.test_btn.set_color(WARN)
        threading.Thread(target=self._quick_test_worker, daemon=True).start()

    def _quick_test_worker(self):
        import socket
        ip = self.gw_ip.get().strip()
        port = self.gw_port.get().strip()
        user = self.gw_user.get().strip()
        pwd = self.gw_pass.get().strip()
        try:
            sock = socket.create_connection((ip, int(port)), timeout=3)
            sock.close()
        except Exception as e:
            self.root.after(0, lambda: self._on_conn_result(False, f"网络不通：{e}"))
            return
        try:
            r = requests.get(f"http://{ip}:{port}/messages", auth=(user, pwd), timeout=5)
            if r.status_code == 401:
                self.root.after(0, lambda: self._on_conn_result(False, "账号或密码错误"))
            elif r.status_code in (200, 404, 405):
                self.root.after(0, lambda: self._on_conn_result(True, ""))
            else:
                self.root.after(0, lambda: self._on_conn_result(False, f"HTTP {r.status_code}"))
        except Exception as e:
            self.root.after(0, lambda: self._on_conn_result(False, str(e)[:60]))

    def _on_conn_result(self, ok, msg):
        self.connection_ok = ok
        if ok:
            self.test_btn.set_text("✓ 已连接"); self.test_btn.set_color(SUCCESS)
            self._set_status("● 已连接网关", SUCCESS, SUCCESS, "已连接")
        else:
            self.test_btn.set_text("✗ 连接失败"); self.test_btn.set_color(DANGER)
            self._set_status(f"● {msg}", DANGER, DANGER, "未连接")

    # ---------- 连接检测 ----------
    def _build_diag(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)

        top = tk.Frame(parent, bg=BG)
        top.grid(row=0, column=0, sticky="ew", padx=4, pady=(8, 8))
        NeonButton(top, "🔍 开始检测", self._run_diag, color=ACCENT,
                   w=160, h=38, font=F_UI_B).pack(side="left")
        self.diag_summary = tk.Label(top, text="点击按钮检测网关连接", bg=BG,
                                    fg=MUTED, font=F_LABEL)
        self.diag_summary.pack(side="left", padx=16)

        card = Card(parent, "检测报告", accent=ACCENT2)
        card.grid(row=1, column=0, sticky="nsew", padx=4, pady=(0, 8))
        self.diag_text = tk.Text(card.body, wrap="word", font=F_UI,
                                 relief="flat", bg=PANEL, fg=TEXT,
                                 padx=12, pady=10, bd=0, state="disabled")
        self.diag_text.pack(fill="both", expand=True)
        self.diag_text.tag_config("ok", foreground=SUCCESS)
        self.diag_text.tag_config("fail", foreground=DANGER)
        self.diag_text.tag_config("warn", foreground=WARN)
        self.diag_text.tag_config("head", foreground=ACCENT, font=F_UI_B)
        self.diag_text.tag_config("muted", foreground=MUTED)
        self.diag_text.tag_config("fix", foreground=TEXT)

    def _diag_log(self, line, tag="muted"):
        self.diag_text.config(state="normal")
        self.diag_text.insert("end", line + "\n", tag)
        self.diag_text.see("end")
        self.diag_text.config(state="disabled")
        self.root.update_idletasks()

    def _run_diag(self):
        threading.Thread(target=self._diag_worker, daemon=True).start()

    def _diag_worker(self):
        import socket, time as _t
        ip = self.gw_ip.get().strip()
        port = self.gw_port.get().strip()
        user = self.gw_user.get().strip()
        pwd = self.gw_pass.get().strip()
        base = f"http://{ip}:{port}"

        self.diag_text.config(state="normal")
        self.diag_text.delete("1.0", "end")
        self.diag_text.config(state="disabled")
        self.diag_summary.config(text="检测中…", fg=WARN)

        passed = 0; failed = 0
        problems = []

        # 步骤1：配置完整性
        self._diag_log("【1/5】检查网关配置", "head")
        miss = []
        if not ip: miss.append("IP")
        if not port: miss.append("端口")
        if not user: miss.append("账号")
        if not pwd: miss.append("密码")
        if miss:
            self._diag_log(f"  ✗ 缺少：{', '.join(miss)}", "fail")
            self._diag_log("  → 解决方案：在顶部网关配置栏填写完整后再测。", "fix")
            failed += 1; problems.append("配置不完整")
        else:
            self._diag_log(f"  ✓ IP={ip}  端口={port}  账号={user}", "ok")
            passed += 1

        # 步骤2：TCP 连通性
        self._diag_log("", "muted")
        self._diag_log("【2/5】网络连通性（TCP 连接）", "head")
        t0 = _t.time()
        try:
            sock = socket.create_connection((ip, int(port)), timeout=3)
            sock.close()
            ms = int((_t.time() - t0) * 1000)
            self._diag_log(f"  ✓ 连上 {ip}:{port}  延迟 {ms} ms", "ok")
            passed += 1
        except Exception as e:
            self._diag_log(f"  ✗ 连不上：{e}", "fail")
            self._diag_log("  → 解决方案：", "fix")
            self._diag_log("     · 确认手机和电脑连在同一个 WiFi / 局域网", "fix")
            self._diag_log("     · 确认手机上的短信网关 App 正在运行、服务已开启", "fix")
            self._diag_log("     · 确认 App 里显示的 IP 和端口与上方一致（手机换网后 IP 会变）", "fix")
            self._diag_log("     · 关掉手机热点/电脑 VPN，或检查防火墙是否拦了端口", "fix")
            failed += 1; problems.append("网络不通")

        # 步骤3：认证
        self._diag_log("", "muted")
        self._diag_log("【3/5】账号认证", "head")
        try:
            r = requests.get(f"{base}/messages", auth=(user, pwd), timeout=5)
            if r.status_code == 401:
                self._diag_log("  ✗ 401 认证失败：账号或密码错误", "fail")
                self._diag_log("  → 解决方案：核对网关 App 里的账号密码，区分大小写。", "fix")
                failed += 1; problems.append("认证失败")
            elif r.status_code == 200:
                self._diag_log(f"  ✓ 认证通过（/messages 返回 200）", "ok")
                passed += 1
            else:
                self._diag_log(f"  ? 状态码 {r.status_code}（认证通过但接口行为异常）", "warn")
                passed += 1
        except Exception as e:
            self._diag_log(f"  ✗ 请求失败：{str(e)[:80]}", "fail")
            self._diag_log("  → 网络不通时跳过此项，先解决第 2 步。", "fix")
            failed += 1; problems.append("认证请求失败")

        # 步骤4：发送接口
        self._diag_log("", "muted")
        self._diag_log("【4/5】发送短信接口（POST /message）", "head")
        try:
            r = requests.post(f"{base}/message", auth=(user, pwd),
                              json={"textMessage": {"text": ""}, "phoneNumbers": []},
                              timeout=5)
            if r.status_code in (200, 202):
                self._diag_log("  ✓ 发送接口可用", "ok")
                passed += 1
            elif r.status_code == 401:
                self._diag_log("  ✗ 401：发送接口认证失败", "fail")
                self._diag_log("  → 解决方案：账号密码错误，回到第 3 步。", "fix")
                failed += 1; problems.append("发送接口认证失败")
            elif r.status_code == 404:
                self._diag_log("  ✗ 404：接口路径 /message 不存在", "fail")
                self._diag_log("  → 解决方案：网关 App 版本过旧或路径不同，尝试升级 App，或联系 App 作者确认接口路径。", "fix")
                failed += 1; problems.append("发送接口404")
            elif r.status_code == 405:
                self._diag_log("  ? 405：方法不允许（接口存在但不接受 POST）", "warn")
                passed += 1
            else:
                self._diag_log(f"  ? 状态码 {r.status_code}（接口存在，返回非预期）", "warn")
                passed += 1
        except Exception as e:
            self._diag_log(f"  ✗ 请求失败：{str(e)[:80]}", "fail")
            failed += 1; problems.append("发送接口请求失败")

        # 步骤5：接收接口
        self._diag_log("", "muted")
        self._diag_log("【5/5】接收短信接口", "head")
        recv_paths = ["/messages", "/api/messages", "/message/inbox", "/api/messages/inbox"]
        ok_path = None
        for p in recv_paths:
            try:
                r = requests.get(f"{base}{p}", auth=(user, pwd), timeout=4)
                if r.status_code == 200:
                    ok_path = p; break
            except Exception:
                continue
        if ok_path:
            self._diag_log(f"  ✓ 接收接口可用（{ok_path}）", "ok")
            passed += 1
        else:
            self._diag_log("  ✗ 所有候选接收接口均不可用", "fail")
            self._diag_log("  → 解决方案：", "fix")
            self._diag_log("     · 确认网关 App 已开启「读取短信 / 通知监听」权限", "fix")
            self._diag_log("     · 确认 App 没有被系统省电策略杀掉（加后台白名单）", "fix")
            self._diag_log("     · 发送功能可能仍正常，只是收不到回复", "fix")
            failed += 1; problems.append("接收接口不可用")

        # 总结
        self._diag_log("", "muted")
        self._diag_log("━━━━━━━━━━━━━━━━━━━━━━━━", "muted")
        if failed == 0:
            self._diag_summary.config(text=f"全部通过 ✓  共 {passed} 项", fg=SUCCESS)
            self._diag_log(f"总结：所有 {passed} 项检测通过，可以正常发送和接收。", "ok")
        else:
            self.diag_summary.config(text=f"发现 {failed} 个问题", fg=DANGER)
            self._diag_log(f"总结：{passed} 项通过，{failed} 项失败：{'; '.join(problems)}", "fail")
            self._diag_log("建议：按上面每步的「→ 解决方案」逐项排查。", "fix")

    # ---------- 状态 ----------
    def _set_status(self, text, color=MUTED, dot=None, label=None):
        self.status.config(text=text, fg=color)
        if dot:
            self.dot.set_color(dot)
        if label:
            self.dot_label.config(text=label, fg=color)

    # ---------- 发送单条 ----------
    def _post_sms(self, phone, text):
        base = f"http://{self.gw_ip.get().strip()}:{self.gw_port.get().strip()}"
        u = self.gw_user.get().strip()
        p = self.gw_pass.get().strip()
        ph = phone if phone.startswith("+") else "+86" + phone
        for attempt in range(RETRY_TIMES):
            try:
                r = requests.post(f"{base}/message", auth=(u, p),
                                  json={"textMessage": {"text": text},
                                        "phoneNumbers": [ph]},
                                  timeout=10)
                if r.status_code in (200, 202):
                    return True, ""
                err = f"HTTP {r.status_code}"
            except Exception as e:
                err = str(e)[:60]
            if attempt < RETRY_TIMES - 1:
                import time; time.sleep(RETRY_WAIT)
        return False, err

    def _send_single(self):
        if not self.connection_ok:
            self.single_result.config(text="请先点「测试连接」", fg=DANGER)
            return
        phone = self.single_phone.get().strip().replace(" ", "")
        text = self.single_msg.get("1.0", "end").strip()
        if not phone:
            messagebox.showwarning("提示", "请输入手机号")
            return
        if not text:
            messagebox.showwarning("提示", "请输入消息内容")
            return
        self.single_result.config(text="发送中…", fg=WARN)
        self.single_btn.set_color(MUTED)

        def work():
            ok, err = self._post_sms(phone, text)
            name = ""
            self.root.after(0, lambda: self._after_single(phone, text, ok, err))
        threading.Thread(target=work, daemon=True).start()

    def _after_single(self, phone, text, ok, err):
        self.single_btn.set_color(ACCENT)
        if ok:
            self.single_result.config(text="✓ 发送成功", fg=SUCCESS)
            self.db.upsert_sent(phone)
            self._set_status(f"● 已发送 {phone}", SUCCESS, SUCCESS, "就绪")
        else:
            self.single_result.config(text=f"✗ {err}", fg=DANGER)
            self._set_status(f"● 发送失败 {err}", DANGER, DANGER, "失败")

    # ---------- 批量发送 ----------
    def _toggle_batch(self):
        if not self.connection_ok:
            self.batch_count.config(text="请先点「测试连接」", fg=DANGER)
            return
        if self.batch_sending:
            self._stop_evt.set()
            return
        if not self.excel_path.get():
            messagebox.showwarning("提示", "请先选择 Excel 文件")
            return
        self.batch_sending = True
        self._stop_evt.clear()
        self.ok_n = self.fail_n = self.total_n = 0
        self.batch_btn.set_text("⏸  停止")
        self.batch_btn.set_color(DANGER)
        threading.Thread(target=self._batch_worker, daemon=True).start()

    def _batch_worker(self):
        try:
            df = pd.read_excel(self.excel_path.get(), dtype=str)
        except Exception as e:
            self.root.after(0, self.batch_count.config, {"text": f"读取失败：{e}"})
            self.root.after(0, self._batch_done)
            return

        # 用用户选择的电话列
        phone_col = self.phone_col_var.get().strip()
        if not phone_col or phone_col not in [str(c) for c in df.columns]:
            self.root.after(0, self.batch_count.config, {"text": "请先在上方选择电话列"})
            self.root.after(0, self._batch_done)
            return

        df = df.dropna(subset=[phone_col])
        df[phone_col] = df[phone_col].astype(str).str.replace(r"\D", "", regex=True)
        tpl = self.batch_tpl.get("1.0", "end").strip()
        total = len(df)
        self.total_n = total
        self.root.after(0, self.batch_count.config, {"text": f"共 {total} 人"})

        for i, (_, row) in enumerate(df.iterrows(), 1):
            if self._stop_evt.is_set():
                break
            phone = str(row[phone_col])
            text = self._render_tpl(tpl, row)
            name = str(row.get("姓名", "")) if "姓名" in row else ""
            ok, err = self._post_sms(phone, text)
            if ok:
                self.ok_n += 1
                self.db.upsert_sent(phone, name)
            else:
                self.fail_n += 1
            pct = i / total * 100
            self.root.after(0, self._draw_progress, pct)
            self.root.after(0, self.batch_count.config,
                            {"text": f"{i}/{total}  成功 {self.ok_n}  失败 {self.fail_n}"})
            import time; time.sleep(WAIT_BETWEEN)

        self.root.after(0, self._batch_done)

    def _draw_progress(self, pct):
        self.batch_progress.delete("all")
        w = self.batch_progress.winfo_width() or 800
        self.batch_progress.create_rectangle(0, 0, w, 8, fill=PANEL, outline="")
        fw = max(4, int(w * pct / 100))
        self.batch_progress.create_rectangle(0, 0, fw, 8, fill=ACCENT, outline="")

    def _batch_done(self):
        self.batch_sending = False
        self._stop_evt.clear()
        self.batch_btn.set_text("▶  开始批量发送")
        self.batch_btn.set_color(ACCENT)
        self._set_status(f"● 批量完成 成功 {self.ok_n} 失败 {self.fail_n}",
                         SUCCESS, SUCCESS, "完成")

    # ---------- 接收 ----------
    def _load_seen_async(self):
        def work():
            csv = os.path.join(self.recv_dir.get(), "回复记录.csv")
            try:
                if os.path.exists(csv):
                    df = pd.read_csv(csv, dtype=str)
                    if "消息ID" in df.columns:
                        self.seen_ids.update(df["消息ID"].dropna().tolist())
            except Exception:
                pass
        threading.Thread(target=work, daemon=True).start()

    def _toggle_listen(self):
        if self.listening:
            self.listening = False
            self._recv_evt.set()
            self.recv_btn.set_text("▶  开始监听")
            self.recv_btn.set_color(ACCENT)
            self._append_recv("■ 正在停止监听…", "warn")
            return
        self._listen_gen += 1
        gen = self._listen_gen
        self.listening = True
        self._recv_evt.clear()
        self.recv_btn.set_text("⏸  停止监听")
        self.recv_btn.set_color(DANGER)
        self._append_recv("──── 开始监听 ────", "dim")
        threading.Thread(target=self._recv_loop, args=(gen,), daemon=True).start()

    def _recv_loop(self, gen):
        base = f"http://{self.gw_ip.get().strip()}:{self.gw_port.get().strip()}"
        u = self.gw_user.get().strip()
        p = self.gw_pass.get().strip()

        data = self._fetch_messages(base, u, p, gen)
        if self._listen_gen != gen:
            return
        if data is None:
            self.root.after(0, self._append_recv, "✗ 连不上网关", "fail")
            self.root.after(0, self._recv_stopped)
            return
        self.root.after(0, self._append_recv, f"✓ 已连接，每 {POLL_INTERVAL} 秒轮询", "ok")

        while not self._recv_evt.is_set():
            data = self._fetch_messages(base, u, p, gen)
            if self._listen_gen != gen:
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
                    self._save_recv(sender, text)
                    self.db.upsert_recv(sender)
                    self.reply_n += 1
                    self.root.after(0, self._on_recv, sender, text)
            waited = 0
            while waited < POLL_INTERVAL and not self._recv_evt.is_set():
                self._recv_evt.wait(1)
                waited += 1

        self.root.after(0, self._append_recv, "■ 已停止监听", "dim")
        self.root.after(0, self._recv_stopped)

    def _fetch_messages(self, base, u, p, gen):
        for ep in ("/messages", "/api/messages", "/message/inbox", "/api/messages/inbox"):
            if self._recv_evt.is_set() or self._listen_gen != gen:
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

    def _save_recv(self, sender, text):
        try:
            folder = self.recv_dir.get().strip() or app_dir()
            csv = os.path.join(folder, "回复记录.csv")
            row = {"发送人": sender, "内容": text,
                   "时间": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
            new = pd.DataFrame([row])
            if os.path.exists(csv):
                old = pd.read_csv(csv, dtype=str)
                new = pd.concat([old, new], ignore_index=True)
            new.to_csv(csv, index=False, encoding="utf-8-sig")
        except Exception:
            pass

    def _on_recv(self, sender, text):
        self.recv_count.config(text=f"本次已收 {self.reply_n} 条")
        self._append_recv(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] {sender}\n  {text}", "ok")

    def _append_recv(self, m, tag="info"):
        self.recv_text.configure(state="normal")
        self.recv_text.insert("end", m + "\n\n", tag)
        self.recv_text.see("end")
        self.recv_text.configure(state="disabled")

    def _recv_stopped(self):
        self._recv_evt.clear()
        if not self.listening:
            self._set_status("● 就绪", MUTED, MUTED, "就绪")


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
