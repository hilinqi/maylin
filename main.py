"""
main.py — 代账客户收费管理系统 GUI
====================================
桌面应用
登录 → 3 个标签页（客户管理 / 按次业务 / 收入统计）
功能：搜索、分页、批量操作、月收周期、临时客户、逾期统计栏、Toast提示、Excel导出
v2: 客户状态（有效/中断/失联）、收费期间、按次业务多选
"""

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
from datetime import datetime
from typing import Optional
import os
import sys
import subprocess
import webbrowser

from database import (
    init_db, admin_exists, register_admin, login_admin,
    add_client, update_client, delete_client, get_clients,
    record_payment, calc_next_due, detect_billing_type_from_period,
    calc_overdue_months, get_overdue_level, get_billing_months,
    add_service, mark_service_done, delete_service, get_services,
    get_monthly_stats, get_year_trend, get_total_arrears,
    search_clients, get_clients_paginated,
    batch_delete_clients, batch_record_payment,
    get_overdue_summary, export_stats_to_excel,
    DB_PATH,
)

# ============================================================
#  全局样式
# ============================================================

FONT_TITLE = ("Microsoft YaHei", 14, "bold")
FONT_NORMAL = ("Microsoft YaHei", 11)
FONT_SMALL = ("Microsoft YaHei", 9)
FONT_CARD = ("Microsoft YaHei", 12, "bold")
FONT_PAGE = ("Microsoft YaHei", 10)

COLOR_PRIMARY = "#1976D2"
COLOR_SUCCESS = "#2E7D32"
COLOR_WARNING = "#D4A017"
COLOR_DANGER = "#D93025"
COLOR_BG = "#F5F5F5"

PAGE_SIZE = 20  # 每页显示条数


# ============================================================
#  Toast 悬浮提示
# ============================================================

class Toast:
    """顶部悬浮提示，3秒自动消失"""
    COLORS = {
        "info": ("#1976D2", "white"),
        "success": ("#2E7D32", "white"),
        "warning": ("#FF9800", "white"),
        "error": ("#D93025", "white"),
    }

    @staticmethod
    def show(parent, message, type="info"):
        bg, fg = Toast.COLORS.get(type, Toast.COLORS["info"])
        toast = tk.Toplevel(parent)
        toast.overrideredirect(True)
        toast.attributes("-topmost", True)

        parent.update_idletasks()
        pw = parent.winfo_width()
        px = parent.winfo_rootx()
        py = parent.winfo_rooty()
        tw = min(len(message) * 18 + 40, pw - 20)
        toast.geometry(f"{tw}x38+{px + (pw - tw) // 2}+{py + 5}")
        toast.configure(bg=bg)

        inner = tk.Frame(toast, bg=bg)
        inner.pack(fill="both", expand=True)
        tk.Label(inner, text=message, font=FONT_SMALL, bg=bg, fg=fg,
                 wraplength=tw - 20).pack(expand=True)

        toast.after(3000, toast.destroy)
        toast.bind("<Button-1>", lambda e: toast.destroy())
        inner.bind("<Button-1>", lambda e: toast.destroy())


# ============================================================
#  登录/注册窗口
# ============================================================

class LoginWindow:
    """管理员登录 / 首次注册"""

    def __init__(self):
        self.window = tk.Tk()
        self.window.title("代账客户收费管理系统 — 登录")
        self.window.geometry("420x340")
        self.window.resizable(False, False)
        self.window.configure(bg="#FFFFFF")

        self._build_ui()
        self.window.mainloop()

    def _build_ui(self):
        tk.Label(self.window, text="📊 代账客户收费管理", font=FONT_TITLE,
                 bg="#FFFFFF", fg=COLOR_PRIMARY).pack(pady=(30, 5))
        tk.Label(self.window, text="客户收费 · 逾期提醒 · 收入统计",
                 font=FONT_SMALL, bg="#FFFFFF", fg="#666").pack()

        frame = tk.Frame(self.window, bg="#FFFFFF")
        frame.pack(pady=20)

        tk.Label(frame, text="用户名", font=FONT_NORMAL, bg="#FFFFFF").grid(
            row=0, column=0, sticky="w", pady=(10, 2))
        self.entry_user = tk.Entry(frame, font=FONT_NORMAL, width=25)
        self.entry_user.grid(row=1, column=0, ipady=4)

        tk.Label(frame, text="密  码", font=FONT_NORMAL, bg="#FFFFFF").grid(
            row=2, column=0, sticky="w", pady=(10, 2))
        self.entry_pass = tk.Entry(frame, font=FONT_NORMAL, width=25, show="●")
        self.entry_pass.grid(row=3, column=0, ipady=4)
        self.entry_pass.bind("<Return>", lambda e: self._do_login())

        btn_frame = tk.Frame(self.window, bg="#FFFFFF")
        btn_frame.pack(pady=15)

        self.btn_login = tk.Button(btn_frame, text="登  录", font=FONT_NORMAL,
                                    bg=COLOR_PRIMARY, fg="white", width=12,
                                    command=self._do_login, cursor="hand2")
        self.btn_login.pack(side="left", padx=5)

        is_first = not admin_exists()
        self.btn_register = tk.Button(btn_frame,
                                       text="注册管理员" if is_first else "重置密码",
                                       font=FONT_NORMAL, width=12,
                                       command=self._do_register, cursor="hand2")
        if is_first:
            self.btn_register.configure(bg="#FF9800", fg="white")
        self.btn_register.pack(side="left", padx=5)

        hint = "首次使用请先注册管理员" if is_first else "请输入管理员账号密码"
        tk.Label(self.window, text=hint, font=FONT_SMALL,
                 bg="#FFFFFF", fg="#999").pack()

    def _do_login(self):
        username = self.entry_user.get().strip()
        password = self.entry_pass.get()
        ok, msg = login_admin(username, password)
        if ok:
            self.window.destroy()
            MainWindow()
        else:
            messagebox.showerror("登录失败", msg)

    def _do_register(self):
        if admin_exists():
            if not messagebox.askyesno("重置密码",
                    "重置将删除现有管理员数据。\n确定重置？"):
                return
            if os.path.exists(DB_PATH):
                os.remove(DB_PATH)
            if os.path.exists("data/clients.db"):
                os.remove("data/clients.db")
            init_db()

        username = self.entry_user.get().strip()
        password = self.entry_pass.get()
        if not username:
            messagebox.showwarning("提示", "请输入用户名")
            return
        if not password:
            messagebox.showwarning("提示", "请输入密码")
            return
        ok, msg = register_admin(username, password)
        if ok:
            messagebox.showinfo("成功", f"注册成功！\n请使用 {username} 登录")
            self.entry_user.delete(0, "end")
            self.entry_pass.delete(0, "end")
        else:
            messagebox.showerror("注册失败", msg)


# ============================================================
#  主窗口
# ============================================================

class MainWindow:
    """主界面 —— 3 个标签页：客户收费管理 / 按次收费业务 / 收入统计"""

    def __init__(self):
        self.window = tk.Tk()
        self.window.title("代账客户收费管理系统")
        self.window.geometry("1250x750")
        self.window.minsize(1000, 600)
        self.window.configure(bg=COLOR_BG)

        style = ttk.Style()
        style.theme_use("clam")

        # 分页状态
        self.current_page = 1
        self.total_pages = 1
        self.total_count = 0
        self.search_keyword = ""
        self.selected_ids = set()  # 勾选的客户 ID 集合

        self._build_header()
        self._build_tabs()
        self.window.mainloop()

    # ============================================================
    #  头部
    # ============================================================

    def _build_header(self):
        header = tk.Frame(self.window, bg=COLOR_PRIMARY, height=45)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(header, text="📊 代账客户收费管理系统",
                 font=FONT_TITLE, bg=COLOR_PRIMARY, fg="white").pack(
            side="left", padx=15, pady=8)
        tk.Label(header, text=datetime.now().strftime("%Y年%m月%d日"),
                 font=FONT_SMALL, bg=COLOR_PRIMARY, fg="#BBDEFB").pack(
            side="right", padx=15)

    # ============================================================
    #  标签页容器
    # ============================================================

    def _build_tabs(self):
        self.notebook = ttk.Notebook(self.window)
        self.notebook.pack(fill="both", expand=True, padx=5, pady=(0, 5))

        self.tab_clients = tk.Frame(self.notebook, bg=COLOR_BG)
        self.tab_services = tk.Frame(self.notebook, bg=COLOR_BG)
        self.tab_stats = tk.Frame(self.notebook, bg=COLOR_BG)

        self.notebook.add(self.tab_clients, text="  🏠 客户收费管理  ")
        self.notebook.add(self.tab_services, text="  📋 按次收费业务  ")
        self.notebook.add(self.tab_stats, text="  📊 收入统计  ")

        self._build_clients_tab()
        self._build_services_tab()
        self._build_stats_tab()

        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_change)

    def _on_tab_change(self, event):
        tab_idx = self.notebook.index(self.notebook.select())
        if tab_idx == 0:
            self._refresh_clients()
        elif tab_idx == 1:
            self._refresh_services()
            self.window.update_idletasks()  # 强制渲染树内容
        elif tab_idx == 2:
            self._refresh_stats()

    # ============================================================
    #  标签页 1：客户收费管理（搜索+双筛选+分页+批量+逾期统计）
    # ============================================================

    def _build_clients_tab(self):
        # ——— 搜索栏 ———
        search_frame = tk.Frame(self.tab_clients, bg=COLOR_BG)
        search_frame.pack(fill="x", padx=10, pady=(10, 3))

        tk.Label(search_frame, text="🔍", font=FONT_NORMAL, bg=COLOR_BG).pack(side="left")
        self.search_var = tk.StringVar()
        self.search_entry = tk.Entry(search_frame, textvariable=self.search_var,
                                      font=FONT_NORMAL, width=28)
        self.search_entry.pack(side="left", padx=5, ipady=2)
        self.search_entry.bind("<Return>", lambda e: self._do_search())
        self._setup_placeholder(self.search_entry, "请输入客户名称 / 备注关键词搜索")

        tk.Button(search_frame, text="搜索", font=FONT_SMALL,
                  bg=COLOR_PRIMARY, fg="white", padx=12, pady=2,
                  command=self._do_search, cursor="hand2").pack(side="left", padx=5)
        tk.Button(search_frame, text="清除", font=FONT_SMALL,
                  bg="#607D8B", fg="white", padx=8, pady=2,
                  command=self._clear_search, cursor="hand2").pack(side="left")

        # ——— 筛选栏 1：收费周期 ———
        filter_frame = tk.Frame(self.tab_clients, bg=COLOR_BG)
        filter_frame.pack(fill="x", padx=10, pady=2)

        tk.Label(filter_frame, text="周期:", font=FONT_SMALL, bg=COLOR_BG, fg="#666").pack(side="left", padx=(0, 5))
        self.filter_var = tk.StringVar(value="全部")
        cycle_filters = [
            ("📋 全部客户", "全部"),
            ("📅 年收客户", "年收"),
            ("📆 季收客户", "季收"),
            ("📌 月收客户", "月收"),
        ]
        for text, val in cycle_filters:
            tk.Radiobutton(filter_frame, text=text, variable=self.filter_var,
                           value=val, font=FONT_SMALL, bg=COLOR_BG,
                           command=self._on_filter_change,
                           indicatoron=0, padx=8, pady=3,
                           selectcolor="#BBDEFB").pack(side="left", padx=1)

        # ——— 筛选栏 2：客户状态 ———
        status_filter_frame = tk.Frame(self.tab_clients, bg=COLOR_BG)
        status_filter_frame.pack(fill="x", padx=10, pady=2)

        tk.Label(status_filter_frame, text="状态:", font=FONT_SMALL, bg=COLOR_BG, fg="#666").pack(side="left", padx=(0, 5))
        self.status_filter_var = tk.StringVar(value="全部")
        status_filters = [
            ("📋 全部状态", "全部"),
            ("✅ 有效", "有效"),
            ("⏸️ 中断", "中断"),
            ("📴 失联", "失联"),
            ("⚠️ 逾期", "逾期"),
        ]
        for text, val in status_filters:
            tk.Radiobutton(status_filter_frame, text=text, variable=self.status_filter_var,
                           value=val, font=FONT_SMALL, bg=COLOR_BG,
                           command=self._on_filter_change,
                           indicatoron=0, padx=8, pady=3,
                           selectcolor="#BBDEFB").pack(side="left", padx=1)

        # ——— 逾期统计栏 ———
        self.overdue_bar = tk.Frame(self.tab_clients, bg="#FFF3E0")
        self.overdue_bar.pack(fill="x", padx=10, pady=2)
        self.overdue_label = tk.Label(self.overdue_bar, text="", font=FONT_SMALL,
                                       bg="#FFF3E0", fg=COLOR_DANGER)
        self.overdue_label.pack(pady=4)
        self.overdue_bar.pack_forget()  # 默认隐藏

        # ——— 批量操作栏 ———
        batch_frame = tk.Frame(self.tab_clients, bg=COLOR_BG)
        batch_frame.pack(fill="x", padx=10, pady=3)

        self.select_all_var = tk.BooleanVar(value=False)
        self.cb_select_all = tk.Checkbutton(batch_frame, text="全选当前页",
                                             variable=self.select_all_var,
                                             font=FONT_SMALL, bg=COLOR_BG,
                                             command=self._toggle_select_all)
        self.cb_select_all.pack(side="left", padx=5)

        tk.Button(batch_frame, text="🗑️ 批量删除", font=FONT_SMALL,
                  bg="#F44336", fg="white", padx=8, pady=2,
                  command=self._batch_delete, cursor="hand2").pack(side="left", padx=2)
        tk.Button(batch_frame, text="💰 批量标记收费", font=FONT_SMALL,
                  bg="#FF9800", fg="white", padx=8, pady=2,
                  command=self._batch_mark_paid, cursor="hand2").pack(side="left", padx=2)

        self.loading_label = tk.Label(batch_frame, text="", font=FONT_SMALL,
                                       bg=COLOR_BG, fg="#999")

        # ——— 操作按钮 ———
        btn_frame = tk.Frame(self.tab_clients, bg=COLOR_BG)
        btn_frame.pack(fill="x", padx=10, pady=3)

        buttons = [
            ("➕ 新增客户", self._add_client_dialog, COLOR_PRIMARY),
            ("💰 记录收款", self._record_payment_dialog, "#FF9800"),
            ("✏️ 编辑客户", self._edit_client_dialog, "#4CAF50"),
            ("🗑️ 删除客户", self._delete_client, "#F44336"),
            ("🔄 刷新列表", self._refresh_clients, "#607D8B"),
        ]
        for text, cmd, color in buttons:
            tk.Button(btn_frame, text=text, font=FONT_SMALL, command=cmd,
                      bg=color, fg="white", padx=8, pady=2,
                      cursor="hand2").pack(side="left", padx=2)

        # ——— 表格区域 ———
        tree_frame = tk.Frame(self.tab_clients, bg=COLOR_BG)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=3)

        columns = ("select", "id", "name", "status", "billing_type", "fee_amount",
                   "payment_status", "charge_period", "next_due", "arrears", "overdue_status", "notes")
        self.client_tree = ttk.Treeview(tree_frame, columns=columns, show="headings",
                                         selectmode="browse", height=12)

        headers = [
            ("☐", 35), ("ID", 35), ("客户名称", 110), ("状态", 50),
            ("收费周期", 55), ("应收金额(元)", 85), ("收款状态", 65),
            ("收费期间", 195),
            ("下次应付期", 80), ("欠费金额", 80), ("逾期状态", 95), ("备注", 400),
        ]
        centers = {"select", "id", "status", "billing_type", "fee_amount",
                   "payment_status", "next_due", "arrears", "overdue_status"}
        for col, (text, width) in zip(columns, headers):
            self.client_tree.heading(col, text=text)
            anchor = "center" if col in centers else "w"
            self.client_tree.column(col, width=width, anchor=anchor, minwidth=width)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.client_tree.yview)
        self.client_tree.configure(yscrollcommand=vsb.set)
        self.client_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)

        self.client_tree.bind("<ButtonRelease-1>", self._on_tree_click)
        self.client_tree.bind("<Double-1>", self._on_client_double_click)

        # 颜色标签
        self.client_tree.tag_configure("green", background="#E8F5E9")
        self.client_tree.tag_configure("yellow", background="#FFFDE7")
        self.client_tree.tag_configure("orange", background="#FFF3E0")
        self.client_tree.tag_configure("red", background="#FFEBEE")
        self.client_tree.tag_configure("darkred", background="#FCE4EC", foreground="#8B0000")
        self.client_tree.tag_configure("unpaid", background="#FFF3E0", foreground="#D93025")
        self.client_tree.tag_configure("interrupted", background="#E3F2FD", foreground="#1565C0")
        self.client_tree.tag_configure("lost", background="#FFEBEE", foreground="#C62828")

        # ——— 分页栏 ———
        self.page_frame = tk.Frame(self.tab_clients, bg=COLOR_BG)
        self.page_frame.pack(fill="x", padx=10, pady=(0, 5))

        self._build_pagination()

        self._refresh_clients()

    def _setup_placeholder(self, entry, placeholder):
        """为输入框设置占位文本"""
        def on_focus_in(e):
            if entry.get() == placeholder:
                entry.delete(0, "end")
                entry.configure(fg="black")
        def on_focus_out(e):
            if not entry.get():
                entry.insert(0, placeholder)
                entry.configure(fg="#999")
        if not entry.get():
            entry.insert(0, placeholder)
            entry.configure(fg="#999")
        entry.bind("<FocusIn>", on_focus_in)
        entry.bind("<FocusOut>", on_focus_out)

    def _build_pagination(self):
        """构建分页按钮"""
        for w in self.page_frame.winfo_children():
            w.destroy()

        self.page_info_label = tk.Label(self.page_frame, text="",
                                         font=FONT_SMALL, bg=COLOR_BG, fg="#666")
        self.page_info_label.pack(side="left", padx=5)

        btn_frame = tk.Frame(self.page_frame, bg=COLOR_BG)
        btn_frame.pack(side="right")

        page_buttons = [
            ("|< 首页", 1),
            ("< 上一页", max(1, self.current_page - 1)),
        ]
        for text, page in page_buttons:
            state = "normal" if self.current_page > 1 else "disabled"
            tk.Button(btn_frame, text=text, font=FONT_SMALL, padx=5,
                      command=lambda p=page: self._go_page(p),
                      state=state, cursor="hand2").pack(side="left", padx=1)

        self.page_num_var = tk.StringVar(value=str(self.current_page))
        page_entry = tk.Entry(btn_frame, textvariable=self.page_num_var,
                               font=FONT_SMALL, width=4, justify="center")
        page_entry.pack(side="left", padx=3)
        page_entry.bind("<Return>", lambda e: self._go_page(
            int(self.page_num_var.get()) if self.page_num_var.get().isdigit() else 1))

        tk.Label(btn_frame, text=f"/ {self.total_pages} 页",
                 font=FONT_SMALL, bg=COLOR_BG, fg="#666").pack(side="left")

        next_buttons = [
            ("下一页 >", min(self.total_pages, self.current_page + 1)),
            ("末页 >|", self.total_pages),
        ]
        for text, page in next_buttons:
            state = "normal" if self.current_page < self.total_pages else "disabled"
            tk.Button(btn_frame, text=text, font=FONT_SMALL, padx=5,
                      command=lambda p=page: self._go_page(p),
                      state=state, cursor="hand2").pack(side="left", padx=1)

    def _go_page(self, page):
        if 1 <= page <= self.total_pages:
            self.current_page = page
            self._refresh_clients()

    def _do_search(self):
        """执行搜索"""
        keyword = self.search_var.get().strip()
        placeholder = "请输入客户名称 / 备注关键词搜索"
        if keyword == placeholder:
            keyword = ""
        self.search_keyword = keyword
        self.current_page = 1
        self.selected_ids.clear()
        self._refresh_clients()

    def _clear_search(self):
        """清除搜索"""
        self.search_var.set("")
        self.search_keyword = ""
        self.current_page = 1
        self._refresh_clients()

    def _on_filter_change(self):
        """筛选条件变更"""
        self.current_page = 1
        self.selected_ids.clear()
        self._refresh_clients()

    def _toggle_select_all(self):
        """全选/反选当前页"""
        if self.select_all_var.get():
            for item in self.client_tree.get_children():
                values = self.client_tree.item(item, "values")
                if values and len(values) > 1:
                    self.selected_ids.add(int(values[1]))
        else:
            for item in self.client_tree.get_children():
                values = self.client_tree.item(item, "values")
                if values and len(values) > 1:
                    self.selected_ids.discard(int(values[1]))
        self._refresh_clients()

    def _on_tree_click(self, event):
        """点击选择列切换勾选"""
        region = self.client_tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        column = self.client_tree.identify_column(event.x)
        if column != "#1":
            return
        item = self.client_tree.identify_row(event.y)
        if not item:
            return
        values = self.client_tree.item(item, "values")
        if not values or len(values) < 2:
            return
        cid = int(values[1])
        if cid in self.selected_ids:
            self.selected_ids.discard(cid)
        else:
            self.selected_ids.add(cid)
        self._refresh_clients()

    def _refresh_clients(self):
        """刷新客户列表（支持搜索+筛选+分页）"""
        try:
            self._do_refresh_clients()
        except Exception as e:
            messagebox.showerror("刷新错误", f"刷新客户列表失败：{e}")

    def _do_refresh_clients(self):
        """实际刷新逻辑"""
        for row in self.client_tree.get_children():
            self.client_tree.delete(row)

        self._show_loading(True)
        self.window.update_idletasks()

        # 确定状态筛选：如果选择了"逾期"则用逾期筛选
        billing_filter = self.filter_var.get()
        status_selection = self.status_filter_var.get()

        if status_selection == "逾期":
            billing_filter = "逾期"
            status_filter = "全部"
        else:
            status_filter = status_selection

        clients, total, pages = get_clients_paginated(
            billing_filter, self.search_keyword, status_filter,
            self.current_page, PAGE_SIZE
        )
        self.total_count = total
        self.total_pages = pages
        if self.current_page > self.total_pages:
            self.current_page = self.total_pages
            clients, total, pages = get_clients_paginated(
                billing_filter, self.search_keyword, status_filter,
                self.current_page, PAGE_SIZE
            )
            self.total_count = total
            self.total_pages = pages

        # 填充表格
        for c in clients:
            checked = "☑" if c["id"] in self.selected_ids else "☐"
            color_map = {
                "正常": "green", "逾期1季度": "yellow",
                "逾期半年": "orange", "逾期1年": "red", "久悬户": "darkred",
            }
            tag = color_map.get(c["overdue_level"], "green")

            arrears = c.get("arrears_amount", 0)

            # 收费期间显示
            charge_period_display = c.get("charge_period_display", "-")

            # 收款状态
            payment_status = c.get("payment_status", "未收")

            # 根据收款状态决定欠费金额和逾期状况的显示
            if payment_status == "已收":
                arrears_text = "正常"
                status_text = "正常"
            else:
                arrears_text = "正常" if arrears == 0 else f"¥{arrears:,.0f}"
                if c["overdue_months"] > 0:
                    status_text = f"逾期{c['overdue_months']}个月"
                else:
                    status_text = "正常"

            # 行颜色：中断=蓝、失联=红、已收=绿、未收=逾期等级色
            client_status = c.get("status", "有效")
            if client_status == "中断":
                final_tag = "interrupted"
            elif client_status == "失联":
                final_tag = "lost"
            elif payment_status == "已收":
                final_tag = "green"
            else:
                final_tag = tag

            self.client_tree.insert("", "end", values=(
                checked, c["id"], c["name"], c.get("status", "有效"),
                c["billing_type"], f"¥{c['fee_amount']:,.0f}",
                payment_status,
                charge_period_display, c["next_due"],
                arrears_text, status_text, c["notes"],
            ), tags=(final_tag,))

        self._show_loading(False)

        # 自动调整备注列宽度（根据最长备注内容）
        self._auto_width_notes_column(clients)

        # 更新全选状态
        if self.selected_ids:
            page_ids = {c["id"] for c in clients}
            self.select_all_var.set(page_ids.issubset(self.selected_ids))
        else:
            self.select_all_var.set(False)

        # 更新逾期统计栏
        if status_selection == "逾期":
            summary = get_overdue_summary()
            self.overdue_label.configure(
                text=f"⚠️ 当前共有 {summary['count']} 个逾期客户，总逾期金额 ¥{summary['total_arrears']:,.0f}")
            self.overdue_bar.pack(fill="x", padx=10, pady=2,
                                   before=self.client_tree.master)
        else:
            self.overdue_bar.pack_forget()

        # 更新分页信息
        self.page_info_label.configure(
            text=f"共 {self.total_count} 条记录，第 {self.current_page}/{self.total_pages} 页"
        )
        self._build_pagination()

    def _show_loading(self, show):
        """显示/隐藏加载状态"""
        if show:
            self.loading_label.configure(text="⏳ 数据加载中...")
            self.loading_label.pack(side="right", padx=10)
        else:
            self.loading_label.pack_forget()

    def _auto_width_notes_column(self, clients):
        """根据当前页所有备注内容的长度自动调整备注列宽度"""
        if not clients:
            self.client_tree.column("notes", width=200)
            return
        max_len = 0
        for c in clients:
            notes = c.get("notes", "") or ""
            # 估算像素宽度：中文字符约12px，ASCII约7px
            px = 0
            for ch in notes:
                if '\u4e00' <= ch <= '\u9fff' or '\u3000' <= ch <= '\u303f' or '\uff00' <= ch <= '\uffef':
                    px += 12
                else:
                    px += 7
            max_len = max(max_len, px)
        # 设置宽度：最小200，最大800
        width = max(200, min(max_len + 20, 800))
        self.client_tree.column("notes", width=width)

    def _on_client_double_click(self, event):
        """双击表格编辑备注"""
        region = self.client_tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        column = self.client_tree.identify_column(event.x)
        if column != "#11":  # 备注列
            return
        item = self.client_tree.selection()
        if not item:
            return
        values = self.client_tree.item(item[0], "values")
        client_id = int(values[1])
        old_notes = values[11]

        new_notes = simpledialog.askstring("编辑备注",
                                            f"客户「{values[2]}」的备注:",
                                            initialvalue=old_notes)
        if new_notes is not None:
            update_client(client_id, notes=new_notes)
            self._refresh_clients()

    def _batch_delete(self):
        """批量删除选中客户"""
        if not self.selected_ids:
            Toast.show(self.window, "请先勾选要删除的客户", "warning")
            return
        ids = list(self.selected_ids)
        client_names = []
        for iid in self.client_tree.get_children():
            vals = self.client_tree.item(iid, "values")
            if vals and int(vals[1]) in ids:
                client_names.append(vals[2])

        if not messagebox.askyesno("确认批量删除",
                                    f"将删除以下 {len(ids)} 个客户：\n" +
                                    "\n".join(client_names[:10]) +
                                    (f"\n...等共{len(ids)}个" if len(ids) > 10 else "") +
                                    "\n\n历史收费记录将保留。确认删除？"):
            return

        ok, msg = batch_delete_clients(ids)
        if ok:
            self.selected_ids.clear()
            self._refresh_clients()
            Toast.show(self.window, msg, "success")
        else:
            Toast.show(self.window, msg, "error")

    def _batch_mark_paid(self):
        """批量标记收费"""
        if not self.selected_ids:
            Toast.show(self.window, "请先勾选要标记收费的客户", "warning")
            return
        ids = list(self.selected_ids)

        paid_date = simpledialog.askstring("批量标记收费",
                                            "请输入收费日期（YYYY-MM-DD）:",
                                            initialvalue=datetime.now().strftime("%Y-%m-%d"))
        if not paid_date:
            return

        ok, msg, count = batch_record_payment(ids, paid_date=paid_date)
        if ok:
            self.selected_ids.clear()
            self._refresh_clients()
            Toast.show(self.window, msg, "success")
        else:
            Toast.show(self.window, msg, "error")

    # ============================================================
    #  客户 CRUD 操作
    # ============================================================

    def _add_client_dialog(self):
        """新增客户弹窗（收费期间 + 状态 + 实时校验）"""
        dialog = tk.Toplevel(self.window)
        dialog.title("新增客户")
        dialog.geometry("450x510")
        dialog.resizable(False, False)
        dialog.configure(bg="#FFFFFF")
        dialog.transient(self.window)
        dialog.grab_set()

        error_var = tk.StringVar()
        now = datetime.now()

        # 生成年月选项列表（前后2年）
        year_months = []
        for y in range(now.year - 2, now.year + 3):
            for m in range(1, 13):
                year_months.append(f"{y}-{m:02d}")

        # —— 客户名称 ——
        row = 0
        tk.Label(dialog, text="客户名称 *", font=FONT_NORMAL, bg="#FFFFFF").grid(
            row=row, column=0, sticky="e", padx=(20, 5), pady=6)
        name_entry = tk.Entry(dialog, font=FONT_NORMAL, width=22)
        name_entry.grid(row=row, column=1, columnspan=2, sticky="w", padx=5)
        name_entry.bind("<KeyRelease>", lambda e: self._validate_name(
            name_entry.get().strip(), error_var))

        # —— 状态 ——
        row += 1
        tk.Label(dialog, text="状态", font=FONT_NORMAL, bg="#FFFFFF").grid(
            row=row, column=0, sticky="e", padx=(20, 5), pady=6)
        status_var = tk.StringVar(value="有效")
        status_cb = ttk.Combobox(dialog, textvariable=status_var,
                                  values=["有效", "中断", "失联"],
                                  font=FONT_NORMAL, state="readonly", width=8)
        status_cb.grid(row=row, column=1, sticky="w", padx=5)

        # —— 收费周期 ——
        row += 1
        tk.Label(dialog, text="收费周期", font=FONT_NORMAL, bg="#FFFFFF").grid(
            row=row, column=0, sticky="e", padx=(20, 5), pady=6)
        billing_var = tk.StringVar(value="季收")
        billing_cb = ttk.Combobox(dialog, textvariable=billing_var,
                                   values=["季收", "年收", "月收"],
                                   font=FONT_NORMAL, state="readonly", width=8)
        billing_cb.grid(row=row, column=1, sticky="w", padx=5)
        tk.Label(dialog, text="（自动检测/手动选择）",
                 font=FONT_SMALL, bg="#FFFFFF", fg="#999").grid(
            row=row, column=2, sticky="w", padx=5)

        # —— 收费期间 ———
        row += 1
        tk.Label(dialog, text="收费期间 *", font=FONT_NORMAL, bg="#FFFFFF").grid(
            row=row, column=0, sticky="e", padx=(20, 5), pady=6)
        period_frame = tk.Frame(dialog, bg="#FFFFFF")
        period_frame.grid(row=row, column=1, columnspan=2, sticky="w", padx=5)

        start_var = tk.StringVar(value=now.strftime("%Y-%m"))
        period_start_cb = ttk.Combobox(period_frame, textvariable=start_var,
                                        values=year_months,
                                        font=FONT_NORMAL, state="readonly", width=8)
        period_start_cb.pack(side="left")

        tk.Label(period_frame, text=" 至 ", font=FONT_NORMAL, bg="#FFFFFF").pack(side="left")

        # 默认截止月（季度：当前月+2）
        end_month = now.month + 2
        end_year = now.year
        while end_month > 12:
            end_month -= 12
            end_year += 1
        end_var = tk.StringVar(value=f"{end_year}-{end_month:02d}")
        period_end_cb = ttk.Combobox(period_frame, textvariable=end_var,
                                      values=year_months,
                                      font=FONT_NORMAL, state="readonly", width=8)
        period_end_cb.pack(side="left")

        tk.Label(period_frame, text="（如：2026年5月-2026年7月）",
                 font=FONT_SMALL, bg="#FFFFFF", fg="#999").pack(side="left", padx=3)

        # 期间变更时自动检测收费周期
        def on_period_change(*args):
            try:
                detected = detect_billing_type_from_period(start_var.get(), end_var.get())
                billing_var.set(detected)
            except (ValueError, AttributeError):
                pass
        start_var.trace_add("write", on_period_change)
        end_var.trace_add("write", on_period_change)

        # —— 应收金额 ——
        row += 1
        tk.Label(dialog, text="应收金额（元）", font=FONT_NORMAL, bg="#FFFFFF").grid(
            row=row, column=0, sticky="e", padx=(20, 5), pady=6)
        fee_entry = tk.Entry(dialog, font=FONT_NORMAL, width=10)
        fee_entry.insert(0, "0")
        fee_entry.grid(row=row, column=1, sticky="w", padx=5)

        # —— 收款状态 ——
        row += 1
        tk.Label(dialog, text="收款状态", font=FONT_NORMAL, bg="#FFFFFF").grid(
            row=row, column=0, sticky="e", padx=(20, 5), pady=6)
        pay_status_var = tk.StringVar(value="未收")
        pay_status_cb = ttk.Combobox(dialog, textvariable=pay_status_var,
                                      values=["未收", "已收"],
                                      font=FONT_NORMAL, state="readonly", width=8)
        pay_status_cb.grid(row=row, column=1, sticky="w", padx=5)

        # —— 电话 ——
        row += 1
        tk.Label(dialog, text="电话", font=FONT_NORMAL, bg="#FFFFFF").grid(
            row=row, column=0, sticky="e", padx=(20, 5), pady=6)
        phone_entry = tk.Entry(dialog, font=FONT_NORMAL, width=22)
        phone_entry.grid(row=row, column=1, columnspan=2, sticky="w", padx=5)

        # —— 备注 ——
        row += 1
        tk.Label(dialog, text="备注", font=FONT_NORMAL, bg="#FFFFFF").grid(
            row=row, column=0, sticky="e", padx=(20, 5), pady=6)
        notes_entry = tk.Entry(dialog, font=FONT_NORMAL, width=22)
        notes_entry.grid(row=row, column=1, columnspan=2, sticky="w", padx=5)

        # 错误提示
        row += 1
        error_label = tk.Label(dialog, textvariable=error_var,
                                font=FONT_SMALL, bg="#FFFFFF", fg=COLOR_DANGER)
        error_label.grid(row=row, column=0, columnspan=3, pady=2)

        def save(keep_open=False):
            name = name_entry.get().strip()
            if not name:
                error_var.set("客户名称不能为空")
                return
            try:
                fee = float(fee_entry.get())
                if fee < 0:
                    error_var.set("请输入合法金额，金额不能为负数")
                    return
            except ValueError:
                error_var.set("金额格式错误")
                return

            period_start = start_var.get().strip()
            period_end = end_var.get().strip()

            if not period_start or not period_end or len(period_start) != 7 or len(period_end) != 7:
                error_var.set("请选择有效的收费期间")
                return

            # 计算 last_paid_period（起始月的前一个月）
            try:
                sy, sm = map(int, period_start.split('-'))
                sm -= 1
                if sm == 0:
                    sm = 12
                    sy -= 1
                last_paid = f"{sy}-{sm:02d}"
            except (ValueError, AttributeError):
                error_var.set("收费期间格式错误")
                return

            billing_type = billing_var.get()

            ok, msg = add_client(
                name=name,
                billing_type=billing_type,
                fee_amount=fee,
                last_paid_period=last_paid,
                phone=phone_entry.get().strip(),
                notes=notes_entry.get().strip(),
                status=status_var.get(),
                payment_status=pay_status_var.get(),
                charge_period_start=period_start,
                charge_period_end=period_end,
            )
            if ok:
                if keep_open:
                    Toast.show(self.window, msg, "success")
                    name_entry.delete(0, "end")
                    error_var.set("")
                    name_entry.focus_set()
                    self.window.after(150, self._refresh_clients)
                else:
                    dialog.destroy()
                    self._refresh_clients()
                    Toast.show(self.window, msg, "success")
            else:
                error_var.set(msg)

        row += 1
        btn_frame = tk.Frame(dialog, bg="#FFFFFF")
        btn_frame.grid(row=row, column=0, columnspan=3, pady=15)

        tk.Button(btn_frame, text="取消", font=FONT_NORMAL,
                  command=dialog.destroy, width=10,
                  cursor="hand2").pack(side="left", padx=5)

        tk.Button(btn_frame, text="保存", font=FONT_NORMAL,
                  command=lambda: save(False),
                  bg=COLOR_PRIMARY, fg="white", width=10,
                  cursor="hand2").pack(side="left", padx=5)

        tk.Button(btn_frame, text="保存并新增下一个", font=FONT_SMALL,
                  command=lambda: save(True),
                  bg="#4CAF50", fg="white", width=14,
                  cursor="hand2").pack(side="left", padx=5)

    def _validate_name(self, name, error_var):
        """实时校验客户名称"""
        if not name:
            error_var.set("")
        elif len(name) > 50:
            error_var.set("客户名称不能超过50个字符")

    def _edit_client_dialog(self):
        """编辑选中客户"""
        sel = self.client_tree.selection()
        if not sel:
            Toast.show(self.window, "请先选择一个客户", "warning")
            return

        values = self.client_tree.item(sel[0], "values")
        client_id = int(values[1])

        dialog = tk.Toplevel(self.window)
        dialog.title(f"编辑客户 — {values[2]}")
        dialog.geometry("450x480")
        dialog.resizable(False, False)
        dialog.configure(bg="#FFFFFF")
        dialog.transient(self.window)
        dialog.grab_set()

        error_var = tk.StringVar()
        now = datetime.now()
        year_months = []
        for y in range(now.year - 2, now.year + 3):
            for m in range(1, 13):
                year_months.append(f"{y}-{m:02d}")

        row = 0
        # 客户名称
        tk.Label(dialog, text="客户名称", font=FONT_NORMAL, bg="#FFFFFF").grid(
            row=row, column=0, sticky="e", padx=(20, 5), pady=6)
        name_entry = tk.Entry(dialog, font=FONT_NORMAL, width=22)
        name_entry.insert(0, values[2])
        name_entry.grid(row=row, column=1, columnspan=2, sticky="w", padx=5)

        # 状态
        row += 1
        tk.Label(dialog, text="状态", font=FONT_NORMAL, bg="#FFFFFF").grid(
            row=row, column=0, sticky="e", padx=(20, 5), pady=6)
        status_var = tk.StringVar(value=values[3])
        status_cb = ttk.Combobox(dialog, textvariable=status_var,
                                  values=["有效", "中断", "失联"],
                                  font=FONT_NORMAL, state="readonly", width=8)
        status_cb.grid(row=row, column=1, sticky="w", padx=5)

        # 收费周期
        row += 1
        tk.Label(dialog, text="收费周期", font=FONT_NORMAL, bg="#FFFFFF").grid(
            row=row, column=0, sticky="e", padx=(20, 5), pady=6)
        billing_var = tk.StringVar(value=values[4])
        billing_cb = ttk.Combobox(dialog, textvariable=billing_var,
                                   values=["季收", "年收", "月收"],
                                   font=FONT_NORMAL, state="readonly", width=8)
        billing_cb.grid(row=row, column=1, sticky="w", padx=5)

        # 收费期间
        row += 1
        tk.Label(dialog, text="收费期间", font=FONT_NORMAL, bg="#FFFFFF").grid(
            row=row, column=0, sticky="e", padx=(20, 5), pady=6)

        # 尝试从 charge_period_display 中解析
        period_display = values[7]  # 收费期间列

        period_frame = tk.Frame(dialog, bg="#FFFFFF")
        period_frame.grid(row=row, column=1, columnspan=2, sticky="w", padx=5)

        # 默认值
        default_start = now.strftime("%Y-%m")
        default_end = now.strftime("%Y-%m")

        # 尝试解析现有期间
        if period_display and period_display != "-" and "年" in period_display:
            try:
                parts = period_display.split("-")
                if len(parts) == 2:
                    s_part = parts[0].replace("年", "-").replace("月", "").strip()
                    e_part = parts[1].replace("年", "-").replace("月", "").strip()
                    # 确保是 YYYY-MM 格式
                    if len(s_part.split("-")) == 2:
                        default_start = s_part
                    if len(e_part.split("-")) == 2:
                        default_end = e_part
            except (ValueError, AttributeError):
                pass

        start_var = tk.StringVar(value=default_start)
        period_start_cb = ttk.Combobox(period_frame, textvariable=start_var,
                                        values=year_months,
                                        font=FONT_NORMAL, state="readonly", width=8)
        period_start_cb.pack(side="left")

        tk.Label(period_frame, text=" 至 ", font=FONT_NORMAL, bg="#FFFFFF").pack(side="left")

        end_var = tk.StringVar(value=default_end)
        period_end_cb = ttk.Combobox(period_frame, textvariable=end_var,
                                      values=year_months,
                                      font=FONT_NORMAL, state="readonly", width=8)
        period_end_cb.pack(side="left")

        # 应收金额
        row += 1
        tk.Label(dialog, text="应收金额（元）", font=FONT_NORMAL, bg="#FFFFFF").grid(
            row=row, column=0, sticky="e", padx=(20, 5), pady=6)
        fee_entry = tk.Entry(dialog, font=FONT_NORMAL, width=10)
        fee_entry.insert(0, values[5].replace("¥", "").replace(",", ""))
        fee_entry.grid(row=row, column=1, sticky="w", padx=5)

        # —— 收款状态 ——
        row += 1
        tk.Label(dialog, text="收款状态", font=FONT_NORMAL, bg="#FFFFFF").grid(
            row=row, column=0, sticky="e", padx=(20, 5), pady=6)
        edit_pay_status_var = tk.StringVar(value=values[6])
        edit_pay_status_cb = ttk.Combobox(dialog, textvariable=edit_pay_status_var,
                                            values=["未收", "已收"],
                                            font=FONT_NORMAL, state="readonly", width=8)
        edit_pay_status_cb.grid(row=row, column=1, sticky="w", padx=5)

        # 备注
        row += 1
        tk.Label(dialog, text="备注", font=FONT_NORMAL, bg="#FFFFFF").grid(
            row=row, column=0, sticky="e", padx=(20, 5), pady=6)
        notes_entry = tk.Entry(dialog, font=FONT_NORMAL, width=22)
        notes_entry.insert(0, values[11])
        notes_entry.grid(row=row, column=1, columnspan=2, sticky="w", padx=5)

        row += 1
        error_label = tk.Label(dialog, textvariable=error_var,
                                font=FONT_SMALL, bg="#FFFFFF", fg=COLOR_DANGER)
        error_label.grid(row=row, column=0, columnspan=3)

        # 期间变更时自动检测收费周期
        def on_period_change(*args):
            try:
                detected = detect_billing_type_from_period(start_var.get(), end_var.get())
                billing_var.set(detected)
            except (ValueError, AttributeError):
                pass
        start_var.trace_add("write", on_period_change)
        end_var.trace_add("write", on_period_change)

        def save():
            try:
                fee = float(fee_entry.get())
                if fee < 0:
                    error_var.set("请输入合法金额")
                    return
            except ValueError:
                error_var.set("金额格式错误")
                return

            period_start = start_var.get().strip()
            period_end = end_var.get().strip()

            ok, msg = update_client(
                client_id,
                name=name_entry.get().strip(),
                status=status_var.get(),
                billing_type=billing_var.get(),
                fee_amount=fee,
                payment_status=edit_pay_status_var.get(),
                notes=notes_entry.get().strip(),
                charge_period_start=period_start,
                charge_period_end=period_end,
            )
            if ok:
                dialog.destroy()
                self._refresh_clients()
                Toast.show(self.window, "客户信息更新成功", "success")
            else:
                error_var.set(msg)

        row += 1
        btn_frame = tk.Frame(dialog, bg="#FFFFFF")
        btn_frame.grid(row=row, column=0, columnspan=3, pady=15)

        tk.Button(btn_frame, text="取消", font=FONT_NORMAL,
                  command=dialog.destroy, width=10, cursor="hand2").pack(side="left", padx=5)
        tk.Button(btn_frame, text="保存修改", font=FONT_NORMAL, command=save,
                  bg=COLOR_PRIMARY, fg="white", width=10,
                  cursor="hand2").pack(side="left", padx=5)

    def _delete_client(self):
        """删除选中客户"""
        sel = self.client_tree.selection()
        if not sel:
            Toast.show(self.window, "请先选择一个客户", "warning")
            return

        values = self.client_tree.item(sel[0], "values")
        client_id = int(values[1])
        client_name = values[2]
        status = values[10]  # 逾期状态

        warning = ""
        if "逾期" in status:
            warning = "\n\n⚠️ 该客户存在逾期未结清费用！"

        if messagebox.askyesno("确认删除",
                                f"删除后客户「{client_name}」将标记为已删除，\n"
                                f"历史收费记录仍保留。是否确认删除？{warning}"):
            delete_client(client_id)
            self.selected_ids.discard(client_id)
            self._refresh_clients()
            Toast.show(self.window, "客户删除成功", "success")

    def _record_payment_dialog(self):
        """记录收款弹窗"""
        sel = self.client_tree.selection()
        if not sel:
            Toast.show(self.window, "请先选择一个客户", "warning")
            return

        values = self.client_tree.item(sel[0], "values")
        client_id = int(values[1])
        client_name = values[2]
        billing_type = values[4]
        fee_str = values[5].replace("¥", "").replace(",", "")
        last_paid = values[7]  # 收费期间
        next_due = values[8]

        # 获取实际 last_paid_period
        clients_list = get_clients("全部")
        client_info = next((c for c in clients_list if c["id"] == client_id), None)
        if client_info:
            last_paid_period = client_info["last_paid_period"]
        else:
            last_paid_period = ""

        try:
            fee_amount = float(fee_str)
        except ValueError:
            fee_amount = 0

        months_overdue = calc_overdue_months(last_paid_period, billing_type)
        level = get_overdue_level(months_overdue)

        dialog = tk.Toplevel(self.window)
        dialog.title(f"记录收款 — {client_name}")
        dialog.geometry("460x400")
        dialog.resizable(False, False)
        dialog.configure(bg="#FFFFFF")
        dialog.transient(self.window)
        dialog.grab_set()

        info_frame = tk.LabelFrame(dialog, text="客户信息", font=FONT_SMALL,
                                    bg="#FFFFFF", fg=COLOR_PRIMARY)
        info_frame.pack(fill="x", padx=15, pady=(15, 5))

        info_text = (f"收费周期: {billing_type}   标准金额: ¥{fee_amount:,.0f}\n"
                     f"最近已收期: {last_paid_period}   下次应付期: {next_due}\n"
                     f"当前状态: {level['level']}（已逾期 {months_overdue} 个月）"
                     if months_overdue > 0 else
                     f"当前状态: {level['level']}")
        tk.Label(info_frame, text=info_text, font=FONT_NORMAL, bg="#FFFFFF",
                 justify="left").pack(padx=10, pady=10)

        form_frame = tk.Frame(dialog, bg="#FFFFFF")
        form_frame.pack(fill="x", padx=15, pady=10)

        fields_form = [
            ("收款金额（元）", fee_amount),
            ("收款日期", datetime.now().strftime("%Y-%m-%d")),
            ("覆盖起始月", last_paid_period),
            ("覆盖截止月", calc_next_due(last_paid_period, billing_type)),
            ("备注", ""),
        ]
        entries_form = {}
        for i, (label, default) in enumerate(fields_form):
            tk.Label(form_frame, text=label, font=FONT_NORMAL, bg="#FFFFFF").grid(
                row=i, column=0, sticky="e", pady=5)
            entry = tk.Entry(form_frame, font=FONT_NORMAL, width=18)
            entry.insert(0, str(default))
            entry.grid(row=i, column=1, padx=10)
            entries_form[label] = entry

        def confirm():
            try:
                amount = float(entries_form["收款金额（元）"].get())
                if amount < 0:
                    Toast.show(self.window, "实收金额不能为负数", "error")
                    return
            except ValueError:
                Toast.show(self.window, "金额格式错误", "error")
                return

            paid_date = entries_form["收款日期"].get().strip()
            if paid_date > datetime.now().strftime("%Y-%m-%d"):
                Toast.show(self.window, "收费日期不能晚于当前日期", "warning")
                return

            ok, msg = record_payment(
                client_id=client_id,
                period_from=entries_form["覆盖起始月"].get().strip(),
                period_to=entries_form["覆盖截止月"].get().strip(),
                amount=amount,
                paid_date=paid_date,
                notes=entries_form["备注"].get().strip(),
            )
            if ok:
                dialog.destroy()
                self._refresh_clients()
                Toast.show(self.window, f"收款 ¥{amount:,.0f} 已记录", "success")
            else:
                Toast.show(self.window, msg, "error")

        tk.Button(dialog, text="确认收款", font=FONT_NORMAL, command=confirm,
                  bg="#FF9800", fg="white", width=15, padx=5, pady=5,
                  cursor="hand2").pack(pady=15)

    # ============================================================
    #  标签页 2：按次收费业务（按钮在上方 + 多选业务类型 + 搜索客户）
    # ============================================================

    def _build_services_tab(self):
        # ——— 筛选 + 新增按钮区域 ———
        top_frame = tk.Frame(self.tab_services, bg=COLOR_BG)
        top_frame.pack(fill="x", padx=10, pady=(10, 5))

        self.svc_filter_var = tk.StringVar(value="全部")
        for text, val in [("📋 全部", "全部"), ("⏳ 未收", "未收"), ("✅ 已收", "已收")]:
            tk.Radiobutton(top_frame, text=text, variable=self.svc_filter_var,
                           value=val, font=FONT_SMALL, bg=COLOR_BG,
                           command=self._refresh_services,
                           indicatoron=0, padx=10, pady=4,
                           selectcolor="#BBDEFB").pack(side="left", padx=2)

        # 新增业务按钮（移到列表上方）
        svc_btn_frame = tk.Frame(self.tab_services, bg=COLOR_BG)
        svc_btn_frame.pack(fill="x", padx=10, pady=(0, 3))

        svc_buttons = [
            ("➕ 新增业务", self._add_service_dialog, COLOR_PRIMARY),
            ("✅ 标记已收", self._mark_service_done, "#4CAF50"),
            ("🗑️ 删除业务", self._delete_service, "#F44336"),
            ("🔄 刷新", self._refresh_services, "#607D8B"),
        ]
        for text, cmd, color in svc_buttons:
            tk.Button(svc_btn_frame, text=text, font=FONT_SMALL, command=cmd,
                      bg=color, fg="white", padx=10, pady=3,
                      cursor="hand2").pack(side="left", padx=2)

        # 调试状态栏（显示刷新状态）
        self.svc_debug_label = tk.Label(
            svc_btn_frame, text="", font=("Microsoft YaHei", 9),
            bg=COLOR_BG, fg="#888", anchor="e"
        )
        self.svc_debug_label.pack(side="right", padx=10)

        # ——— 表格区域 ———
        tree_frame = tk.Frame(self.tab_services, bg=COLOR_BG)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=5)

        cols = ("id", "client_name", "service_type", "fee_standard",
                "actual_fee", "status", "completed_date", "notes")
        self.svc_tree = ttk.Treeview(tree_frame, columns=cols, show="headings",
                                      selectmode="browse", height=15)

        svc_headers = [
            ("ID", 40), ("客户", 130), ("业务类型", 150),
            ("业务金额(元)", 90), ("实收金额(元)", 90), ("收款状态", 65),
            ("收款/完成日期", 100), ("备注", 150),
        ]
        for col, (text, width) in zip(cols, svc_headers):
            self.svc_tree.heading(col, text=text)
            self.svc_tree.column(col, width=width, anchor="center")
        self.svc_tree.column("client_name", anchor="w")
        self.svc_tree.column("service_type", anchor="w")
        self.svc_tree.column("notes", anchor="w")

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.svc_tree.yview)
        self.svc_tree.configure(yscrollcommand=vsb.set)
        self.svc_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)

        self.svc_tree.tag_configure("paid", background="#E8F5E9")
        self.svc_tree.tag_configure("unpaid", background="#FFF3E0")

        self._refresh_services()

    def _refresh_services(self):
        try:
            self._do_refresh_services()
        except Exception as e:
            # 静默异常会导致列表空白，写入日志文件
            log_path = os.path.join(os.path.dirname(DB_PATH), "error.log")
            import traceback
            with open(log_path, "a", encoding="utf-8") as lf:
                lf.write(f"[{datetime.now()}] _refresh_services 异常: {e}\n")
                lf.write(traceback.format_exc() + "\n")

    def _do_refresh_services(self):
        for row in self.svc_tree.get_children():
            self.svc_tree.delete(row)

        filter_val = self.svc_filter_var.get()
        services = get_services(filter_val)
        self.svc_debug_label.config(
            text=f"DB: {os.path.basename(DB_PATH)} | 筛选: {filter_val} | 记录数: {len(services)}"
        )
        if not services:
            # 空列表时插入占位提示
            self.svc_tree.insert("", "end", values=(
                "-", "暂无业务记录", "", "", "", "", "", ""
            ))
        else:
            for s in services:
                tag = "paid" if s["status"] == "已收" else "unpaid"
                client_display = s["client_name"]
                if s.get("temp_customer_name") and not s.get("client_id"):
                    client_display = f"⚡临时: {s['temp_customer_name']}"

                self.svc_tree.insert("", "end", values=(
                    s["id"], client_display,
                    s.get("service_type_display", s["service_type"]),
                    f"¥{s['fee_standard']:,.0f}",
                    f"¥{s['actual_fee']:,.0f}" if s["actual_fee"] else "-",
                    s["status"],
                    s["completed_date"] if s["completed_date"] else "-",
                    s["notes"],
                ), tags=(tag,))

    def _add_service_dialog(self):
        """新增按次业务（支持搜索客户 + 手动输入 + 多选业务类型 + 收款状态）"""
        dialog = tk.Toplevel(self.window)
        dialog.title("新增按次业务")
        dialog.geometry("520x560")
        dialog.configure(bg="#FFFFFF")
        dialog.transient(self.window)
        dialog.grab_set()

        # ——— 搜索已有客户 ———
        row = 0
        tk.Label(dialog, text="🔍 搜索客户", font=FONT_NORMAL, bg="#FFFFFF").grid(
            row=row, column=0, sticky="e", padx=(20, 5), pady=(10, 3))

        search_frame = tk.Frame(dialog, bg="#FFFFFF")
        search_frame.grid(row=row, column=1, columnspan=2, sticky="w", padx=5)

        self.svc_search_var = tk.StringVar()
        search_entry = tk.Entry(search_frame, textvariable=self.svc_search_var,
                                 font=FONT_NORMAL, width=18)
        search_entry.pack(side="left", padx=(0, 5))
        search_entry.bind("<Return>", lambda e: self._do_svc_client_search())

        tk.Button(search_frame, text="搜索", font=FONT_SMALL,
                  bg=COLOR_PRIMARY, fg="white", padx=8,
                  command=self._do_svc_client_search,
                  cursor="hand2").pack(side="left")

        # 搜索结果列表
        row += 1
        tk.Label(dialog, text="选择客户", font=FONT_NORMAL, bg="#FFFFFF").grid(
            row=row, column=0, sticky="ne", padx=(20, 5), pady=3)

        list_frame = tk.Frame(dialog, bg="#FFFFFF")
        list_frame.grid(row=row, column=1, columnspan=2, sticky="w", padx=5)

        self.svc_client_listbox = tk.Listbox(list_frame, font=FONT_SMALL,
                                              width=30, height=4,
                                              exportselection=False)
        self.svc_client_listbox.pack(side="left", fill="both", expand=True)

        list_scroll = ttk.Scrollbar(list_frame, orient="vertical",
                                     command=self.svc_client_listbox.yview)
        self.svc_client_listbox.configure(yscrollcommand=list_scroll.set)
        list_scroll.pack(side="right", fill="y")

        # 双击列表项直接保存
        self.svc_client_listbox.bind("<Double-1>", lambda e: save())

        # 初始加载所有客户
        self._do_svc_client_search()

        # ——— 手动输入客户名称 ———
        row += 1
        self.svc_manual_var = tk.BooleanVar(value=False)
        tk.Checkbutton(dialog, text="手动输入新客户名称",
                       variable=self.svc_manual_var,
                       font=FONT_SMALL, bg="#FFFFFF",
                       command=lambda: self._toggle_svc_manual()).grid(
            row=row, column=0, sticky="e", padx=(20, 5), pady=3)

        self.svc_manual_entry = tk.Entry(dialog, font=FONT_NORMAL, width=22,
                                          state="disabled")
        self.svc_manual_entry.grid(row=row, column=1, columnspan=2, sticky="w", padx=5)

        # ——— 业务类型（多选） ———
        row += 1
        tk.Label(dialog, text="业务类型（多选）", font=FONT_NORMAL, bg="#FFFFFF").grid(
            row=row, column=0, sticky="ne", padx=(20, 5), pady=(10, 3))

        svc_type_frame = tk.Frame(dialog, bg="#FFFFFF")
        svc_type_frame.grid(row=row, column=1, columnspan=2, sticky="w", padx=5)

        self.svc_type_vars = {}
        svc_types = ["设立/变更", "税务处理", "商标", "医社保", "个体"]
        for i, st in enumerate(svc_types):
            var = tk.BooleanVar(value=False)
            self.svc_type_vars[st] = var
            tk.Checkbutton(svc_type_frame, text=st, variable=var,
                           font=FONT_SMALL, bg="#FFFFFF").grid(
                row=0, column=i, padx=3, sticky="w")

        # "其他" + 手动自定义输入
        other_var = tk.BooleanVar(value=False)
        self.svc_type_vars["其他"] = other_var
        tk.Checkbutton(svc_type_frame, text="其他", variable=other_var,
                       font=FONT_SMALL, bg="#FFFFFF").grid(
            row=1, column=0, padx=3, sticky="w")
        self.svc_custom_entry = tk.Entry(svc_type_frame, font=FONT_SMALL, width=18)
        self.svc_custom_entry.grid(row=1, column=1, columnspan=4, padx=3, sticky="w")
        self._setup_placeholder(self.svc_custom_entry, "手动自定义输入...")

        # ——— 当次业务金额 ———
        row += 1
        tk.Label(dialog, text="当次业务金额（元）", font=FONT_NORMAL, bg="#FFFFFF").grid(
            row=row, column=0, sticky="e", padx=(20, 5), pady=6)
        self.svc_amount_entry = tk.Entry(dialog, font=FONT_NORMAL, width=12)
        self.svc_amount_entry.insert(0, "0")
        self.svc_amount_entry.grid(row=row, column=1, sticky="w", padx=5)

        # ——— 收款状态 ———
        row += 1
        tk.Label(dialog, text="收款状态", font=FONT_NORMAL, bg="#FFFFFF").grid(
            row=row, column=0, sticky="e", padx=(20, 5), pady=6)
        status_frame = tk.Frame(dialog, bg="#FFFFFF")
        status_frame.grid(row=row, column=1, columnspan=2, sticky="w", padx=5)

        self.svc_status_var = tk.StringVar(value="未收")
        tk.Radiobutton(status_frame, text="未收", variable=self.svc_status_var,
                       value="未收", font=FONT_SMALL, bg="#FFFFFF",
                       command=self._toggle_svc_date).pack(side="left", padx=3)
        tk.Radiobutton(status_frame, text="已收", variable=self.svc_status_var,
                       value="已收", font=FONT_SMALL, bg="#FFFFFF",
                       command=self._toggle_svc_date).pack(side="left", padx=3)

        # 收款日期（已收时显示，默认隐藏）
        self.svc_date_frame = tk.Frame(dialog, bg="#FFFFFF")
        tk.Label(self.svc_date_frame, text="收款日期",
                 font=FONT_SMALL, bg="#FFFFFF", fg="#666").pack(side="left", padx=(5, 3))
        self.svc_date_entry = tk.Entry(self.svc_date_frame, font=FONT_SMALL, width=12)
        self.svc_date_entry.insert(0, datetime.now().strftime("%Y-%m-%d"))
        self.svc_date_entry.pack(side="left")
        # 放在 row+1 位置（status 之后），默认隐藏
        self.svc_date_frame.grid(row=row + 1, column=1, columnspan=2, sticky="w", padx=5, pady=3)
        self.svc_date_frame.grid_remove()

        # ——— 备注 ———
        row += 2  # 跳过日期框行
        tk.Label(dialog, text="备注", font=FONT_NORMAL, bg="#FFFFFF").grid(
            row=row, column=0, sticky="e", padx=(20, 5), pady=6)
        self.svc_notes_entry = tk.Entry(dialog, font=FONT_NORMAL, width=22)
        self.svc_notes_entry.grid(row=row, column=1, columnspan=2, sticky="w", padx=5)

        # 存储搜索结果
        self.svc_search_results = []

        def save():
            # 确定客户
            client_id = None
            temp_name = ""

            if self.svc_manual_var.get():
                temp_name = self.svc_manual_entry.get().strip()
                if not temp_name:
                    Toast.show(self.window, "请输入客户名称", "warning")
                    return
            else:
                sel = self.svc_client_listbox.curselection()
                if not sel:
                    Toast.show(self.window, "请选择客户或勾选手动输入", "warning")
                    return
                idx = sel[0]
                if idx < len(self.svc_search_results):
                    client_id = self.svc_search_results[idx]["id"]

            # 收集选中的业务类型
            selected_types = [st for st, var in self.svc_type_vars.items() if var.get()]
            if not selected_types:
                Toast.show(self.window, "请至少选择一个业务类型", "warning")
                return

            # 如果选中"其他"且有自定义输入，替换为自定义内容
            custom_text = self.svc_custom_entry.get().strip()
            if "其他" in selected_types and custom_text and custom_text != "手动自定义输入...":
                selected_types = [custom_text if st == "其他" else st for st in selected_types]

            service_type = ",".join(selected_types)

            # 金额
            try:
                amount = float(self.svc_amount_entry.get())
                if amount < 0:
                    Toast.show(self.window, "金额不能为负数", "warning")
                    return
            except ValueError:
                Toast.show(self.window, "金额格式错误", "warning")
                return

            # 状态
            status = self.svc_status_var.get()
            completed_date = ""
            actual_fee = 0
            if status == "已收":
                completed_date = self.svc_date_entry.get().strip()
                actual_fee = amount

            ok, msg = add_service(
                client_id=client_id,
                service_type=service_type,
                fee_standard=amount,
                actual_fee=actual_fee,
                temp_customer_name=temp_name,
                notes=self.svc_notes_entry.get().strip(),
                status=status,
                completed_date=completed_date,
            )
            if ok:
                dialog.destroy()
                # 延迟刷新：确保对话框完全销毁、grab释放后再更新界面
                def do_refresh():
                    self.notebook.select(self.tab_services)
                    self._refresh_services()
                    self.window.update_idletasks()
                    Toast.show(self.window, msg, "success")
                self.window.after(50, do_refresh)
            else:
                Toast.show(self.window, msg, "error")

        row += 1
        btn_frame = tk.Frame(dialog, bg="#FFFFFF")
        btn_frame.grid(row=row, column=0, columnspan=3, pady=15)

        tk.Button(btn_frame, text="取消", font=FONT_NORMAL,
                  command=dialog.destroy, width=10,
                  cursor="hand2").pack(side="left", padx=5)
        tk.Button(btn_frame, text="保 存", font=FONT_NORMAL, command=save,
                  bg=COLOR_PRIMARY, fg="white", width=12,
                  cursor="hand2").pack(side="left", padx=5)

    def _do_svc_client_search(self):
        """搜索客户并填充列表"""
        self.svc_client_listbox.delete(0, "end")
        keyword = self.svc_search_var.get().strip()

        result, _ = search_clients(keyword, "全部", "全部", 0, 50)
        self.svc_search_results = result

        for c in result:
            display = f"{c['name']}  [{c.get('status', '有效')}]  {c['billing_type']} ¥{c['fee_amount']:,.0f}"
            self.svc_client_listbox.insert("end", display)

    def _toggle_svc_manual(self):
        """切换手动输入客户名称"""
        if self.svc_manual_var.get():
            self.svc_client_listbox.configure(state="disabled")
            self.svc_manual_entry.configure(state="normal")
            self.svc_search_entry_state = "disabled"
        else:
            self.svc_client_listbox.configure(state="normal")
            self.svc_manual_entry.configure(state="disabled")

    def _toggle_svc_date(self):
        """切换收款日期显示"""
        if self.svc_status_var.get() == "已收":
            self.svc_date_frame.grid()
        else:
            self.svc_date_frame.grid_remove()

    def _mark_service_done(self):
        sel = self.svc_tree.selection()
        if not sel:
            Toast.show(self.window, "请先选择一个业务", "warning")
            return

        values = self.svc_tree.item(sel[0], "values")
        if values[5] == "已收":
            Toast.show(self.window, "该业务已标记为已收", "info")
            return

        fee_str = values[3].replace("¥", "").replace(",", "")
        try:
            default_fee = float(fee_str)
        except ValueError:
            default_fee = 0

        actual = simpledialog.askfloat("确认收款",
                                        f"业务「{values[2]}」\n实收金额（元）:",
                                        initialvalue=default_fee)
        if actual is not None:
            mark_service_done(int(values[0]), actual)
            self._refresh_services()
            Toast.show(self.window, "业务已标记完成，统计到当月业绩", "success")

    def _delete_service(self):
        sel = self.svc_tree.selection()
        if not sel:
            Toast.show(self.window, "请先选择一个业务", "warning")
            return
        values = self.svc_tree.item(sel[0], "values")
        if messagebox.askyesno("确认删除", f"确定删除业务「{values[2]}」？"):
            delete_service(int(values[0]))
            self._refresh_services()
            Toast.show(self.window, "业务已删除", "success")

    # ============================================================
    #  标签页 3：收入统计（+Excel导出）
    # ============================================================

    def _build_stats_tab(self):
        top_frame = tk.Frame(self.tab_stats, bg=COLOR_BG)
        top_frame.pack(fill="x", padx=10, pady=10)

        tk.Label(top_frame, text="统计月份:", font=FONT_NORMAL, bg=COLOR_BG).pack(side="left")
        now = datetime.now()
        self.stats_year = tk.IntVar(value=now.year)
        self.stats_month = tk.IntVar(value=now.month)

        ttk.Spinbox(top_frame, from_=2020, to=2099, textvariable=self.stats_year,
                    width=6, font=FONT_NORMAL).pack(side="left", padx=5)
        tk.Label(top_frame, text="年", font=FONT_NORMAL, bg=COLOR_BG).pack(side="left")

        ttk.Combobox(top_frame, textvariable=self.stats_month,
                     values=list(range(1, 13)), width=3,
                     font=FONT_NORMAL, state="readonly").pack(side="left", padx=5)
        tk.Label(top_frame, text="月", font=FONT_NORMAL, bg=COLOR_BG).pack(side="left")

        tk.Button(top_frame, text="查询", font=FONT_SMALL,
                  bg=COLOR_PRIMARY, fg="white", padx=12, pady=2,
                  command=self._refresh_stats,
                  cursor="hand2").pack(side="left", padx=15)

        tk.Button(top_frame, text="📥 导出 Excel", font=FONT_SMALL,
                  bg="#1565C0", fg="white", padx=10, pady=2,
                  command=self._export_excel,
                  cursor="hand2").pack(side="right", padx=10)

        # 收入卡片
        self.cards_frame = tk.Frame(self.tab_stats, bg=COLOR_BG)
        self.cards_frame.pack(fill="x", padx=10, pady=5)

        # 柱状图
        self.chart_frame = tk.Frame(self.tab_stats, bg="white", relief="groove", bd=1)
        self.chart_frame.pack(fill="x", padx=10, pady=5)
        tk.Label(self.chart_frame, text="📊 近 12 个月收入趋势",
                 font=FONT_NORMAL, bg="white", fg=COLOR_PRIMARY).pack(
            anchor="w", padx=10, pady=5)
        self.chart_canvas = tk.Canvas(self.chart_frame, height=180, bg="white")
        self.chart_canvas.pack(fill="x", padx=10, pady=(0, 10))

        # 收入明细表
        detail_frame = tk.Frame(self.tab_stats, bg=COLOR_BG)
        detail_frame.pack(fill="both", expand=True, padx=10, pady=5)

        detail_cols = ("client_name", "type", "amount", "period", "paid_date")
        self.detail_tree = ttk.Treeview(detail_frame, columns=detail_cols,
                                         show="headings", height=8)
        detail_headers = [
            ("客户/业务", 150), ("类型", 60), ("金额(元)", 90),
            ("收费期", 150), ("收款日期", 100),
        ]
        for (col, _), (text, width) in zip(detail_cols, detail_headers):
            self.detail_tree.heading(col, text=text)
            self.detail_tree.column(col, width=width, anchor="center")
        self.detail_tree.column("client_name", anchor="w")
        self.detail_tree.column("period", anchor="w")

        vsb = ttk.Scrollbar(detail_frame, orient="vertical", command=self.detail_tree.yview)
        self.detail_tree.configure(yscrollcommand=vsb.set)
        self.detail_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        detail_frame.columnconfigure(0, weight=1)
        detail_frame.rowconfigure(0, weight=1)

        self._refresh_stats()

    def _refresh_stats(self):
        year = self.stats_year.get()
        month = self.stats_month.get()
        stats = get_monthly_stats(year, month)

        for w in self.cards_frame.winfo_children():
            w.destroy()

        cards_data = [
            ("💰 本月总收入", f"¥{stats['total_collected']:,.0f}", COLOR_PRIMARY),
            ("📅 年收客户贡献", f"¥{stats['annual_contribution']:,.0f}", "#1565C0"),
            ("📆 季收客户贡献", f"¥{stats['quarterly_contribution']:,.0f}", "#00838F"),
            ("📋 按次业务贡献", f"¥{stats['service_contribution']:,.0f}", "#6A1B9A"),
            ("✅ 已收金额", f"¥{stats['total_collected']:,.0f}", COLOR_SUCCESS),
            ("⏳ 应收未收", f"¥{stats['unpaid']:,.0f}",
             COLOR_DANGER if stats['unpaid'] > 0 else COLOR_SUCCESS),
            ("📛 欠款总额", f"¥{stats.get('total_arrears', 0):,.0f}",
             COLOR_DANGER if stats.get('total_arrears', 0) > 0 else COLOR_SUCCESS),
            ("📈 收款率", f"{stats['collection_rate']}%",
             COLOR_SUCCESS if stats['collection_rate'] >= 80 else COLOR_WARNING),
        ]

        cols_per_row = 4
        for i, (label, value, color) in enumerate(cards_data):
            card = tk.Frame(self.cards_frame, bg="white", relief="groove", bd=1,
                            width=170, height=65)
            card.grid(row=i // cols_per_row, column=i % cols_per_row,
                      padx=4, pady=4, sticky="nsew")
            card.pack_propagate(False)
            tk.Label(card, text=label, font=FONT_SMALL, bg="white", fg="#888").pack(
                pady=(6, 1))
            tk.Label(card, text=value, font=FONT_CARD, bg="white", fg=color).pack()

        for i in range(cols_per_row):
            self.cards_frame.columnconfigure(i, weight=1)

        self._draw_chart(year)

        for row in self.detail_tree.get_children():
            self.detail_tree.delete(row)
        for d in stats["details"]:
            self.detail_tree.insert("", "end", values=(
                d["client_name"], d["type"], f"¥{d['amount']:,.0f}",
                d["period"], d["paid_date"],
            ))

    def _draw_chart(self, year: int):
        self.chart_canvas.delete("all")
        trend = get_year_trend(year)
        w = self.chart_canvas.winfo_width()
        if w < 100:
            w = 900
        h = 160
        pad_left, pad_right, pad_top, pad_bottom = 40, 20, 20, 30
        chart_w = w - pad_left - pad_right
        chart_h = h - pad_top - pad_bottom

        max_amount = max((t["amount"] for t in trend), default=1)
        if max_amount == 0:
            max_amount = 1

        bar_w = chart_w / 12 * 0.55
        gap = chart_w / 12

        for i, t in enumerate(trend):
            x = pad_left + i * gap + (gap - bar_w) / 2
            bar_h = (t["amount"] / max_amount) * chart_h if t["amount"] > 0 else 2
            y = h - pad_bottom - bar_h

            now = datetime.now()
            color = COLOR_PRIMARY if (year == now.year and t["month"] == now.month) else "#90CAF9"

            self.chart_canvas.create_rectangle(x, y, x + bar_w, h - pad_bottom,
                                               fill=color, outline="")
            if t["amount"] > 0:
                self.chart_canvas.create_text(x + bar_w / 2, y - 8,
                                              text=f"¥{t['amount']:,.0f}",
                                              font=("Microsoft YaHei", 7), fill="#333")
            self.chart_canvas.create_text(x + bar_w / 2, h - pad_bottom + 12,
                                          text=f"{t['month']}月",
                                          font=("Microsoft YaHei", 8), fill="#666")

        self.chart_canvas.create_line(pad_left, h - pad_bottom,
                                       pad_left + chart_w, h - pad_bottom,
                                       fill="#CCC")

    def _export_excel(self):
        """导出 Excel"""
        year = self.stats_year.get()
        month = self.stats_month.get()

        filepath = filedialog.asksaveasfilename(
            title="导出收入统计",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx"), ("CSV 文件", "*.csv")],
            initialfile=f"收入统计_{year}年{month}月.xlsx",
        )
        if not filepath:
            return

        try:
            result_path = export_stats_to_excel(year, month, filepath)
            Toast.show(self.window, f"导出成功：{os.path.basename(result_path)}", "success")
            if messagebox.askyesno("导出成功", f"文件已保存至：\n{result_path}\n\n是否打开文件夹？"):
                if sys.platform == "win32":
                    os.startfile(os.path.dirname(result_path))
                elif sys.platform == "darwin":
                    subprocess.run(["open", os.path.dirname(result_path)])
                else:
                    subprocess.run(["xdg-open", os.path.dirname(result_path)])
        except Exception as e:
            Toast.show(self.window, f"导出失败：{e}", "error")


# ============================================================
#  入口
# ============================================================

if __name__ == "__main__":
    init_db()

    if not admin_exists():
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo("欢迎",
            "首次使用，请注册管理员账号。\n\n数据将保存在软件同目录下的 data/clients.db")
        root.destroy()

    LoginWindow()
