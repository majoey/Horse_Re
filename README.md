# Horse_Re
import tkinter as tk
from tkinter import ttk, messagebox, Canvas
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
import json
import os
import logging
from tkcalendar import Calendar
from ttkthemes import ThemedTk

logging.basicConfig(level=logging.DEBUG, format="%(asctime)s - %(levelname)s - %(message)s")

class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip_window = None
        self.widget.bind("<Enter>", self.show_tip)
        self.widget.bind("<Leave>", self.hide_tip)

    def show_tip(self, event):
        if self.tip_window or not self.text:
            return
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height()
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        label = tk.Label(tw, text=self.text, justify="left", bg="#FFFFE0", relief="solid", borderwidth=1, font=("Arial", 10))
        label.pack()

    def hide_tip(self, event):
        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None
class DataManager:
    def __init__(self, data_file="xiaoma_records.json"):
        self.data_file = data_file
        self.expense_categories = {
            "生活消费": {"日常开销": ["日常饮食", "小额购", "交通", "高铁", "快递", "理发", "话费", "手续费"], "品质提升": ["餐饮", "大额购", "娱乐", "会员", "学习", "旅行"]},
            "固定支出": {"居住成本": ["房租&水电"], "长期义务": ["生活费", "团费"]},
            "特殊场景": {"健康医疗": ["治病", "体检"], "亲友往来": ["红包", "礼物"], "募资": ["本金借入", "利息支出"], "投资": ["股票", "其他项目"]}
        }
        self.income_categories = {
            "固定收入": {"工作": ["工资", "补贴", "公积金", "垫付报销"]},
            "生活收入": {"日常收入": ["返现", "高铁冲抵"]},
            "特殊场景": {"亲友往来": ["红包"], "筹资": ["本金借入"], "投资": ["股票", "其他项目"]}
        }
        self.accounts = ["银行卡", "现金", "微信/支付宝"]
        self.budgets = {"2025-03": 5000}
        self.records = []
        self.load_data()

    def load_data(self):
        try:
            if os.path.exists(self.data_file):
                with open(self.data_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if not isinstance(data, dict):
                        logging.error(f"数据文件格式错误: {data}")
                        self.records = []
                        return
                    self.records = data.get("records", [])
                    self.expense_categories = data.get("expense_categories", self.expense_categories)
                    self.income_categories = data.get("income_categories", self.income_categories)
                    logging.info("数据加载成功")
            else:
                self.records = []
                logging.info("未找到数据文件，使用默认值")
        except Exception as e:
            logging.error(f"加载数据失败: {e}")
            self.records = []

    def save_data(self):
        try:
            with open(self.data_file, "w", encoding="utf-8") as f:
                json.dump({
                    "records": self.records,
                    "expense_categories": self.expense_categories,
                    "income_categories": self.income_categories
                }, f, ensure_ascii=False, indent=4)
            logging.info("数据保存成功")
        except Exception as e:
            logging.error(f"保存数据失败: {e}")

    def add_record(self, record):
        self.records.append(record)
        self.check_budget(record)
        self.save_data()

    def edit_record(self, index, new_record):
        self.records[index] = new_record
        self.save_data()

    def delete_record(self, index):
        del self.records[index]
        self.save_data()

    def add_category(self, is_income, cat, subcat, item=None):
        cats = self.income_categories if is_income else self.expense_categories
        if cat not in cats:
            cats[cat] = {}
        if subcat not in cats[cat]:
            cats[cat][subcat] = []
        if item and item not in cats[cat][subcat]:
            cats[cat][subcat].append(item)
        self.save_data()

    def delete_category(self, is_income, cat, subcat=None, item=None):
        cats = self.income_categories if is_income else self.expense_categories

        if item is not None and subcat is not None:  # 删除小类
            if item in cats[cat][subcat]:
                cats[cat][subcat].remove(item)
                self.records = [r for r in self.records if
                                not (r["大类"] == cat and r["中类"] == subcat and r["小类"] == item)]
        elif subcat is not None:  # 删除中类
            if subcat in cats[cat]:
                del cats[cat][subcat]
                self.records = [r for r in self.records if not (r["大类"] == cat and r["中类"] == subcat)]
        else:  # 删除大类
            if cat in cats:
                del cats[cat]
                self.records = [r for r in self.records if r["大类"] != cat]

        self.save_data()

    def check_budget(self, record):
        if record["金额"] < 0:
            month = record["日期"][:7]
            if month in self.budgets:
                total = abs(sum(r["金额"] for r in self.records if r["日期"].startswith(month) and r["金额"] < 0))
                if total > self.budgets[month]:
                    messagebox.showwarning("预算超支", f"{month} 已超出预算 {self.budgets[month]} 元！")
class XiaoMaApp:
    def __init__(self, root):
        self.root = root
        self.root.title("小马记账")
        self.root.geometry("1000x600")
        self.root.configure(bg="#E0F2E9")
        self.data_manager = DataManager()
        self.style = ttk.Style()
        self.style.theme_use("radiance")
        self.style.configure("TButton", borderwidth=0, relief="flat")
        self.setup_ui()

    def setup_ui(self):
        tk.Label(self.root, text="小马记账", font=("Arial", 20, "bold"), bg="#E0F2E9", fg="#42A5F5").pack(pady=10)

        button_frame = tk.Frame(self.root, bg="#E0F2E9")
        button_frame.pack(fill="x", pady=5)
        tk.Button(button_frame, text="+ 支出", command=lambda: self.open_add_window(False), bg="#66BB6A", fg="white",
                  font=("Arial", 14, "bold"), relief="flat", bd=0).pack(side="left", padx=5)
        tk.Button(button_frame, text="+ 收入", command=lambda: self.open_add_window(True), bg="#EF5350", fg="white",
                  font=("Arial", 14, "bold"), relief="flat", bd=0).pack(side="left", padx=5)
        tk.Button(button_frame, text="编辑", command=self.edit_record, bg="#FFB300", fg="white", font=("Arial", 12),
                  relief="flat", bd=0).pack(side="left", padx=5)
        tk.Button(button_frame, text="删除", command=self.delete_record, bg="#EF5350", fg="white", font=("Arial", 12),
                  relief="flat", bd=0).pack(side="left", padx=5)
        tk.Button(button_frame, text="周期", command=self.add_periodic_record, bg="#78909C", fg="white",
                  font=("Arial", 12), relief="flat", bd=0).pack(side="left", padx=5)
        tk.Button(button_frame, text="类别", command=self.manage_categories, bg="#FF7043", fg="white",
                  font=("Arial", 12), relief="flat", bd=0).pack(side="left", padx=5)
        export_btn = tk.Button(button_frame, text="导出", command=self.export_to_excel, bg="#4CAF50", fg="white",
                               font=("Arial", 12), relief="flat", bd=0)
        export_btn.pack(side="right", padx=5)
        ToolTip(export_btn, "导出为Excel文件")

        filter_frame = tk.Frame(self.root, bg="#E0F2E9")
        filter_frame.pack(fill="x", pady=5)
        self.filter_vars = {}
        columns = ["大类", "中类", "小类", "金额", "账户"]
        for i, col in enumerate(columns):
            tk.Label(filter_frame, text=col, bg="#E0F2E9", fg="#333333", font=("Arial", 10)).pack(side="left", padx=5)
            if col == "金额":
                self.filter_vars[col] = (tk.StringVar(value=""), tk.StringVar(value=""))
                min_entry = tk.Entry(filter_frame, textvariable=self.filter_vars[col][0], width=8, font=("Arial", 10),
                                     relief="flat", bd=1, highlightthickness=1)
                max_entry = tk.Entry(filter_frame, textvariable=self.filter_vars[col][1], width=8, font=("Arial", 10),
                                     relief="flat", bd=1, highlightthickness=1)
                min_entry.pack(side="left", padx=2)
                tk.Label(filter_frame, text="-", bg="#E0F2E9", fg="#333333").pack(side="left")
                max_entry.pack(side="left", padx=2)
                min_entry.bind("<KeyRelease>", self.apply_filters)
                max_entry.bind("<KeyRelease>", self.apply_filters)
            else:
                values = sorted(set(r[col] for r in self.data_manager.records)) if self.data_manager.records else []
                self.filter_vars[col] = tk.StringVar(value="所有")
                combo = ttk.Combobox(filter_frame, textvariable=self.filter_vars[col], values=["所有"] + values,
                                     state="readonly", width=10)
                combo.pack(side="left", padx=5)
                combo.bind("<<ComboboxSelected>>", self.apply_filters)

        search_frame = tk.Frame(self.root, bg="#E0F2E9")
        search_frame.pack(fill="x", pady=5)
        tk.Label(search_frame, text="搜索备注:", bg="#E0F2E9", fg="#333333", font=("Arial", 12)).pack(side="left",
                                                                                                      padx=5)
        self.search_var = tk.StringVar()
        tk.Entry(search_frame, textvariable=self.search_var, font=("Arial", 12), width=20, relief="flat", bd=1,
                 highlightthickness=1).pack(side="left", padx=5)
        tk.Button(search_frame, text="\U0001F50D", command=self.search_records, bg="#42A5F5", fg="white",
                  font=("Arial", 12), relief="flat", bd=0).pack(side="left", padx=5)

        main_frame = tk.Frame(self.root, bg="#E0F2E9")
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        header_frame = tk.Frame(main_frame, bg="#F5F6F5", height=30)
        header_frame.pack(fill="x")
        tk.Label(header_frame, text="大类", font=("Arial", 12, "bold"), bg="#F5F6F5", anchor="center").place(x=50,
                                                                                                             y=5)  # 保持居中
        tk.Label(header_frame, text="中类", font=("Arial", 12, "bold"), bg="#F5F6F5", anchor="center").place(x=200, y=5)
        tk.Label(header_frame, text="小类", font=("Arial", 12, "bold"), bg="#F5F6F5", anchor="center").place(x=350, y=5)
        tk.Label(header_frame, text="金额",
                 font=("Consolas", 12, "bold"),  # 改用等宽字体
                 bg="#F5F6F5",
                 anchor="e",  # 强制右对齐
                 width=15).place(x=500, y=5)  # 精确控制列宽

        tk.Label(header_frame, text="备注",
                 font=("Arial", 12, "bold"),
                 bg="#F5F6F5",
                 anchor="w",  # 强制左对齐
                 width=25).place(x=650, y=5)  # 扩大列宽

        tk.Label(header_frame, text="账户",
                 font=("Arial", 12, "bold"),
                 bg="#F5F6F5",
                 anchor="center",
                 width=12).place(x=800, y=5)

        canvas_frame = tk.Frame(main_frame, bg="#F5F6F5")
        canvas_frame.pack(fill="both", expand=True)
        self.canvas = Canvas(canvas_frame, bg="#F5F6F5", width=1000, height=400)
        self.scrollbar = tk.Scrollbar(canvas_frame, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        self.canvas.bind_all("<MouseWheel>", lambda e: self.canvas.yview_scroll(-1 * (e.delta // 120), "units"))

        current_month = datetime.now().strftime("%Y年%m月")
        self.stats_label = tk.Label(self.root, text=f"{current_month}支出: 0.00 元 | 收入: 0.00 元", font=("Arial", 12),
                                    bg="#E0F2E9", fg="#333333")
        self.stats_label.pack(pady=5)

        self.update_treeview()
        self.update_stats()

    def update_treeview(self, filtered_records=None):
        self.canvas.delete("all")
        records = filtered_records or sorted(self.data_manager.records, key=lambda x: x["日期"], reverse=True)
        y_offset = 10
        daily_records = {}
        for record in records:
            date = record["日期"]
            if date not in daily_records:
                daily_records[date] = []
            daily_records[date].append(record)

        for date, recs in daily_records.items():
            if y_offset > 10:
                self.canvas.create_line(10, y_offset, 990, y_offset, fill="#D3D3D3", width=1)
                y_offset += 10
            daily_out = sum(r["金额"] for r in recs if r["金额"] < 0)
            daily_in = sum(r["金额"] for r in recs if r["金额"] > 0)
            stats_text = f"{date} {recs[0]['星期']}  支出: {abs(daily_out):.2f} 元" + (
                f" | 收入: {daily_in:.2f} 元" if daily_in > 0 else "")
            self.canvas.create_text(20, y_offset + 10, text=stats_text, font=("Arial", 12, "italic"), anchor="w",
                                    fill="#666666")
            y_offset += 20
            for rec in recs:
                self.canvas.create_text(50, y_offset + 10, text=rec["大类"], font=("Arial", 12), anchor="center")
                self.canvas.create_text(200, y_offset + 10, text=rec["中类"], font=("Arial", 12), anchor="center")
                self.canvas.create_text(350, y_offset + 10, text=rec["小类"], font=("Arial", 12), anchor="center")
                amount_tag = self.canvas.create_text(
                    575, y_offset + 10,
                    text=f"{rec['金额']:+,.2f}",
                    font=("Consolas", 12),  # 必须使用等宽字体
                    anchor="e",  # 严格右对齐
                    width=150  # 设置绘制区域宽度（关键参数）
                )
                self.canvas.create_text(650, y_offset + 10, text=rec["备注"], font=("Arial", 12), anchor="center")
                self.canvas.create_text(800, y_offset + 10, text=rec["账户"], font=("Arial", 12), anchor="center")
                self.canvas.tag_bind(amount_tag, "<Button-1>", lambda e, r=rec: self.select_record(r))
                y_offset += 20

        self.canvas.configure(scrollregion=(0, 0, 1000, y_offset + 10))

    def select_record(self, record):
        self.selected_record = record

    def update_stats(self):
        current_month = datetime.now().strftime("%Y-%m")
        monthly_out = sum(r["金额"] for r in self.data_manager.records if r["日期"].startswith(current_month) and r["金额"] < 0)
        monthly_in = sum(r["金额"] for r in self.data_manager.records if r["日期"].startswith(current_month) and r["金额"] > 0)
        self.stats_label.config(text=f"{current_month[:4]}年{current_month[5:]}月支出: {abs(monthly_out):.2f} 元 | 收入: {monthly_in:.2f} 元")

    def apply_filters(self, event=None):
        filtered = self.data_manager.records[:]
        for col, var in self.filter_vars.items():
            if col == "金额":
                min_val, max_val = var[0].get(), var[1].get()
                try:
                    min_val = float(min_val) if min_val else float("-inf")
                    max_val = float(max_val) if max_val else float("inf")
                    filtered = [r for r in filtered if min_val <= (r["金额"] if r["金额"] > 0 else -r["金额"]) <= max_val]
                except ValueError:
                    pass
            else:
                value = var.get()
                if value != "所有":
                    filtered = [r for r in filtered if r[col] == value]
        self.update_treeview(filtered)

    def search_records(self):
        search_term = self.search_var.get().strip()
        if not search_term:
            self.update_treeview()
            return
        try:
            amount = float(search_term)
            filtered = [r for r in self.data_manager.records if abs(r["金额"]) == amount or search_term in r["备注"]]
        except ValueError:
            filtered = [r for r in self.data_manager.records if search_term in r["备注"]]
        self.update_treeview(filtered)

    def open_add_window(self, is_income):
        add_window = tk.Toplevel(self.root)
        add_window.title("记账")
        add_window.geometry("400x700")
        add_window.configure(bg="#ffffff")

        date_var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        account_var = tk.StringVar(value=self.data_manager.accounts[0])
        category_var = tk.StringVar()
        subcategory_var = tk.StringVar()
        item_var = tk.StringVar()
        amount_var = tk.StringVar()
        note_var = tk.StringVar()

        tk.Label(add_window, text="记账", font=("Arial", 16, "bold"), bg="#ffffff", fg="#333333").grid(row=0, column=0, columnspan=2, pady=10)
        tk.Label(add_window, text="日期", font=("Arial", 12), bg="#ffffff", fg="#333333").grid(row=1, column=0, pady=5, sticky="w")
        date_entry = tk.Entry(add_window, textvariable=date_var, font=("Arial", 12), relief="flat", bd=1, highlightthickness=1)
        date_entry.grid(row=1, column=1, pady=5, sticky="e")
        date_entry.bind("<FocusIn>", lambda e: self.show_calendar(add_window, date_entry, date_var))

        tk.Label(add_window, text="账户", font=("Arial", 12), bg="#ffffff", fg="#333333").grid(row=2, column=0, pady=5, sticky="w")
        ttk.OptionMenu(add_window, account_var, self.data_manager.accounts[0], *self.data_manager.accounts).grid(row=2, column=1, pady=5)

        tk.Label(add_window, text="分类", font=("Arial", 12, "bold"), bg="#ffffff", fg="#333333").grid(row=3, column=0, columnspan=2, pady=10)

        tree = ttk.Treeview(add_window, height=15, show="tree")
        tree.grid(row=4, column=0, columnspan=2, padx=10, pady=5, sticky="ew")
        cats = self.data_manager.income_categories if is_income else self.data_manager.expense_categories
        for cat, subcats in cats.items():
            cat_id = tree.insert("", "end", text=cat, open=True)
            for subcat, items in subcats.items():
                subcat_id = tree.insert(cat_id, "end", text=subcat, open=True)
                for item in items:
                    tree.insert(subcat_id, "end", text=item)

        amount_entry = tk.Entry(add_window, textvariable=amount_var, font=("Arial", 12), relief="flat", bd=1, highlightthickness=1)
        note_entry = tk.Entry(add_window, textvariable=note_var, font=("Arial", 12), relief="flat", bd=1, highlightthickness=1)

        def on_select(event):
            amount_entry.grid_forget()
            note_entry.grid_forget()
            add_btn.grid_forget()
            delete_btn.grid_forget()
            selected = tree.selection()
            if selected:
                item = tree.item(selected[0])["text"]
                parent = tree.parent(selected[0])
                grandparent = tree.parent(parent) if parent else ""
                if grandparent:
                    category_var.set(tree.item(grandparent)["text"])
                    subcategory_var.set(tree.item(parent)["text"])
                    item_var.set(item)
                    tk.Label(add_window, text="金额", font=("Arial", 12), bg="#ffffff", fg="#333333").grid(row=5, column=0, pady=5, sticky="w")
                    amount_entry.grid(row=5, column=1, pady=5)
                    tk.Label(add_window, text="备注", font=("Arial", 12), bg="#ffffff", fg="#333333").grid(row=6, column=0, pady=5, sticky="w")
                    note_entry.grid(row=6, column=1, pady=5)
                    add_btn.grid(row=7, column=0, pady=10)
                    delete_btn.grid(row=7, column=1, pady=10)

        tree.bind("<<TreeviewSelect>>", on_select)

        def add_record():
            try:
                date = datetime.strptime(date_var.get(), "%Y-%m-%d")
                amount = float(amount_var.get()) if is_income else -float(amount_var.get())
                if not all([category_var.get(), subcategory_var.get(), item_var.get()]):
                    raise ValueError("分类未选择完整")
                record = {
                    "日期": date_var.get(),
                    "星期": date.strftime("星期%u").replace("1", "一").replace("2", "二").replace("3", "三").replace("4", "四").replace("5", "五").replace("6", "六").replace("7", "日"),
                    "金额": amount,
                    "大类": category_var.get(),
                    "中类": subcategory_var.get(),
                    "小类": item_var.get(),
                    "备注": note_var.get(),
                    "账户": account_var.get()
                }
                self.data_manager.add_record(record)
                self.update_treeview()
                self.update_stats()
                add_window.destroy()
            except ValueError as e:
                messagebox.showerror("错误", str(e) or "请输入有效金额和日期！")

        def delete_item():
            if item_var.get() in cats[category_var.get()][subcategory_var.get()]:
                if messagebox.askyesno("警告", f"删除 '{item_var.get()}' 将移除所有相关记录，确认吗？"):
                    self.data_manager.delete_category(is_income, category_var.get(), subcategory_var.get(), item_var.get())
                    self.update_treeview()
            add_window.destroy()

        add_btn = tk.Button(add_window, text="添加记录", command=add_record, bg="#66BB6A", fg="white", font=("Arial", 12), relief="flat", bd=0)
        delete_btn = tk.Button(add_window, text="删除记录", command=delete_item, bg="#EF5350", fg="white", font=("Arial", 12), relief="flat", bd=0)

    def show_calendar(self, parent, entry, var):
        cal = tk.Toplevel(parent)
        cal.title("选择日期")
        cal.geometry(f"+{entry.winfo_rootx()}+{entry.winfo_rooty() + entry.winfo_height() + 10}")

        # 年份选择
        year_var = tk.StringVar(value=datetime.now().strftime("%Y"))
        tk.Label(cal, text="年").pack(side="left", padx=5, pady=10)
        year_combo = ttk.Combobox(cal, textvariable=year_var, values=[str(y) for y in range(2020, 2031)], state="readonly", width=6)
        year_combo.pack(side="left", padx=5)

        # 月份选择
        month_var = tk.StringVar(value=datetime.now().strftime("%m"))
        tk.Label(cal, text="月").pack(side="left", padx=5, pady=10)
        month_combo = ttk.Combobox(cal, textvariable=month_var, values=[f"{m:02d}" for m in range(1, 13)], state="readonly", width=4)
        month_combo.pack(side="left", padx=5)

        # 日期选择
        day_var = tk.StringVar(value=datetime.now().strftime("%d"))
        tk.Label(cal, text="日").pack(side="left", padx=5, pady=10)
        day_combo = ttk.Combobox(cal, textvariable=day_var, values=[f"{d:02d}" for d in range(1, 32)], state="readonly", width=4)
        day_combo.pack(side="left", padx=5)

        def set_and_close():
            date_str = f"{year_var.get()}-{month_var.get()}-{day_var.get()}"
            try:
                datetime.strptime(date_str, "%Y-%m-%d")
                var.set(date_str)
                print("日期已设置，尝试关闭窗口")  # 调试信息
                cal.destroy()  # 直接销毁
            except ValueError:
                messagebox.showerror("错误", "无效日期！")

        # 确定按钮
        tk.Button(cal, text="确定", command=set_and_close, bg="#66BB6A", fg="white", relief="flat", bd=0).pack(pady=5)

        # 处理“×”关闭
        def close_window():
            print("点击×，尝试关闭窗口")  # 调试信息
            cal.destroy()

        cal.protocol("WM_DELETE_WINDOW", close_window)

    def edit_record(self):
        if not hasattr(self, "selected_record"):
            messagebox.showerror("错误", "请选择一条记录！（点击金额选择）")
            return
        item = self.selected_record
        edit_window = tk.Toplevel(self.root)
        edit_window.title("编辑")
        edit_window.geometry("400x600")
        edit_window.configure(bg="#ffffff")

        date_var = tk.StringVar(value=item["日期"])
        amount_var = tk.StringVar(value=str(abs(item["金额"])))
        note_var = tk.StringVar(value=item["备注"])
        account_var = tk.StringVar(value=item["账户"])
        category_var = tk.StringVar(value=item["大类"])
        subcategory_var = tk.StringVar(value=item["中类"])
        item_var = tk.StringVar(value=item["小类"])

        tk.Label(edit_window, text="日期", font=("Arial", 12), bg="#ffffff", fg="#333333").grid(row=0, column=0, pady=5, sticky="w")
        date_entry = tk.Entry(edit_window, textvariable=date_var, font=("Arial", 12), relief="flat", bd=1, highlightthickness=1)
        date_entry.grid(row=0, column=1, pady=5, sticky="e")
        date_entry.bind("<FocusIn>", lambda e: self.show_calendar(edit_window, date_entry, date_var))

        tk.Label(edit_window, text="金额", font=("Arial", 12), bg="#ffffff", fg="#333333").grid(row=1, column=0, pady=5, sticky="w")
        tk.Entry(edit_window, textvariable=amount_var, font=("Arial", 12), relief="flat", bd=1, highlightthickness=1).grid(row=1, column=1, pady=5)
        tk.Label(edit_window, text="备注", font=("Arial", 12), bg="#ffffff", fg="#333333").grid(row=2, column=0, pady=5, sticky="w")
        tk.Entry(edit_window, textvariable=note_var, font=("Arial", 12), relief="flat", bd=1, highlightthickness=1).grid(row=2, column=1, pady=5)
        tk.Label(edit_window, text="账户", font=("Arial", 12), bg="#ffffff", fg="#333333").grid(row=3, column=0, pady=5, sticky="w")
        ttk.OptionMenu(edit_window, account_var, item["账户"], *self.data_manager.accounts).grid(row=3, column=1, pady=5)

        tk.Label(edit_window, text="分类", font=("Arial", 12, "bold"), bg="#ffffff", fg="#333333").grid(row=4, column=0, columnspan=2, pady=10)
        tree = ttk.Treeview(edit_window, height=10, show="tree")
        tree.grid(row=5, column=0, columnspan=2, padx=10, pady=5, sticky="ew")
        cats = self.data_manager.income_categories if item["金额"] > 0 else self.data_manager.expense_categories
        for cat, subcats in cats.items():
            cat_id = tree.insert("", "end", text=cat, open=True)
            for subcat, items in subcats.items():
                subcat_id = tree.insert(cat_id, "end", text=subcat, open=True)
                for item in items:
                    tree.insert(subcat_id, "end", text=item)

        def on_select(event):
            selected = tree.selection()
            if selected:
                item_text = tree.item(selected[0])["text"]
                parent = tree.parent(selected[0])
                grandparent = tree.parent(parent) if parent else ""
                if grandparent:
                    category_var.set(tree.item(grandparent)["text"])
                    subcategory_var.set(tree.item(parent)["text"])
                    item_var.set(item_text)

        tree.bind("<<TreeviewSelect>>", on_select)

        def save_edit():
            try:
                date = datetime.strptime(date_var.get(), "%Y-%m-%d")
                amount = float(amount_var.get()) if item["金额"] > 0 else -float(amount_var.get())
                index = self.data_manager.records.index(item)
                new_record = {
                    "日期": date_var.get(),
                    "星期": date.strftime("星期%u").replace("1", "一").replace("2", "二").replace("3", "三").replace("4", "四").replace("5", "五").replace("6", "六").replace("7", "日"),
                    "金额": amount,
                    "大类": category_var.get(),
                    "中类": subcategory_var.get(),
                    "小类": item_var.get(),
                    "备注": note_var.get(),
                    "账户": account_var.get()
                }
                self.data_manager.edit_record(index, new_record)
                self.update_treeview()
                self.update_stats()
                edit_window.destroy()
            except ValueError:
                messagebox.showerror("错误", "请输入有效金额和日期！")

        tk.Button(edit_window, text="保存", command=save_edit, bg="#66BB6A", fg="white", font=("Arial", 12), relief="flat", bd=0).grid(row=6, column=0, columnspan=2, pady=10)

    def delete_record(self):
        if not hasattr(self, "selected_record"):
            messagebox.showerror("错误", "请选择一条记录！（点击金额选择）")
            return
        index = self.data_manager.records.index(self.selected_record)
        self.data_manager.delete_record(index)
        delattr(self, "selected_record")
        self.update_treeview()
        self.update_stats()

    def export_to_excel(self):
        if not self.data_manager.records:
            messagebox.showerror("错误", "没有记录可导出！")
            return
        df = pd.DataFrame(self.data_manager.records)
        df["金额"] = pd.to_numeric(df["金额"])
        filename = f"小马记账_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        df.to_excel(filename, index=False, float_format="%.2f")
        messagebox.showinfo("成功", f"已导出到 {filename}")

    def add_periodic_record(self):
        periodic_window = tk.Toplevel(self.root)
        periodic_window.title("周期记账")
        periodic_window.geometry("400x700")
        periodic_window.configure(bg="#ffffff")

        is_income = tk.BooleanVar(value=False)
        income_btn = tk.Button(periodic_window, text="收入", command=lambda: [is_income.set(True), income_btn.config(bg="#EF5350"), expense_btn.config(bg="#D3D3D3")], bg="#EF5350" if is_income.get() else "#D3D3D3", fg="white", font=("Arial", 12), relief="flat", bd=0)
        expense_btn = tk.Button(periodic_window, text="支出", command=lambda: [is_income.set(False), income_btn.config(bg="#D3D3D3"), expense_btn.config(bg="#66BB6A")], bg="#66BB6A" if not is_income.get() else "#D3D3D3", fg="white", font=("Arial", 12), relief="flat", bd=0)
        tk.Label(periodic_window, text="周期记账", font=("Arial", 16, "bold"), bg="#ffffff", fg="#333333").grid(row=0, column=0, columnspan=2, pady=10)
        expense_btn.grid(row=1, column=0, pady=5)
        income_btn.grid(row=1, column=1, pady=5)

        amount_var = tk.StringVar()
        note_var = tk.StringVar()
        period_var = tk.StringVar(value="每年")
        start_date_var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        end_date_var = tk.StringVar(value=(datetime.now() + relativedelta(years=1)).strftime("%Y-%m-%d"))
        category_var = tk.StringVar()
        subcategory_var = tk.StringVar()
        item_var = tk.StringVar()

        tk.Label(periodic_window, text="金额", font=("Arial", 12), bg="#ffffff", fg="#333333").grid(row=2, column=0, pady=5, sticky="w")
        tk.Entry(periodic_window, textvariable=amount_var, font=("Arial", 12), relief="flat", bd=1, highlightthickness=1).grid(row=2, column=1, pady=5)
        tk.Label(periodic_window, text="周期", font=("Arial", 12), bg="#ffffff", fg="#333333").grid(row=3, column=0, pady=5, sticky="w")
        ttk.OptionMenu(periodic_window, period_var, "每年", "每年", "每季", "每月", "每周").grid(row=3, column=1, pady=5)
        tk.Label(periodic_window, text="开始日期", font=("Arial", 12), bg="#ffffff", fg="#333333").grid(row=4, column=0, pady=5, sticky="w")
        start_date_entry = tk.Entry(periodic_window, textvariable=start_date_var, font=("Arial", 12), relief="flat", bd=1, highlightthickness=1)
        start_date_entry.grid(row=4, column=1, pady=5)
        start_date_entry.bind("<FocusIn>", lambda e: self.show_calendar(periodic_window, start_date_entry, start_date_var))
        tk.Label(periodic_window, text="截止日期", font=("Arial", 12), bg="#ffffff", fg="#333333").grid(row=5, column=0, pady=5, sticky="w")
        end_date_entry = tk.Entry(periodic_window, textvariable=end_date_var, font=("Arial", 12), relief="flat", bd=1, highlightthickness=1)
        end_date_entry.grid(row=5, column=1, pady=5)
        end_date_entry.bind("<FocusIn>", lambda e: self.show_calendar(periodic_window, end_date_entry, end_date_var))
        tk.Label(periodic_window, text="备注", font=("Arial", 12), bg="#ffffff", fg="#333333").grid(row=6, column=0, pady=5, sticky="w")
        tk.Entry(periodic_window, textvariable=note_var, font=("Arial", 12), relief="flat", bd=1, highlightthickness=1).grid(row=6, column=1, pady=5)

        tree = ttk.Treeview(periodic_window, height=10, show="tree")
        tree.grid(row=7, column=0, columnspan=2, padx=10, pady=5, sticky="ew")

        def populate_tree():
            tree.delete(*tree.get_children())
            cats = self.data_manager.income_categories if is_income.get() else self.data_manager.expense_categories
            for cat, subcats in cats.items():
                cat_id = tree.insert("", "end", text=cat, open=True)
                for subcat, items in subcats.items():
                    subcat_id = tree.insert(cat_id, "end", text=subcat, open=True)
                    for item in items:
                        tree.insert(subcat_id, "end", text=item)

        populate_tree()
        is_income.trace("w", lambda *args: populate_tree())

        def on_select(event):
            selected = tree.selection()
            if selected:
                item = tree.item(selected[0])["text"]
                parent = tree.parent(selected[0])
                grandparent = tree.parent(parent) if parent else ""
                if grandparent:
                    category_var.set(tree.item(grandparent)["text"])
                    subcategory_var.set(tree.item(parent)["text"])
                    item_var.set(item)

        tree.bind("<<TreeviewSelect>>", on_select)

        def save_periodic():
            try:
                amount = float(amount_var.get()) if is_income.get() else -float(amount_var.get())
                if not all([category_var.get(), subcategory_var.get(), item_var.get()]):
                    raise ValueError("分类未选择完整")
                start_date = datetime.strptime(start_date_var.get(), "%Y-%m-%d")
                end_date = datetime.strptime(end_date_var.get(), "%Y-%m-%d")
                period = period_var.get()
                current_date = start_date
                while current_date <= end_date:
                    date_str = current_date.strftime("%Y-%m-%d")
                    record = {
                        "日期": date_str,
                        "星期": current_date.strftime("星期%u").replace("1", "一").replace("2", "二").replace("3", "三").replace("4", "四").replace("5", "五").replace("6", "六").replace("7", "日"),
                        "金额": amount,
                        "大类": category_var.get(),
                        "中类": subcategory_var.get(),
                        "小类": item_var.get(),
                        "备注": note_var.get(),
                        "账户": "银行卡"
                    }
                    self.data_manager.add_record(record)
                    if period == "每年":
                        current_date += relativedelta(years=1)
                    elif period == "每季":
                        current_date += relativedelta(months=3)
                    elif period == "每月":
                        current_date += relativedelta(months=1)
                    else:  # 每周
                        current_date += relativedelta(weeks=1)
                self.update_treeview()
                self.update_stats()
                periodic_window.destroy()
            except ValueError as e:
                messagebox.showerror("错误", str(e) or "请输入有效金额和日期！")

        tk.Button(periodic_window, text="保存", command=save_periodic, bg="#66BB6A", fg="white", font=("Arial", 12), relief="flat", bd=0).grid(row=8, column=0, columnspan=2, pady=10)

    def manage_categories(self):
        manage_window = tk.Toplevel(self.root)
        manage_window.title("管理类别")
        manage_window.geometry("500x700")
        manage_window.configure(bg="#ffffff")

        is_income = tk.BooleanVar(value=False)
        income_btn = tk.Button(manage_window, text="收入", command=lambda: [is_income.set(True), income_btn.config(bg="#EF5350"), expense_btn.config(bg="#D3D3D3")], bg="#EF5350" if is_income.get() else "#D3D3D3", fg="white", font=("Arial", 12), relief="flat", bd=0)
        expense_btn = tk.Button(manage_window, text="支出", command=lambda: [is_income.set(False), income_btn.config(bg="#D3D3D3"), expense_btn.config(bg="#66BB6A")], bg="#66BB6A" if not is_income.get() else "#D3D3D3", fg="white", font=("Arial", 12), relief="flat", bd=0)
        tk.Label(manage_window, text="类别管理", font=("Arial", 16, "bold"), bg="#ffffff", fg="#333333").grid(row=0, column=0, columnspan=3, pady=10)
        expense_btn.grid(row=1, column=0, pady=5)
        income_btn.grid(row=1, column=1, pady=5)

        tk.Label(manage_window, text="已有类别", font=("Arial", 12, "bold"), bg="#ffffff", fg="#333333").grid(row=2, column=0, columnspan=3, pady=10)
        tree = ttk.Treeview(manage_window, height=10, show="tree")
        tree.grid(row=3, column=0, columnspan=3, padx=10, pady=5, sticky="ew")

        def populate_tree():
            tree.delete(*tree.get_children())
            cats = self.data_manager.income_categories if is_income.get() else self.data_manager.expense_categories
            for cat, subcats in cats.items():
                cat_id = tree.insert("", "end", text=cat, open=True)
                for subcat, items in subcats.items():
                    subcat_id = tree.insert(cat_id, "end", text=subcat, open=True)
                    for item in items:
                        tree.insert(subcat_id, "end", text=item)

        populate_tree()
        is_income.trace("w", lambda *args: [populate_tree(), update_combos()])

        tk.Label(manage_window, text="添加类别", font=("Arial", 12, "bold"), bg="#ffffff", fg="#333333").grid(row=4, column=0, columnspan=3, pady=10)
        cat_var = tk.StringVar()
        subcat_var = tk.StringVar()
        new_item_var = tk.StringVar()

        def update_combos(event=None):
            cats = self.data_manager.income_categories if is_income.get() else self.data_manager.expense_categories
            cat_combo["values"] = list(cats.keys())
            if event is None and cat_combo["values"]:
                cat_var.set(cat_combo["values"][0])
            subcat_combo["values"] = list(cats.get(cat_var.get(), {}).keys())
            if subcat_combo["values"]:
                subcat_var.set(subcat_combo["values"][0])
            else:
                subcat_var.set("其他")

        tk.Label(manage_window, text="大类", font=("Arial", 12), bg="#ffffff", fg="#333333").grid(row=5, column=0,
                                                                                                  pady=5, sticky="w")
        cat_combo = ttk.Combobox(manage_window, textvariable=cat_var,
                                 values=list(self.data_manager.expense_categories.keys()), state="readonly")
        cat_combo.grid(row=5, column=1, pady=5)
        cat_combo.bind("<<ComboboxSelected>>", update_combos)
        tk.Label(manage_window, text="中类", font=("Arial", 12), bg="#ffffff", fg="#333333").grid(row=6, column=0,
                                                                                                  pady=5, sticky="w")
        subcat_combo = ttk.Combobox(manage_window, textvariable=subcat_var, state="readonly")
        subcat_combo.grid(row=6, column=1, pady=5)
        tk.Label(manage_window, text="小类", font=("Arial", 12), bg="#ffffff", fg="#333333").grid(row=7, column=0, pady=5, sticky="w")
        tk.Entry(manage_window, textvariable=new_item_var, font=("Arial", 12), relief="flat", bd=1, highlightthickness=1).grid(row=7, column=1, pady=5)

        tk.Label(manage_window, text="新大类/中类", font=("Arial", 12), bg="#ffffff", fg="#333333").grid(row=8, column=0, pady=5, sticky="w")
        new_cat_var = tk.StringVar()
        new_subcat_var = tk.StringVar()
        tk.Entry(manage_window, textvariable=new_cat_var, font=("Arial", 12), relief="flat", bd=1, highlightthickness=1).grid(row=8, column=1, pady=5)
        tk.Entry(manage_window, textvariable=new_subcat_var, font=("Arial", 12), relief="flat", bd=1, highlightthickness=1).grid(row=9, column=1, pady=5)

        def add_category():
            cat = new_cat_var.get().strip() or cat_var.get()
            subcat = new_subcat_var.get().strip() or subcat_var.get()
            item = new_item_var.get().strip()
            if not item and not new_cat_var.get().strip() and not new_subcat_var.get().strip():
                messagebox.showerror("错误", "至少填写新小类！")
                return
            self.data_manager.add_category(is_income.get(), cat, subcat, item if item else None)
            new_item_var.set("")
            new_cat_var.set("")
            new_subcat_var.set("")
            populate_tree()
            update_combos()

        def delete_category():
            selected = tree.selection()
            if not selected:
                messagebox.showerror("错误", "请选择一个类别！")
                return

            item = tree.item(selected[0])["text"]
            parent = tree.parent(selected[0])
            grandparent = tree.parent(parent) if parent else ""

            cats = self.data_manager.income_categories if is_income.get() else self.data_manager.expense_categories

            if grandparent:  # 删除小类
                cat = tree.item(grandparent)["text"]
                subcat = tree.item(parent)["text"]
                if messagebox.askyesno("警告", f"删除小类 '{item}' 将移除所有相关记录，确认吗？"):
                    self.data_manager.delete_category(is_income.get(), cat, subcat, item)
            elif parent:  # 删除中类
                cat = tree.item(parent)["text"]
                if messagebox.askyesno("警告", f"删除中类 '{item}' 将移除其下所有小类及相关记录，确认吗？"):
                    self.data_manager.delete_category(is_income.get(), cat, item, None)
            else:  # 删除大类
                if messagebox.askyesno("警告", f"删除大类 '{item}' 将移除其下所有中类、小类及相关记录，确认吗？"):
                    self.data_manager.delete_category(is_income.get(), item, None, None)

            self.update_treeview()
            populate_tree()

        update_combos()
        tk.Button(manage_window, text="添加类别", command=add_category, bg="#66BB6A", fg="white", font=("Arial", 12), relief="flat", bd=0).grid(row=10, column=0, pady=5)
        tk.Button(manage_window, text="删除类别", command=delete_category, bg="#EF5350", fg="white", font=("Arial", 12), relief="flat", bd=0).grid(row=10, column=1, pady=5)

if __name__ == "__main__":
    root = ThemedTk(theme="radiance")
    app = XiaoMaApp(root)
    root.mainloop()
