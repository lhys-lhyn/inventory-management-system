#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Author: lhys
# File  : app.py

import os
import tkinter as tk
from tkcalendar import DateEntry
from tkinter import ttk, messagebox, filedialog

from datetime import datetime, timedelta

import pandas as pd
from utils import SqliteOperation, Record, make_path_exists

class SearchCombobox(ttk.Combobox):
    def __init__(self, root, db, id_var=None, **kwargs):
        super().__init__(root, **kwargs)
        self.db = db
        self.id_var = id_var
        self.after_id = None
        self.bind('<KeyRelease>', self.on_input)
        self.bind("<<ComboboxSelected>>", self.on_select)
        self.last_text = ''

    def on_select(self, *args):
        kw = self.get()
        self.last_text = kw
        # 如果需要则设置编号
        if self.id_var is not None:
            ids, _ = self.db.search(
                'products',
                (kw, ),
                field='id',
                limit='where `Name` = ?'
            )
            self.id_var.set(ids[0][0] if ids else '')

    def on_input(self, *args):
        text = self.get()
        # 等待用户输入新关键词
        if not text or text == self.last_text:
            self.event_generate('<Escape>')
            if self.after_id: self.after_cancel(self.after_id)
            return
        self.after_id = self.after(500, self.show_suggestions)

    def show_suggestions(self):
        kw = self.get()
        self.last_text = kw
        search_text = '%'.join(kw)
        results, _ = self.db.search(
            'products',
            (f'%{search_text}%', ),
            field='Name',
            limit='where `Name` LIKE ?'
        )
        product_names = [row[0] for row in results]
        # 保持用户输入
        self.set(kw)
        # 更新下拉选项
        self.configure(values=product_names)
        # 打开下拉框
        self.event_generate('<Down>')
        # 删除编号值
        if self.id_var is not None:
            self.id_var.set('')

    def bind_return(self, *args, **kwargs):
        self.on_select()
        self.bind('<Return>', *args, **kwargs)

class App:
    def __init__(self, title="出入库管理系统", menu_style='bar'):
        self.root = tk.Tk()
        self.root.title(title)

        # 设置窗口大小
        self.root.geometry('780x650')
        # 禁止调整窗口大小
        self.root.resizable(False, False)

        self.create_widgets(menu_style=menu_style)
        self.recorder = Record(self.log_text)
        self.db = SqliteOperation("data.db", self.recorder)
        self.init_db()
        self.show_home_page()

        # 在窗口关闭时调用 self.on_closing 方法
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.root.mainloop()

    def init_db(self):
        self.db.exec_sql('''
        CREATE TABLE IF NOT EXISTS in_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            in_time DATETIME DEFAULT (datetime('now', 'localtime')) NOT NULL,
            product_id TEXT NOT NULL,
            bottles INTEGER NOT NULL,
            unit_price REAL NOT NULL,
            total_price REAL NOT NULL,
            settled BOOLEAN DEFAULT FALSE, 
            FOREIGN KEY (product_id) REFERENCES products (`id`)
        )
        ''')

        self.db.exec_sql('''
        CREATE TABLE IF NOT EXISTS out_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            out_time DATETIME DEFAULT CURRENT_TIMESTAMP NOT NULL,
            product_id TEXT NOT NULL,
            bottles INTEGER NOT NULL,
            unit_price REAL NOT NULL,
            total_price REAL NOT NULL,
            settled BOOLEAN DEFAULT FALSE, 
            FOREIGN KEY (product_id) REFERENCES products (`id`)
        )
        ''')

        self.db.exec_sql('''
        CREATE TABLE IF NOT EXISTS products (
            id TEXT PRIMARY KEY,
            Name TEXT,
            BottlesPerBox INTEGER DEFAULT 0
        )
        ''')

        self.db.exec_sql('''
        CREATE TABLE IF NOT EXISTS statistics (
            product_id TEXT,
            day TEXT DEFAULT (DATE('now')) NOT NULL,
            origin INTEGER DEFAULT 0 NOT NULL,
            in_total INTEGER DEFAULT 0 NOT NULL,
            out_total INTEGER DEFAULT 0 NOT NULL,
            in_unit_price REAL DEFAULT 0 NOT NULL,
            out_unit_price REAL DEFAULT 0 NOT NULL,
            settled BOOLEAN DEFAULT FALSE, 
            profit REAL DEFAULT 0 NOT NULL,
            FOREIGN KEY (product_id) REFERENCES products (`id`)
        )
        ''')

    def create_widgets(self, menu_style='bar'):
        if menu_style == 'bar':
            self.create_menubar()
        else:
            self.create_menulist()

        # 创建主框架，将主框架放置在窗口顶部，填充窗口并允许扩展
        self.main_frame = tk.Frame(self.root)
        self.main_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=5)

        # 创建日志框架，将日志框架放置在窗口底部，并水平填充
        self.log_frame = tk.Frame(self.root)
        self.log_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=3, pady=5)

        # 创建文本框用于日志输出，高度为 5 行，水平填充文本框
        self.log_text = tk.Text(self.log_frame, height=5)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        # 设置 Text 小部件为不可编辑
        self.log_text.config(state='disabled')

        # 创建垂直 Scrollbar 并与 Text 小部件关联
        scrollbar = tk.Scrollbar(self.log_frame, orient='vertical', command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 将 Text 小部件的 yscrollcommand 设置为 Scrollbar 的 set 方法
        self.log_text.config(yscrollcommand=scrollbar.set)

    def create_menulist(self):
        self.menu_frame = tk.Frame(self.root)
        self.menu_frame.pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=10)

        tk.Label(self.menu_frame, text="菜单", font=("Arial", 16)).pack(padx=5, pady=10)
        ttk.Button(self.menu_frame, text="主\n页", command=self.show_home_page).pack(padx=5, pady=5)
        ttk.Button(self.menu_frame, text="新增\n数据", command=self.show_add_data_page).pack(padx=5, pady=5)
        ttk.Button(self.menu_frame, text="查询\n数据", command=self.show_query_page).pack(padx=5, pady=5)
        ttk.Button(self.menu_frame, text="导出\n数据", command=self.show_export_page).pack(padx=5, pady=5)

    def create_menubar(self):
        # 创建菜单栏
        menubar = tk.Menu(self.root)
        # 将菜单栏配置到根窗口
        self.root.config(menu=menubar)

        # 创建主页菜单
        # home_menu = tk.Menu(menubar, tearoff=0)
        # menubar.add_cascade(label="主页", menu=home_menu)
        # home_menu.add_command(label="主页", command=self.show_home_page)
        menubar.add_command(label='主页', command=self.show_home_page)

        # 创建添加数据菜单
        add_data_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="新增数据", menu=add_data_menu)
        add_data_menu.add_command(label="添加出入库记录", command=self.show_add_data_page)
        add_data_menu.add_command(label="批量导入商品", command=self.show_add_product_batches)
        # menubar.add_command(label='新增数据', command=self.show_add_data_page)

        # 创建查询菜单
        # query_menu = tk.Menu(menubar, tearoff=0)
        # menubar.add_cascade(label="查询", menu=query_menu)
        # query_menu.add_command(label="查询数据", command=self.show_query_page)
        menubar.add_command(label='查询数据', command=self.show_query_page)

        # 创建导出菜单
        # export_menu = tk.Menu(menubar, tearoff=0)
        # menubar.add_cascade(label="导出", menu=export_menu)
        # export_menu.add_command(label="导出数据", command=self.show_export_page)
        menubar.add_command(label='导出数据', command=self.show_export_page)

    def log(self, message, level='info', out='log'):
        self.recorder.lock_output(message, level=level, out=out)

    def show_home_page(self):
        # 清除主框架中的所有小部件
        for widget in self.main_frame.winfo_children():
            widget.destroy()

        # 添加主页标题
        tk.Label(self.main_frame, text="销售情况", font=("华文楷体", 24)).pack(pady=10)

        count = self.get_statistics()
        if count.empty:
            tk.Label(self.main_frame, text='当前无数据', font=("华文楷体", 24)).pack(pady=10)
        else:
            import matplotlib.pyplot as plt
            from matplotlib.figure import Figure
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

            # 显示中文和负号
            plt.rcParams['font.sans-serif'] = ['SimHei']
            plt.rcParams['axes.unicode_minus'] = False

            # 创建一个 Figure 对象
            fig = Figure(figsize=(5, 4), dpi=100)
            # 添加子图
            ax1 = fig.add_subplot(111)
            # ax2 = fig.add_subplot(223)
            # ax3 = fig.add_subplot(224)
            # ax4 = fig.add_subplot(224)

            # 第一张图绘制最近一周利润走势图
            count['day'] = pd.to_datetime(count['day']).dt.date
            # 获取今天的日期
            today = datetime.today().date()
            # 最近一周的日期
            date_range = pd.date_range(start=today - timedelta(days=8), end=today - timedelta(days=1), freq='D').to_series().dt.date
            recent_week_profit = pd.Series(index=date_range, dtype=float, name='profit')
            # 过滤出最近一周的数据
            recent_week_profit.update(count[count['day'].isin(date_range)].groupby('day')['profit'].sum())
            recent_week_profit.fillna(0, inplace=True)
            ax1.plot(recent_week_profit)
            ax1.set_title("一周利润走势图")
            # 设置 x 轴刻度和标签
            ax1.set_xticks(date_range)
            ax1.set_xticklabels([d.strftime('%m-%d') for d in date_range])

            # ax2.plot([1, 2, 3], [6, 5, 4])
            # ax3.plot([1, 2, 3], [4, 6, 5])
            # ax4.plot([1, 2, 3], [5, 4, 6])

            # 创建一个 FigureCanvasTkAgg 对象
            canvas = FigureCanvasTkAgg(fig, master=self.main_frame)
            # 绘制图形
            canvas.draw()
            # 将画布放置在主框架中，填充并扩展
            canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        self.log('欢迎使用。', out='text')

    def init_a_row(self, side=tk.TOP, fill=tk.BOTH, expand=True):
        row = tk.Frame(self.main_frame)
        row.pack(side=side, fill=fill, expand=expand, padx=5, pady=10)
        return row

    def show_add_data_page(self):
        # 清除主框架中的所有小部件
        for widget in self.main_frame.winfo_children():
            widget.destroy()

        # 添加“添加出入库记录”标题
        tk.Label(self.main_frame, text="添加出入库记录", font=("华文楷体", 24)).pack(pady=10)

        # 保存字段条目
        self.entries = {
            '商品编号': tk.StringVar(),
            '商品名称': tk.StringVar(),
            '数量': tk.IntVar(),
            '单价': tk.DoubleVar(),
        } if not hasattr(self, 'entries') else self.entries

        # 创建一个行框架
        for i, (field, var) in enumerate(self.entries.items()):
            if field in ['结算状态', '出入库']: continue
            row = self.init_a_row()
            tk.Label(row, width=15, text=field, anchor='w').pack(side=tk.LEFT)

            if i == 1:
                id_var = self.entries.get('商品编号')
                entry = SearchCombobox(row, self.db, id_var)
                entry.pack(side=tk.RIGHT, expand=tk.YES, fill=tk.X)
                self.entries['商品名称'] = entry
            else:
                entry = tk.Entry(row, textvariable=var)
                if i == 2:
                    entry.pack(side=tk.LEFT, expand=tk.YES, fill=tk.X, padx=3)
                    self.unit_var = tk.StringVar()
                    self.unit_var.set('箱')
                    box = ttk.Combobox(row, width=1, textvariable=self.unit_var, values=['箱', '瓶'])
                    box.pack(side=tk.RIGHT, expand=tk.YES, fill=tk.X)
                    # 打开下拉框
                    box.event_generate('<Down>')
                else:
                    entry.pack(side=tk.RIGHT, expand=tk.YES, fill=tk.X)

        row = self.init_a_row()
        tk.Label(row, width=15, text="结算状态", anchor='w').pack(side=tk.LEFT)
        status_var = tk.BooleanVar()
        status_var.set(True)
        # 创建结算状态单选按钮
        tk.Radiobutton(row, text="已结清", variable=status_var, value=True).pack(side=tk.LEFT)
        # 创建结算状态单选按钮
        tk.Radiobutton(row, text="未结清", variable=status_var, value=False).pack(side=tk.LEFT)
        self.entries['结算状态'] = status_var

        row = self.init_a_row()
        # 在类型选择框架中添加标签
        tk.Label(row, width=15, text="类型", anchor='w').pack(side=tk.LEFT)
        # 创建一个 StringVar 用于保存入库/出库选择
        type_var = tk.StringVar()
        # 默认选择入库
        type_var.set('in')
        # 创建入库单选按钮
        tk.Radiobutton(row, text="入库", variable=type_var, value='in').pack(side=tk.LEFT)
        # 创建出库单选按钮
        tk.Radiobutton(row, text="出库", variable=type_var, value='out').pack(side=tk.LEFT)
        self.entries['出入库'] = type_var

        row = self.init_a_row()
        # 添加保存按钮
        ttk.Button(row, text="保存", command=self.save_data).pack(pady=20)

    def get_statistics(self, start=None, end=datetime.today().strftime('%Y-%m-%d')):
        if not start:
            result, _ = self.db.search('statistics', field='MIN(day)')
            start = result[0][0]
        start = end if not start else start

        result, _ = self.db.search(
            'statistics',
            (start, end),
            limit='where `day` BETWEEN ? AND ?'
        )
        columns = self.db.get_column_names('statistics')
        return pd.DataFrame(result, columns=columns)

    def update_statistics(self, data, mode='in'):
        # data = [product_id, bottles, unit_price, total_price, settled]
        product_id = data.pop(0)

        today = datetime.today().strftime('%Y-%m-%d')
        history, _ = self.db.search(
            'statistics',
            (product_id, ),
            limit='WHERE `product_id` = ? ORDER BY day DESC LIMIT 1'
        )
        if history: history = history[0]

        # data = [bottles, unit_price, total_price, settled]
        in_total = (history[3] if history else 0) + (data[0] if mode == 'in' else 0)
        out_total = (history[4] if history else 0) + (data[0] if mode != 'in' else 0)
        in_total_price = (history[5] * history[3] if history else 0) + (data[2] if mode == 'in' else 0)
        out_total_price = (history[6] * history[4] if history else 0) + (data[2] if mode != 'in' else 0)
        data = [
            product_id,
            today,
            history[2] if history else 0,
            in_total,
            out_total,
            in_total_price / in_total if in_total else 0,
            out_total_price / out_total if out_total else 0,
            (history[7] if history else True) and data[-1],
            (history[8] if history else 0) + (data[0] * (data[1] - (history[5] if history else 0)) if mode != 'in' else 0),
        ]
        if not history or history[0][1] != today:
            self.db.insert(
                'statistics',
                data
            )
        else:
            columns = self.db.get_column_names('statistics')
            columns = columns[3:]
            self.db.modify(
                'statistics',
                columns,
                data[3:],
                limit=f'WHERE `product_id` = "{product_id}" and `day` = "{today}"'
            )

    def save_data(self):
        # 出入库选择
        type_var = self.entries.get('出入库')
        table = 'in_records' if type_var.get() == 'in' else 'out_records'

        # 出入库数据
        order = ['商品编号', '商品名称', '数量', '单价', '结算状态']
        data = [self.entries.get(key).get() for key in order]

        # 获得商品编号，如果没有，则新增
        product_id, product_name = data[:2]
        if not product_id and not product_name:
            messagebox.showerror('错误', '请输入商品编号或者商品名称。')
            return

        data = data[2:]

        result, _ = self.db.search(
            'products',
            (product_id, ),
            field='id',
            limit='where `id` = ?'
        )

        if not product_id or not result:
            messagebox.showwarning('提示', f'未找到商品，请点击确定新增数据。')
            return self.add_new_product(product_id, product_name)

        if sum([item == 0 for item in data[:2]]) == 2:
            messagebox.showerror('错误', '请输入数量及价格。')
            return

        data.insert(-1, data[0] * data[1])

        if self.unit_var.get() == '箱':
            unit = int(self.db.search(
                'products',
                (product_id, ),
                field='BottlesPerBox',
                limit='where `id` = ?'
            )[0][0][0])
            num, unit_price = data[:2]
            num *= unit
            unit_price /= unit
            data[:2] = [num, unit_price]

        data = [product_id] + data
        columns = self.db.get_column_names(table)
        # 删除多余列名
        columns.remove('id')
        columns.remove(f'{type_var.get()}_time')
        # 新增数据
        self.db.insert(table, data, keywords=columns)

        # 记录日志
        name = self.db.search(
            'products',
            (product_id, ),
            field='Name',
            limit='where `id` == ?'
        )[0][0][0]
        self.log(
            f"已{'入' if type_var.get() == 'in' else '出'}库 {self.entries.get('数量').get()} {self.unit_var.get()}{name}。",
            out='text'
        )
        messagebox.showinfo('提示', '新增数据成功')

        self.update_statistics(data, type_var.get())

        del self.entries
        return self.show_add_data_page()

    def show_add_product_batches(self):
        for widget in self.main_frame.winfo_children():
            widget.destroy()

        tk.Label(self.main_frame, text="批量导入商品", font=("华文楷体", 24)).pack(pady=10)

        path_frame = tk.Frame(self.main_frame)
        path_frame.pack()
        tk.Label(path_frame, text="请选择 Excel 文件：").pack(side=tk.LEFT, padx=5)
        path_var = tk.StringVar()
        path_entry = tk.Entry(path_frame, textvariable=path_var, state='readonly', width=50)
        path_entry.pack(side=tk.LEFT, padx=5)

        def select_path():
            path = filedialog.askopenfilename()
            if not os.path.isfile(path) or not path.endswith('.xlsx'):
                messagebox.showerror('错误', '请选择 Excel 文件。')
                return
            path_var.set(path)

        path_button = tk.Button(path_frame, text="选择", command=select_path)
        path_button.pack(side=tk.LEFT, padx=5)

        def save():
            path = path_var.get()
            if not path:
                messagebox.showerror('错误', '请选择 Excel 文件。')
                return
            data = pd.read_excel(
                path,
                header=None,
                skiprows=1,
                names=['id', 'name', 'num'],
                dtype={'id': str}
            )
            data.dropna(subset=['name', 'num'], inplace=True)

            data = [
                tuple(row) for row in data.itertuples(index=False, name=None)
                if self.db.search(
                    'products',
                    (row[0], ),
                    limit='where `id` = ?',
                ) == ([], True)
            ]
            if data:
                _, flag = self.db.insert('products', data, mode='multi')
                if not flag: return
                messagebox.showinfo("提示", f'成功导入 {len(data)} 条数据。')
            else:
                messagebox.showinfo("提示", '没有新的商品数据。')

            self.show_add_product_batches()

        export_button = tk.Button(self.main_frame, text="导入", command=save)
        export_button.pack(pady=20)

    def add_new_product(self, product_id, product_name):
        # 创建新窗口
        new_product_window = tk.Toplevel(self.root)
        new_product_window.title("新增商品")
        new_product_window.geometry('500x500')

        tk.Label(new_product_window, text="商品编号:").pack(pady=5)
        product_id_entry = tk.Entry(new_product_window)
        product_id_entry.pack(pady=5)
        product_id_entry.insert(0, product_id)

        tk.Label(new_product_window, text="商品名称:").pack(pady=5)
        product_name_entry = tk.Entry(new_product_window)
        product_name_entry.pack(pady=5)
        product_name_entry.insert(0, product_name)

        tk.Label(new_product_window, text="每箱瓶数:").pack(pady=5)
        unit_entry = tk.Entry(new_product_window)
        unit_entry.pack(pady=5)
        unit_entry.insert(0, '1')

        def save():
            product_id = product_id_entry.get()
            product_name = product_name_entry.get()
            bottles_per_box = unit_entry.get()

            self.db.insert('products', [product_id, product_name, bottles_per_box])
            self.log(f"新增商品: {product_id} - {product_name} - {bottles_per_box}", 'info', out='all')

            new_product_window.destroy()
            self.show_add_data_page()

        tk.Button(new_product_window, text="保存", command=save).pack(pady=10)

    def show_query_page(self):
        for widget in self.main_frame.winfo_children():
            widget.destroy()

        row = self.init_a_row(expand=False)
        tk.Label(row, text="搜索", font=("华文楷体", 24)).pack(pady=10)

        row = self.init_a_row(expand=False)
        search_frame = tk.Frame(row)
        search_frame.pack(pady=10)
        tk.Label(search_frame, text="关键词：").pack(side=tk.LEFT)

        box = SearchCombobox(search_frame, self.db, width=50)
        box.pack(side=tk.LEFT)

        search_button = tk.Button(search_frame, text="搜索")
        search_button.pack(side=tk.LEFT, padx=5)

        row = self.init_a_row(side=None)

        tree = ttk.Treeview(
            row,
            columns=("时间", "商品编号", "商品名称", "数量", "单价（瓶）", "总价", "是否结清"),
            show='headings'
        )

        # 设置每列标题
        widths = [130, 100, 150, 50, 80, 80, 80]
        for i, col in enumerate(tree['columns']):
            tree.heading(col, text=col)
            tree.column(col, width=widths[i], minwidth=widths[i], anchor=tk.CENTER, stretch=True)

        # 创建垂直 Scrollbar 并与 Treeview 关联
        v_scrollbar = ttk.Scrollbar(row, orient='vertical', command=tree.yview)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 创建水平 Scrollbar 并与 Treeview 关联
        h_scrollbar = ttk.Scrollbar(row, orient='horizontal', command=tree.xview)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)

        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree.configure(xscrollcommand=h_scrollbar.set, yscrollcommand=v_scrollbar.set)

        def search_data():
            kw = box.get()
            product_id, _ = self.db.search(
                'products',
                (kw, kw),
                field='id',
                limit='where `id` = ? or `Name` = ?'
            )
            if not product_id:
                messagebox.showerror('错误', '未找到商品。')
                return
            result, _ = self.db.union_search(
                'in_records',
                'out_records',
                (product_id[0][0], product_id[0][0]),
                field1='*, 1',
                field2='*, -1',
                limit='where `product_id` = ?'
            )
            columns = self.db.get_column_names('in_records')
            columns[1] = 'time'
            columns.append('weight')
            result = pd.DataFrame(result, columns=columns)
            result.drop('id', axis=1)
            result = result.sort_values(by=['time'], ascending=False)
            result['bottles'] = result['bottles'] * result['weight']
            result['total_price'] = -1 * result['total_price'] * result['weight']

            # 删除 Treeview 中的所有行
            for row in tree.get_children():
                tree.delete(row)

            # 将查询结果插入到 Treeview 中
            for row in result.itertuples():
                row = list(row)[2:-1]
                name = self.db.search(
                    'products',
                    (row[1], ),
                    field='Name',
                    limit='WHERE `id` = ?'
                )[0][0][0]
                row.insert(2, name)
                row[-1] = '已结清' if row[-1] else '未结清'
                tree.insert("", "end", values=row)

        search_button['command'] = search_data

    def show_export_page(self):
        for widget in self.main_frame.winfo_children():
            widget.destroy()

        tk.Label(self.main_frame, text="导出数据", font=("华文楷体", 24)).pack(pady=10)

        path_frame = self.init_a_row(expand=False, fill=None)
        tk.Label(path_frame, text="请输入要保存的路径:").pack(side=tk.LEFT, padx=5)
        path_var = tk.StringVar()
        path_entry = tk.Entry(path_frame, textvariable=path_var, state='readonly', width=50)
        path_entry.pack(side=tk.LEFT, padx=5)

        def select_path():
            path = filedialog.askdirectory()
            path_var.set(path)

        path_button = tk.Button(path_frame, text="选择", command=select_path)
        path_button.pack(side=tk.LEFT, padx=5)

        # 创建时间框架
        time_frame = self.init_a_row(expand=False, fill=None)
        today = datetime.today()
        # 添加开始时间标签
        tk.Label(time_frame, text="开始时间：").pack(side=tk.LEFT, padx=5)
        # 创建开始时间选择器
        start_time_entry = DateEntry(time_frame, date_pattern='yyyy-mm-dd', maxdate=today)
        # 将开始时间选择器放置在时间框架中
        start_time_entry.pack(side=tk.LEFT)
        # 添加结束时间标签
        tk.Label(time_frame, text="结束时间：").pack(side=tk.LEFT, padx=5)
        # 创建结束时间选择器
        end_time_entry = DateEntry(time_frame, date_pattern='yyyy-mm-dd', maxdate=today)
        # 将结束时间选择器放置在时间框架中
        end_time_entry.pack(side=tk.LEFT)

        def export_one_table(name, start_time, end_time, keep_dir):
            try:
                table = f'{name}_records'
                records, _ = self.db.search(
                    table,
                    (start_time, end_time),
                    limit=f'WHERE `{name}_time` BETWEEN ? AND ?'
                )

                columns = ['', '时间', '商品编号', '数量（瓶）', '单价（瓶）', '总价', '是否结清']
                records = pd.DataFrame(records, columns=columns)
                records['是否结清'].map({0: '未结清', 1: '已结清'})
                records.to_excel(os.path.join(keep_dir, f'{name}.xlsx'), index=False)
                return True
            except Exception as e:
                messagebox.showerror('错误', str(e))
                return False

        def export_data():
            path = path_var.get()
            start_time = start_time_entry.get()
            end_time = end_time_entry.get()
            if start_time > end_time:
                messagebox.showerror('错误', '开始时间要小于终止时间。')

            # 实现导出逻辑
            keep_dir = os.path.join(path, f'{start_time}-{end_time}')
            make_path_exists(keep_dir)

            search_end_time = datetime.strptime(end_time, '%Y-%m-%d')
            search_end_time = (search_end_time + timedelta(days=1)).strftime('%Y-%m-%d')
            flag = export_one_table('in', start_time, search_end_time, keep_dir)
            if not flag: return
            flag = export_one_table('out', start_time, search_end_time, keep_dir)
            if not flag: return

            messagebox.showinfo("提示", f'{start_time} 到 {end_time} 的数据已导出到 "{path}"。')

        export_button = tk.Button(self.main_frame, text="导出", command=export_data)
        export_button.pack(pady=20)

    def on_closing(self):
        # 关闭数据库连接
        self.db.close()
        # 销毁窗口
        self.root.destroy()

if __name__ == "__main__":
    app = App(menu_style='list')