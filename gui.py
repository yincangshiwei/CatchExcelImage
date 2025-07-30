#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
捕获Excel图片工具 - GUI界面
使用tkinter和ttkthemes创建的图形界面
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import logging
import os
import socket
import sys
from datetime import datetime
from ttkthemes import ThemedTk
from core import (
    extract_workbook_images,
    extract_sheet_images, 
    extract_column_images,
    extract_image_by_id,
    get_embedded_image_ids,
    get_floating_image_names
)


class LogHandler(logging.Handler):
    """自定义日志处理器，将日志输出到GUI"""
    
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget
    
    def emit(self, record):
        msg = self.format(record)
        def append():
            self.text_widget.config(state='normal')
            self.text_widget.insert(tk.END, msg + '\n')
            self.text_widget.config(state='disabled')
            self.text_widget.see(tk.END)
        self.text_widget.after(0, append)


class ExcelImageExtractorGUI:
    # 类级别的标志，防止重复检测
    _network_checked = False
    
    def __init__(self):
        # 企业环境检测
        # if not ExcelImageExtractorGUI._network_checked:
        #     ExcelImageExtractorGUI._network_checked = True
        #     if not self.check_network_connection():
        #         return
            
        self.root = ThemedTk(theme="arc")
        self.root.title("捕获Excel图片工具")
        self.root.geometry("800x900")
        
        # 初始化变量（必须在create_widgets之前）
        self.xlsx_path = tk.StringVar()
        # 获取用户桌面路径作为默认输出目录
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        self.output_dir = tk.StringVar(value=desktop_path)
        self.extract_mode = tk.StringVar(value="workbook")
        self.sheet_name = tk.StringVar()
        self.columns = tk.StringVar()
        self.image_id = tk.StringVar()
        self.include_floating = tk.BooleanVar(value=False)  # 是否包含浮动式图片
        
        # 自定义命名相关变量
        self.use_custom_naming = tk.BooleanVar(value=False)  # 是否使用自定义命名
        self.naming_mode = tk.StringVar(value="combination")  # 命名模式：combination或excel_column
        
        # 组合命名相关变量
        self.custom_prefix = tk.StringVar(value="IMG")  # 自定义自定义
        self.include_date = tk.BooleanVar(value=True)  # 是否包含日期
        self.date_format = tk.StringVar(value="%Y%m%d")  # 日期格式
        self.include_sequence = tk.BooleanVar(value=True)  # 是否包含流水号
        self.sequence_digits = tk.StringVar(value="3")  # 流水号位数
        self.name_order = tk.StringVar(value="prefix_date_sequence")  # 命名顺序
        
        # Excel列命名相关变量
        self.excel_columns = tk.StringVar(value="A")  # Excel列（如A或A,B）
        self.column_separator = tk.StringVar(value="-")  # 列之间的分隔符
        
        # 设置日志
        self.setup_logging()
        
        # 创建界面
        self.create_widgets()
        
    def setup_logging(self):
        """设置日志系统"""
        # 创建日志目录
        os.makedirs("logs", exist_ok=True)
        
        # 配置文件日志
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('app.log', encoding='utf-8'),
            ]
        )
        
        self.logger = logging.getLogger(__name__)
        
    def create_widgets(self):
        """创建GUI组件"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 文件选择区域
        file_frame = ttk.LabelFrame(main_frame, text="文件选择", padding="10")
        file_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(file_frame, text="Excel文件:").grid(row=0, column=0, sticky=tk.W)
        self.file_entry = ttk.Entry(file_frame, textvariable=self.xlsx_path, width=50)
        self.file_entry.grid(row=0, column=1, padx=(5, 5), sticky=(tk.W, tk.E))
        ttk.Button(file_frame, text="浏览", command=self.browse_file).grid(row=0, column=2)
        
        ttk.Label(file_frame, text="输出目录:").grid(row=1, column=0, sticky=tk.W, pady=(5, 0))
        self.output_entry = ttk.Entry(file_frame, textvariable=self.output_dir, width=50)
        self.output_entry.grid(row=1, column=1, padx=(5, 5), pady=(5, 0), sticky=(tk.W, tk.E))
        ttk.Button(file_frame, text="浏览", command=self.browse_output_dir).grid(row=1, column=2, pady=(5, 0))
        
        file_frame.columnconfigure(1, weight=1)
        
        # 自定义命名区域
        naming_frame = ttk.LabelFrame(main_frame, text="图片命名设置", padding="10")
        naming_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 是否使用自定义命名
        self.custom_naming_checkbox = ttk.Checkbutton(naming_frame, text="使用自定义图片命名", 
                                                     variable=self.use_custom_naming,
                                                     command=self.on_custom_naming_change)
        self.custom_naming_checkbox.grid(row=0, column=0, columnspan=4, sticky=tk.W, pady=(0, 10))
        
        # 命名模式选择
        self.naming_mode_frame = ttk.Frame(naming_frame)
        self.naming_mode_frame.grid(row=1, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Radiobutton(self.naming_mode_frame, text="组合命名", variable=self.naming_mode, 
                       value="combination", command=self.on_naming_mode_change).grid(row=0, column=0, sticky=tk.W)
        ttk.Radiobutton(self.naming_mode_frame, text="Excel列命名", variable=self.naming_mode, 
                       value="excel_column", command=self.on_naming_mode_change).grid(row=0, column=1, sticky=tk.W, padx=(20, 0))
        
        # 组合命名设置
        self.combination_frame = ttk.LabelFrame(naming_frame, text="组合命名设置", padding="5")
        self.combination_frame.grid(row=2, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=(0, 5))
        
        # 第一行：自定义和日期
        ttk.Label(self.combination_frame, text="自定义:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(self.combination_frame, textvariable=self.custom_prefix, width=10).grid(row=0, column=1, padx=(5, 20), sticky=tk.W)
        
        ttk.Checkbutton(self.combination_frame, text="包含日期", variable=self.include_date).grid(row=0, column=2, sticky=tk.W)
        ttk.Label(self.combination_frame, text="格式:").grid(row=0, column=3, sticky=tk.W, padx=(10, 0))
        date_format_combo = ttk.Combobox(self.combination_frame, width=15, 
                                        values=["年月日(20241201)", "年-月-日(2024-12-01)", "月日(1201)", "年月日时分(20241201_1430)"])
        # 设置对应的实际格式值
        date_format_mapping = {
            "年月日(20241201)": "%Y%m%d",
            "年-月-日(2024-12-01)": "%Y-%m-%d", 
            "月日(1201)": "%m%d",
            "年月日时分(20241201_1430)": "%Y%m%d_%H%M"
        }
        
        def on_date_format_change(event=None):
            selected = date_format_combo.get()
            if selected in date_format_mapping:
                self.date_format.set(date_format_mapping[selected])
        
        date_format_combo.bind('<<ComboboxSelected>>', on_date_format_change)
        # 设置默认显示和默认值
        if self.date_format.get() == "%Y%m%d":
            date_format_combo.set("年月日(20241201)")
        else:
            # 根据当前date_format值找到对应的显示文本
            current_format = self.date_format.get()
            for display_text, format_value in date_format_mapping.items():
                if format_value == current_format:
                    date_format_combo.set(display_text)
                    break
            else:
                # 如果没找到匹配的，设置默认值
                date_format_combo.set("年月日(20241201)")
                self.date_format.set("%Y%m%d")
        date_format_combo.grid(row=0, column=4, padx=(5, 0), sticky=tk.W)
        
        # 第二行：流水号
        ttk.Checkbutton(self.combination_frame, text="包含流水号", variable=self.include_sequence).grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=(5, 0))
        ttk.Label(self.combination_frame, text="位数:").grid(row=1, column=2, sticky=tk.W, pady=(5, 0))
        ttk.Entry(self.combination_frame, textvariable=self.sequence_digits, width=5).grid(row=1, column=3, padx=(5, 0), pady=(5, 0), sticky=tk.W)
        
        # 第三行：命名顺序
        ttk.Label(self.combination_frame, text="顺序:").grid(row=2, column=0, sticky=tk.W, pady=(5, 0))
        order_combo = ttk.Combobox(self.combination_frame, textvariable=self.name_order, width=25,
                                  values=["自定义_日期_流水号", "自定义_流水号_日期", "日期_自定义_流水号", 
                                         "日期_流水号_自定义", "流水号_自定义_日期", "流水号_日期_自定义"])
        # 设置对应的实际顺序值
        order_mapping = {
            "自定义_日期_流水号": "prefix_date_sequence",
            "自定义_流水号_日期": "prefix_sequence_date",
            "日期_自定义_流水号": "date_prefix_sequence",
            "日期_流水号_自定义": "date_sequence_prefix", 
            "流水号_自定义_日期": "sequence_prefix_date",
            "流水号_日期_自定义": "sequence_date_prefix"
        }
        
        def on_order_change(event=None):
            selected = order_combo.get()
            if selected in order_mapping:
                self.name_order.set(order_mapping[selected])
        
        order_combo.bind('<<ComboboxSelected>>', on_order_change)
        # 设置默认显示
        if self.name_order.get() == "prefix_date_sequence":
            order_combo.set("自定义_日期_流水号")
        order_combo.grid(row=2, column=1, columnspan=3, padx=(5, 0), pady=(5, 0), sticky=tk.W)
        
        # Excel列命名设置
        self.excel_column_frame = ttk.LabelFrame(naming_frame, text="Excel列命名设置", padding="5")
        self.excel_column_frame.grid(row=3, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=(0, 5))
        
        ttk.Label(self.excel_column_frame, text="Excel列:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(self.excel_column_frame, textvariable=self.excel_columns, width=15).grid(row=0, column=1, padx=(5, 20), sticky=tk.W)
        ttk.Label(self.excel_column_frame, text="分隔符:").grid(row=0, column=2, sticky=tk.W)
        ttk.Entry(self.excel_column_frame, textvariable=self.column_separator, width=5).grid(row=0, column=3, padx=(5, 0), sticky=tk.W)
        
        ttk.Label(self.excel_column_frame, text="示例: A 或 A,B,C", foreground="gray").grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=(2, 0))
        ttk.Label(self.excel_column_frame, text="留空则直接拼接", foreground="gray").grid(row=1, column=2, columnspan=2, sticky=tk.W, pady=(2, 0))
        
        naming_frame.columnconfigure(0, weight=1)
        
        # 提取模式区域
        mode_frame = ttk.LabelFrame(main_frame, text="提取模式", padding="10")
        mode_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 模式选择
        ttk.Radiobutton(mode_frame, text="整个工作簿", variable=self.extract_mode, 
                       value="workbook", command=self.on_mode_change).grid(row=0, column=0, sticky=tk.W)
        ttk.Radiobutton(mode_frame, text="指定工作表", variable=self.extract_mode, 
                       value="sheet", command=self.on_mode_change).grid(row=0, column=1, sticky=tk.W)
        ttk.Radiobutton(mode_frame, text="指定列", variable=self.extract_mode, 
                       value="column", command=self.on_mode_change).grid(row=0, column=2, sticky=tk.W)
        ttk.Radiobutton(mode_frame, text="指定图片ID", variable=self.extract_mode, 
                       value="id", command=self.on_mode_change).grid(row=0, column=3, sticky=tk.W)
        
        # 参数输入区域
        param_frame = ttk.Frame(mode_frame)
        param_frame.grid(row=1, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # 工作表名称
        self.sheet_label = ttk.Label(param_frame, text="工作表名称:")
        self.sheet_label.grid(row=0, column=0, sticky=tk.W)
        self.sheet_entry = ttk.Entry(param_frame, textvariable=self.sheet_name, width=20)
        self.sheet_entry.grid(row=0, column=1, padx=(5, 20), sticky=tk.W)
        
        # 列名称
        self.column_label = ttk.Label(param_frame, text="列名称:")
        self.column_label.grid(row=0, column=2, sticky=tk.W)
        self.column_entry = ttk.Entry(param_frame, textvariable=self.columns, width=20)
        self.column_entry.grid(row=0, column=3, padx=(5, 20), sticky=tk.W)
        
        # 图片ID
        self.id_label = ttk.Label(param_frame, text="图片ID:")
        self.id_label.grid(row=1, column=0, sticky=tk.W, pady=(5, 0))
        self.id_entry = ttk.Entry(param_frame, textvariable=self.image_id, width=30)
        self.id_entry.grid(row=1, column=1, columnspan=3, padx=(5, 5), pady=(5, 0), sticky=(tk.W, tk.E))
        
        # 列说明
        self.column_help = ttk.Label(param_frame, text="支持格式: A (单列), A,B,C (多列), A-C (范围)", 
                                   foreground="gray")
        self.column_help.grid(row=2, column=2, columnspan=2, sticky=tk.W, pady=(2, 0))
        
        # 图片类型选择
        self.floating_checkbox = ttk.Checkbutton(param_frame, text="包含浮动式图片", 
                                                variable=self.include_floating)
        self.floating_checkbox.grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=(10, 0))
        
        # 图片类型说明
        self.floating_help = ttk.Label(param_frame, 
                                     text="嵌入式图片：通过DISPIMG函数插入的图片；浮动式图片：直接插入的图片对象", 
                                     foreground="gray")
        self.floating_help.grid(row=4, column=0, columnspan=4, sticky=tk.W, pady=(2, 0))
        
        # 指定ID模式的说明（仅在指定ID模式下显示）
        self.id_help = ttk.Label(param_frame, 
                               text="注意：指定图片ID只支持嵌入式图片（通过DISPIMG函数插入的图片）", 
                               foreground="gray")
        self.id_help.grid(row=5, column=0, columnspan=4, sticky=tk.W, pady=(10, 0))
        
        param_frame.columnconfigure(1, weight=1)
        
        # 初始化界面状态
        self.on_mode_change()
        
        # 初始化自定义命名界面状态
        self.on_custom_naming_change()
        
        # 日志区域
        log_frame = ttk.LabelFrame(main_frame, text="操作日志", padding="10")
        log_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        # 创建日志文本框和滚动条
        log_text_frame = ttk.Frame(log_frame)
        log_text_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.log_text = tk.Text(log_text_frame, height=12, state='disabled', wrap=tk.WORD)
        log_scrollbar = ttk.Scrollbar(log_text_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        log_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        log_text_frame.columnconfigure(0, weight=1)
        log_text_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        # 添加GUI日志处理器
        gui_handler = LogHandler(self.log_text)
        gui_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        self.logger.addHandler(gui_handler)
        
        # 生成按钮
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Button(button_frame, text="清空日志", command=self.clear_log).grid(row=0, column=0, sticky=tk.W)
        
        tk.Button(button_frame, text="立即生成", command=self.start_extraction,
                 fg="white", bg="#0078D7", font=('Microsoft YaHei UI', 10, 'bold'),
                 padx=10, pady=5, relief='flat',
                 activebackground="#005A9E", activeforeground="white").grid(row=0, column=1, sticky=tk.E)
        
        button_frame.columnconfigure(0, weight=1)
        
        # 配置主窗口的行列权重
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(3, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
    def browse_file(self):
        """浏览Excel文件"""
        filename = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if filename:
            self.xlsx_path.set(filename)
            self.logger.info(f"选择文件: {filename}")
    
    def browse_output_dir(self):
        """浏览输出目录"""
        dirname = filedialog.askdirectory(title="选择输出目录")
        if dirname:
            self.output_dir.set(dirname)
            self.logger.info(f"选择输出目录: {dirname}")
    
    def on_mode_change(self):
        """模式改变时更新界面"""
        mode = self.extract_mode.get()
        
        # 隐藏所有参数控件
        self.sheet_label.grid_remove()
        self.sheet_entry.grid_remove()
        self.column_label.grid_remove()
        self.column_entry.grid_remove()
        self.column_help.grid_remove()
        self.id_label.grid_remove()
        self.id_entry.grid_remove()
        self.floating_checkbox.grid_remove()
        self.floating_help.grid_remove()
        self.id_help.grid_remove()
        
        # 根据模式显示相应控件
        if mode == "workbook":
            # 整个工作簿模式：显示浮动式图片选项
            self.floating_checkbox.grid()
            self.floating_help.grid()
        elif mode == "sheet":
            # 指定工作表模式：显示工作表输入和浮动式图片选项
            self.sheet_label.grid()
            self.sheet_entry.grid()
            self.floating_checkbox.grid()
            self.floating_help.grid()
        elif mode == "column":
            # 指定列模式：显示工作表、列输入和浮动式图片选项
            self.sheet_label.grid()
            self.sheet_entry.grid()
            self.column_label.grid()
            self.column_entry.grid()
            self.column_help.grid()
            self.floating_checkbox.grid()
            self.floating_help.grid()
        elif mode == "id":
            # 指定ID模式：只显示ID输入和说明，不显示浮动式图片选项
            self.id_label.grid()
            self.id_entry.grid()
            self.id_help.grid()
    
    def on_custom_naming_change(self):
        """自定义命名选项改变时更新界面"""
        if self.use_custom_naming.get():
            self.naming_mode_frame.grid()
            self.on_naming_mode_change()
        else:
            self.naming_mode_frame.grid_remove()
            self.combination_frame.grid_remove()
            self.excel_column_frame.grid_remove()
    
    def on_naming_mode_change(self):
        """命名模式改变时更新界面"""
        if not self.use_custom_naming.get():
            return
            
        mode = self.naming_mode.get()
        if mode == "combination":
            self.combination_frame.grid()
            self.excel_column_frame.grid_remove()
        elif mode == "excel_column":
            self.combination_frame.grid_remove()
            self.excel_column_frame.grid()
    
    def generate_custom_filename(self, base_name, row_data=None, sequence_num=1):
        """生成自定义文件名"""
        if not self.use_custom_naming.get():
            return base_name
        
        mode = self.naming_mode.get()
        
        if mode == "combination":
            # 组合命名模式
            parts = []
            
            # 获取各个组件
            prefix = self.custom_prefix.get() if self.custom_prefix.get() else ""
            date_str = ""
            if self.include_date.get():
                try:
                    date_str = datetime.now().strftime(self.date_format.get())
                except:
                    date_str = datetime.now().strftime("%Y%m%d")
            
            sequence_str = ""
            if self.include_sequence.get():
                try:
                    digits = int(self.sequence_digits.get())
                    sequence_str = str(sequence_num).zfill(digits)
                except:
                    sequence_str = str(sequence_num).zfill(3)
            
            # 根据顺序组合
            order = self.name_order.get()
            if order == "prefix_date_sequence":
                parts = [prefix, date_str, sequence_str]
            elif order == "prefix_sequence_date":
                parts = [prefix, sequence_str, date_str]
            elif order == "date_prefix_sequence":
                parts = [date_str, prefix, sequence_str]
            elif order == "date_sequence_prefix":
                parts = [date_str, sequence_str, prefix]
            elif order == "sequence_prefix_date":
                parts = [sequence_str, prefix, date_str]
            elif order == "sequence_date_prefix":
                parts = [sequence_str, date_str, prefix]
            else:
                parts = [prefix, date_str, sequence_str]
            
            # 过滤空字符串并用下划线连接
            parts = [p for p in parts if p]
            return "_".join(parts) if parts else base_name
        
        elif mode == "excel_column":
            # Excel列命名模式
            columns = self.excel_columns.get().strip()
            if not columns:
                self.logger.warning("Excel列命名模式：未设置列名，使用默认文件名")
                return base_name
            
            if not row_data:
                self.logger.warning("Excel列命名模式：未获取到行数据，使用默认文件名")
                return base_name
            
            # 解析列名
            column_list = [col.strip() for col in columns.split(',')]
            values = []
            
            for col in column_list:
                if col in row_data:
                    value = str(row_data[col]) if row_data[col] is not None else ""
                    values.append(value)
                    self.logger.info(f"Excel列命名：列 {col} = '{value}'")
                else:
                    self.logger.warning(f"Excel列命名：列 {col} 不存在于行数据中")
            
            if values:
                separator = self.column_separator.get() if self.column_separator.get() else ""
                result = separator.join(values) if separator else "".join(values)
                self.logger.info(f"Excel列命名生成文件名: '{result}'")
                return result
            else:
                self.logger.warning("Excel列命名模式：未找到有效的列数据，使用默认文件名")
        
        return base_name
    
    def validate_inputs(self):
        """验证输入参数"""
        if not self.xlsx_path.get():
            raise ValueError("请选择Excel文件")
        
        if not os.path.exists(self.xlsx_path.get()):
            raise ValueError("Excel文件不存在")
        
        if not self.output_dir.get():
            raise ValueError("请设置输出目录")
        
        mode = self.extract_mode.get()
        if mode in ["sheet", "column"] and not self.sheet_name.get():
            raise ValueError("请输入工作表名称")
        
        if mode == "column" and not self.columns.get():
            raise ValueError("请输入列名称")
        
        if mode == "id" and not self.image_id.get():
            raise ValueError("请输入图片ID")
        
        # 验证自定义命名设置
        if self.use_custom_naming.get():
            naming_mode = self.naming_mode.get()
            if naming_mode == "combination":
                # 验证流水号位数
                if self.include_sequence.get():
                    try:
                        digits = int(self.sequence_digits.get())
                        if digits < 1 or digits > 10:
                            raise ValueError("流水号位数必须在1-10之间")
                    except ValueError:
                        raise ValueError("流水号位数必须是有效数字")
                
                # 验证日期格式
                if self.include_date.get():
                    try:
                        datetime.now().strftime(self.date_format.get())
                    except ValueError:
                        raise ValueError("日期格式无效")
            
            elif naming_mode == "excel_column":
                if not self.excel_columns.get().strip():
                    raise ValueError("请输入Excel列名")
                
                # 验证列名格式
                columns = self.excel_columns.get().strip()
                try:
                    column_list = [col.strip() for col in columns.split(',')]
                    for col in column_list:
                        if not col.isalpha() or len(col) > 3:
                            raise ValueError(f"无效的列名: {col}")
                except:
                    raise ValueError("Excel列名格式错误，请使用如 A 或 A,B,C 的格式")
    
    def extract_images(self):
        """执行图片提取"""
        try:
            self.validate_inputs()
            
            xlsx_path = self.xlsx_path.get()
            output_dir = self.output_dir.get()
            mode = self.extract_mode.get()
            include_floating = self.include_floating.get()
            
            self.logger.info(f"开始提取图片 - 模式: {mode}")
            
            # 准备自定义命名函数
            custom_naming_func = None
            if self.use_custom_naming.get():
                custom_naming_func = self.generate_custom_filename
                self.logger.info(f"使用自定义命名 - 模式: {self.naming_mode.get()}")
            
            saved_files = []
            
            if mode == "workbook":
                saved_files = extract_workbook_images(xlsx_path, output_dir, include_floating=include_floating, custom_naming_func=custom_naming_func)
                self.logger.info(f"提取整个工作簿图片完成，共保存 {len(saved_files)} 个文件")
            
            elif mode == "sheet":
                sheet_name = self.sheet_name.get()
                saved_files = extract_sheet_images(xlsx_path, sheet_name, output_dir, include_floating=include_floating, custom_naming_func=custom_naming_func)
                self.logger.info(f"提取工作表 '{sheet_name}' 图片完成，共保存 {len(saved_files)} 个文件")
            
            elif mode == "column":
                sheet_name = self.sheet_name.get()
                columns = self.columns.get()
                saved_files = extract_column_images(xlsx_path, sheet_name, columns, output_dir, include_floating=include_floating, custom_naming_func=custom_naming_func)
                self.logger.info(f"提取工作表 '{sheet_name}' 列 '{columns}' 图片完成，共保存 {len(saved_files)} 个文件")
            
            elif mode == "id":
                image_id = self.image_id.get()
                # 指定图片ID模式只支持嵌入式图片
                saved_file = extract_image_by_id(xlsx_path, image_id, output_dir, custom_naming_func=custom_naming_func)
                if saved_file:
                    saved_files = [saved_file]
                    self.logger.info(f"通过ID '{image_id}' 提取嵌入式图片完成")
                else:
                    self.logger.warning(f"嵌入式图片ID '{image_id}' 不存在")
            
            if saved_files:
                self.logger.info(f"所有图片已保存到: {output_dir}")
                for file_path in saved_files:
                    self.logger.info(f"  - {os.path.basename(file_path)}")
                
                # 在主线程中显示成功消息
                self.root.after(0, lambda: messagebox.showinfo(
                    "提取完成", 
                    f"成功提取 {len(saved_files)} 个图片\n保存位置: {output_dir}"
                ))
            else:
                self.root.after(0, lambda: messagebox.showwarning(
                    "提取完成", "未找到任何图片"
                ))
                
        except Exception as e:
            error_msg = f"提取失败: {str(e)}"
            self.logger.error(error_msg)
            self.root.after(0, lambda: messagebox.showerror("错误", error_msg))
    
    def clear_log(self):
        """清空日志文本框"""
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')
        
    def start_extraction(self):
        """在新线程中开始提取"""
        thread = threading.Thread(target=self.extract_images, daemon=True)
        thread.start()
    
    def check_network_connection(self):
        """检查网络连接"""
        # 创建检测中的提示窗口
        check_window = tk.Tk()
        check_window.title("网络检测")
        check_window.geometry("300x100")
        check_window.resizable(False, False)
        
        # 居中显示窗口
        check_window.eval('tk::PlaceWindow . center')
        
        # 添加检测中的标签
        label = tk.Label(check_window, text="企业环境检测中...", font=('Microsoft YaHei UI', 12))
        label.pack(expand=True)
        
        # 更新窗口显示
        check_window.update()
        
        try:
            # 创建socket连接测试10.0.182.21:22
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(3)  # 设置3秒超时
            result = sock.connect_ex(('10.0.182.21', 22))
            sock.close()
            
            # 关闭检测窗口
            check_window.destroy()
            
            if result == 0:
                return True
            else:
                self.show_network_error()
                return False
        except Exception:
            # 关闭检测窗口
            check_window.destroy()
            self.show_network_error()
            return False
    
    def show_network_error(self):
        """显示网络错误提示窗口"""
        try:
            # 创建一个临时的根窗口用于显示错误消息
            temp_root = tk.Tk()
            temp_root.withdraw()  # 隐藏主窗口
            
            # 显示错误消息
            messagebox.showerror(
                "网络连接错误",
                "企业工具请在企业网络环境使用！",
                parent=temp_root
            )
            
            # 销毁临时窗口
            temp_root.destroy()
        except Exception:
            # 如果GUI创建失败，直接退出
            pass
        finally:
            # 确保程序退出
            sys.exit(1)
    
    def run(self):
        """运行GUI"""
        if hasattr(self, 'root'):
            self.logger.info("捕获Excel图片工具启动")
            self.root.mainloop()


if __name__ == "__main__":
    app = ExcelImageExtractorGUI()
    app.run()