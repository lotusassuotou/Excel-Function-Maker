#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel函数制作器 / Excel Function Maker
一个用于生成常用Excel函数的图形界面工具
A GUI tool for generating common Excel functions
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import pyperclip
from typing import Dict, List, Callable
import re
import json
import os


class LanguageManager:
    """语言管理器"""
    
    def __init__(self):
        self.current_language = "zh"  # 默认中文
        self.config_file = "language_config.json"
        self.load_language_preference()
        
        # 文本字典
        self.texts = {
            "zh": {
                "title": "Excel函数制作器 v2.0",
                "select_function": "选择函数",
                "function_type": "函数类型:",
                "please_select": "请选择一个函数",
                "parameter_settings": "参数设置",
                "parameter": "参数",
                "generated_function": "生成的函数",
                "example": "示例: ",
                "generate_function": "生成函数",
                "copy_to_clipboard": "复制到剪贴板",
                "clear": "清空",
                "success": "成功",
                "error": "错误",
                "copy_success": "函数已复制到剪贴板！",
                "generate_error": "生成函数时出错: ",
                "startup_error": "应用程序启动失败: ",
                "please_input_chars": "请输入要统计的字符",
                "language": "语言",
                "chinese": "中文",
                "english": "English"
            },
            "en": {
                "title": "Excel Function Maker v2.0",
                "select_function": "Select Function",
                "function_type": "Function Type:",
                "please_select": "Please select a function",
                "parameter_settings": "Parameter Settings",
                "parameter": "Parameter",
                "generated_function": "Generated Function",
                "example": "Example: ",
                "generate_function": "Generate Function",
                "copy_to_clipboard": "Copy to Clipboard",
                "clear": "Clear",
                "success": "Success",
                "error": "Error",
                "copy_success": "Function copied to clipboard!",
                "generate_error": "Error generating function: ",
                "startup_error": "Application startup failed: ",
                "please_input_chars": "Please input characters to count",
                "language": "Language",
                "chinese": "中文",
                "english": "English"
            }
        }
        
        # 函数定义字典
        self.functions = {
            "zh": {
                "SUM": {
                    "name": "求和函数",
                    "description": "计算选定单元格范围的总和",
                    "params": ["单元格范围 (如: A1:A10)"],
                    "template": "=SUM({0})",
                    "example": "=SUM(A1:A10)"
                },
                "AVERAGE": {
                    "name": "平均值函数",
                    "description": "计算选定单元格范围的平均值",
                    "params": ["单元格范围 (如: A1:A10)"],
                    "template": "=AVERAGE({0})",
                    "example": "=AVERAGE(A1:A10)"
                },
                "COUNT": {
                    "name": "计数函数",
                    "description": "计算包含数字的单元格数量",
                    "params": ["单元格范围 (如: A1:A10)"],
                    "template": "=COUNT({0})",
                    "example": "=COUNT(A1:A10)"
                },
                "COUNTIF": {
                    "name": "条件计数函数",
                    "description": "计算满足特定条件的单元格数量",
                    "params": ["单元格范围 (如: A1:A10)", "条件 (如: \">10\" 或 \"苹果\")"],
                    "template": "=COUNTIF({0},{1})",
                    "example": "=COUNTIF(A1:A10,\">10\")"
                },
                "SUMIF": {
                    "name": "条件求和函数",
                    "description": "对满足特定条件的单元格求和",
                    "params": ["条件范围 (如: A1:A10)", "条件 (如: \">10\")", "求和范围 (如: B1:B10，可选)"],
                    "template": "=SUMIF({0},{1},{2})",
                    "example": "=SUMIF(A1:A10,\">10\",B1:B10)"
                },
                "VLOOKUP": {
                    "name": "垂直查找函数",
                    "description": "在表格中垂直查找指定值",
                    "params": ["查找值", "查找表格范围 (如: A1:D10)", "返回列号 (数字)", "精确匹配 (TRUE/FALSE)"],
                    "template": "=VLOOKUP({0},{1},{2},{3})",
                    "example": "=VLOOKUP(\"张三\",A1:D10,3,FALSE)"
                },
                "HLOOKUP": {
                    "name": "水平查找函数",
                    "description": "在表格中水平查找指定值",
                    "params": ["查找值", "查找表格范围 (如: A1:J4)", "返回行号 (数字)", "精确匹配 (TRUE/FALSE)"],
                    "template": "=HLOOKUP({0},{1},{2},{3})",
                    "example": "=HLOOKUP(\"销售额\",A1:J4,3,FALSE)"
                },
                "IF": {
                    "name": "条件判断函数",
                    "description": "根据条件返回不同的值",
                    "params": ["条件表达式 (如: A1>10)", "条件为真时的值", "条件为假时的值"],
                    "template": "=IF({0},{1},{2})",
                    "example": "=IF(A1>10,\"合格\",\"不合格\")"
                },
                "CONCATENATE": {
                    "name": "文本连接函数",
                    "description": "将多个文本字符串连接成一个",
                    "params": ["文本1", "文本2", "文本3 (可选)", "文本4 (可选)"],
                    "template": "=CONCATENATE({0},{1},{2},{3})",
                    "example": "=CONCATENATE(A1,\" \",B1)"
                },
                "LEFT": {
                    "name": "左取字符函数",
                    "description": "从文本左侧提取指定数量的字符",
                    "params": ["文本", "字符数量"],
                    "template": "=LEFT({0},{1})",
                    "example": "=LEFT(A1,5)"
                },
                "RIGHT": {
                    "name": "右取字符函数",
                    "description": "从文本右侧提取指定数量的字符",
                    "params": ["文本", "字符数量"],
                    "template": "=RIGHT({0},{1})",
                    "example": "=RIGHT(A1,3)"
                },
                "MID": {
                    "name": "中间取字符函数",
                    "description": "从文本中间提取指定数量的字符",
                    "params": ["文本", "开始位置", "字符数量"],
                    "template": "=MID({0},{1},{2})",
                    "example": "=MID(A1,2,5)"
                },
                "CHAR_COUNT": {
                    "name": "字符计数统计函数",
                    "description": "统计指定范围内多个字符的出现次数并格式化输出",
                    "params": ["统计范围 (如: A1:J1)", "字符1", "字符2 (可选)", "字符3 (可选)", "字符4 (可选)"],
                    "template": "=\"{1}\"&SUMPRODUCT(LEN({0})-LEN(SUBSTITUTE({0},\"{1}\",\"\")))&\"{2}\"&SUMPRODUCT(LEN({0})-LEN(SUBSTITUTE({0},\"{2}\",\"\")))&\"{3}\"&SUMPRODUCT(LEN({0})-LEN(SUBSTITUTE({0},\"{3}\",\"\")))&\"{4}\"&SUMPRODUCT(LEN({0})-LEN(SUBSTITUTE({0},\"{4}\",\"\")))",
                    "example": "=\"生\"&SUMPRODUCT(LEN(A1:J1)-LEN(SUBSTITUTE(A1:J1,\"生\",\"\")))&\"库\"&SUMPRODUCT(LEN(A1:J1)-LEN(SUBSTITUTE(A1:J1,\"库\",\"\")))&\"旺\"&SUMPRODUCT(LEN(A1:J1)-LEN(SUBSTITUTE(A1:J1,\"旺\",\"\")))"
                }
            },
            "en": {
                "SUM": {
                    "name": "Sum Function",
                    "description": "Calculate the sum of selected cell range",
                    "params": ["Cell range (e.g.: A1:A10)"],
                    "template": "=SUM({0})",
                    "example": "=SUM(A1:A10)"
                },
                "AVERAGE": {
                    "name": "Average Function",
                    "description": "Calculate the average of selected cell range",
                    "params": ["Cell range (e.g.: A1:A10)"],
                    "template": "=AVERAGE({0})",
                    "example": "=AVERAGE(A1:A10)"
                },
                "COUNT": {
                    "name": "Count Function",
                    "description": "Count cells containing numbers",
                    "params": ["Cell range (e.g.: A1:A10)"],
                    "template": "=COUNT({0})",
                    "example": "=COUNT(A1:A10)"
                },
                "COUNTIF": {
                    "name": "Conditional Count Function",
                    "description": "Count cells meeting specific criteria",
                    "params": ["Cell range (e.g.: A1:A10)", "Criteria (e.g.: \">10\" or \"apple\")"],
                    "template": "=COUNTIF({0},{1})",
                    "example": "=COUNTIF(A1:A10,\">10\")"
                },
                "SUMIF": {
                    "name": "Conditional Sum Function",
                    "description": "Sum cells meeting specific criteria",
                    "params": ["Criteria range (e.g.: A1:A10)", "Criteria (e.g.: \">10\")", "Sum range (e.g.: B1:B10, optional)"],
                    "template": "=SUMIF({0},{1},{2})",
                    "example": "=SUMIF(A1:A10,\">10\",B1:B10)"
                },
                "VLOOKUP": {
                    "name": "Vertical Lookup Function",
                    "description": "Look up values vertically in a table",
                    "params": ["Lookup value", "Table array (e.g.: A1:D10)", "Column index (number)", "Exact match (TRUE/FALSE)"],
                    "template": "=VLOOKUP({0},{1},{2},{3})",
                    "example": "=VLOOKUP(\"John\",A1:D10,3,FALSE)"
                },
                "HLOOKUP": {
                    "name": "Horizontal Lookup Function",
                    "description": "Look up values horizontally in a table",
                    "params": ["Lookup value", "Table array (e.g.: A1:J4)", "Row index (number)", "Exact match (TRUE/FALSE)"],
                    "template": "=HLOOKUP({0},{1},{2},{3})",
                    "example": "=HLOOKUP(\"Sales\",A1:J4,3,FALSE)"
                },
                "IF": {
                    "name": "Conditional Function",
                    "description": "Return different values based on condition",
                    "params": ["Condition (e.g.: A1>10)", "Value if true", "Value if false"],
                    "template": "=IF({0},{1},{2})",
                    "example": "=IF(A1>10,\"Pass\",\"Fail\")"
                },
                "CONCATENATE": {
                    "name": "Text Concatenation Function",
                    "description": "Join multiple text strings",
                    "params": ["Text1", "Text2", "Text3 (optional)", "Text4 (optional)"],
                    "template": "=CONCATENATE({0},{1},{2},{3})",
                    "example": "=CONCATENATE(A1,\" \",B1)"
                },
                "LEFT": {
                    "name": "Left Characters Function",
                    "description": "Extract characters from the left of text",
                    "params": ["Text", "Number of characters"],
                    "template": "=LEFT({0},{1})",
                    "example": "=LEFT(A1,5)"
                },
                "RIGHT": {
                    "name": "Right Characters Function",
                    "description": "Extract characters from the right of text",
                    "params": ["Text", "Number of characters"],
                    "template": "=RIGHT({0},{1})",
                    "example": "=RIGHT(A1,3)"
                },
                "MID": {
                    "name": "Middle Characters Function",
                    "description": "Extract characters from the middle of text",
                    "params": ["Text", "Start position", "Number of characters"],
                    "template": "=MID({0},{1},{2})",
                    "example": "=MID(A1,2,5)"
                },
                "CHAR_COUNT": {
                    "name": "Character Count Statistics Function",
                    "description": "Count occurrences of multiple characters in range with formatted output",
                    "params": ["Count range (e.g.: A1:J1)", "Character1", "Character2 (optional)", "Character3 (optional)", "Character4 (optional)"],
                    "template": "=\"{1}\"&SUMPRODUCT(LEN({0})-LEN(SUBSTITUTE({0},\"{1}\",\"\")))&\"{2}\"&SUMPRODUCT(LEN({0})-LEN(SUBSTITUTE({0},\"{2}\",\"\")))&\"{3}\"&SUMPRODUCT(LEN({0})-LEN(SUBSTITUTE({0},\"{3}\",\"\")))&\"{4}\"&SUMPRODUCT(LEN({0})-LEN(SUBSTITUTE({0},\"{4}\",\"\")))",
                    "example": "=\"a\"&SUMPRODUCT(LEN(A1:J1)-LEN(SUBSTITUTE(A1:J1,\"a\",\"\")))&\"b\"&SUMPRODUCT(LEN(A1:J1)-LEN(SUBSTITUTE(A1:J1,\"b\",\"\")))&\"c\"&SUMPRODUCT(LEN(A1:J1)-LEN(SUBSTITUTE(A1:J1,\"c\",\"\")))"
                }
            }
        }
    
    def load_language_preference(self):
        """加载语言偏好设置"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    self.current_language = config.get('language', 'zh')
        except:
            self.current_language = 'zh'
    
    def save_language_preference(self):
        """保存语言偏好设置"""
        try:
            config = {'language': self.current_language}
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except:
            pass
    
    def get_text(self, key):
        """获取当前语言的文本"""
        return self.texts.get(self.current_language, {}).get(key, key)
    
    def get_functions(self):
        """获取当前语言的函数定义"""
        return self.functions.get(self.current_language, {})
    
    def switch_language(self):
        """切换语言"""
        self.current_language = "en" if self.current_language == "zh" else "zh"
        self.save_language_preference()


class ExcelFunctionMaker:
    """Excel函数制作器主类"""
    
    def __init__(self):
        self.lang_manager = LanguageManager()
        self.root = tk.Tk()
        self.setup_window()
        self.setup_ui()
        
    def setup_window(self):
        """设置窗口"""
        self.root.title(self.lang_manager.get_text("title"))
        self.root.geometry("850x650")
        self.root.resizable(True, True)
        self.root.configure(bg='#f0f0f0')
        
    def setup_ui(self):
        """设置用户界面"""
        # 清空现有界面
        for widget in self.root.winfo_children():
            widget.destroy()
            
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(4, weight=1)
        
        # 顶部框架（标题和语言切换）
        top_frame = ttk.Frame(main_frame)
        top_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 20))
        top_frame.columnconfigure(0, weight=1)
        
        # 标题
        title_label = ttk.Label(top_frame, text=self.lang_manager.get_text("title"), 
                               font=("Microsoft YaHei", 16, "bold"))
        title_label.grid(row=0, column=0, sticky=tk.W)
        
        # 语言切换按钮
        language_frame = ttk.Frame(top_frame)
        language_frame.grid(row=0, column=1, sticky=tk.E)
        
        ttk.Label(language_frame, text=self.lang_manager.get_text("language") + ":").pack(side=tk.LEFT, padx=(0, 5))
        
        self.language_var = tk.StringVar()
        language_combo = ttk.Combobox(language_frame, textvariable=self.language_var, 
                                     values=[self.lang_manager.get_text("chinese"), 
                                            self.lang_manager.get_text("english")], 
                                     state="readonly", width=10)
        language_combo.pack(side=tk.LEFT)
        language_combo.bind('<<ComboboxSelected>>', self.on_language_changed)
        
        # 设置当前语言显示
        if self.lang_manager.current_language == "zh":
            language_combo.set(self.lang_manager.get_text("chinese"))
        else:
            language_combo.set(self.lang_manager.get_text("english"))
        
        # 函数选择区域
        func_frame = ttk.LabelFrame(main_frame, text=self.lang_manager.get_text("select_function"), padding="10")
        func_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        func_frame.columnconfigure(1, weight=1)
        
        ttk.Label(func_frame, text=self.lang_manager.get_text("function_type")).grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.function_var = tk.StringVar()
        self.function_combo = ttk.Combobox(func_frame, textvariable=self.function_var, 
                                          values=list(self.lang_manager.get_functions().keys()), 
                                          state="readonly", width=30)
        self.function_combo.grid(row=0, column=1, sticky=(tk.W, tk.E))
        self.function_combo.bind('<<ComboboxSelected>>', self.on_function_selected)
        
        # 函数描述标签
        self.description_label = ttk.Label(func_frame, text=self.lang_manager.get_text("please_select"), 
                                          foreground="blue", font=("Microsoft YaHei", 9))
        self.description_label.grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=(5, 0))
        
        # 参数输入区域
        self.param_frame = ttk.LabelFrame(main_frame, text=self.lang_manager.get_text("parameter_settings"), padding="10")
        self.param_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        self.param_frame.columnconfigure(1, weight=1)
        
        # 参数输入控件列表
        self.param_entries = []
        self.param_labels = []
        self.param_desc_labels = []
        
        # 结果显示区域
        result_frame = ttk.LabelFrame(main_frame, text=self.lang_manager.get_text("generated_function"), padding="10")
        result_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        result_frame.columnconfigure(0, weight=1)
        result_frame.rowconfigure(1, weight=1)
        
        # 示例标签
        self.example_label = ttk.Label(result_frame, text=self.lang_manager.get_text("example"), 
                                      font=("Microsoft YaHei", 9), foreground="gray")
        self.example_label.grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        
        # 结果文本框
        self.result_text = scrolledtext.ScrolledText(result_frame, height=8, width=70,
                                                    font=("Consolas", 11))
        self.result_text.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=2, pady=(10, 0))
        
        # 生成按钮
        self.generate_btn = ttk.Button(button_frame, text=self.lang_manager.get_text("generate_function"), 
                                      command=self.generate_function, state="disabled")
        self.generate_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 复制按钮
        self.copy_btn = ttk.Button(button_frame, text=self.lang_manager.get_text("copy_to_clipboard"), 
                                  command=self.copy_to_clipboard, state="disabled")
        self.copy_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 清空按钮
        self.clear_btn = ttk.Button(button_frame, text=self.lang_manager.get_text("clear"), 
                                   command=self.clear_all)
        self.clear_btn.pack(side=tk.LEFT)
        
        # 设置默认选择
        functions = self.lang_manager.get_functions()
        if functions:
            self.function_combo.set(list(functions.keys())[0])
            self.on_function_selected(None)
    
    def on_language_changed(self, event):
        """语言切换回调"""
        self.lang_manager.switch_language()
        self.root.title(self.lang_manager.get_text("title"))
        self.setup_ui()
    
    def on_function_selected(self, event):
        """当函数被选择时的回调"""
        selected_func = self.function_var.get()
        functions = self.lang_manager.get_functions()
        
        if selected_func in functions:
            func_info = functions[selected_func]
            
            # 更新描述
            self.description_label.config(text=f"{func_info['name']}: {func_info['description']}")
            
            # 更新示例
            self.example_label.config(text=f"{self.lang_manager.get_text('example')}{func_info['example']}")
            
            # 清空现有参数输入框
            for entry in self.param_entries:
                entry.destroy()
            for label in self.param_labels:
                label.destroy()
            for desc_label in self.param_desc_labels:
                desc_label.destroy()
            
            self.param_entries.clear()
            self.param_labels.clear()
            self.param_desc_labels.clear()
            
            # 创建新的参数输入框
            for i, param_desc in enumerate(func_info['params']):
                label = ttk.Label(self.param_frame, text=f"{self.lang_manager.get_text('parameter')}{i+1}:")
                label.grid(row=i, column=0, sticky=tk.W, pady=2, padx=(0, 10))
                self.param_labels.append(label)
                
                entry = ttk.Entry(self.param_frame, width=50)
                entry.grid(row=i, column=1, sticky=(tk.W, tk.E), pady=2)
                entry.bind('<KeyRelease>', self.on_param_change)
                self.param_entries.append(entry)
                
                # 添加参数说明
                desc_label = ttk.Label(self.param_frame, text=param_desc, 
                                     font=("Microsoft YaHei", 8), foreground="gray")
                desc_label.grid(row=i, column=2, sticky=tk.W, padx=(10, 0), pady=2)
                self.param_desc_labels.append(desc_label)
            
            # 启用生成按钮
            self.generate_btn.config(state="normal")
            
            # 清空结果
            self.result_text.delete('1.0', tk.END)
            self.copy_btn.config(state="disabled")
    
    def on_param_change(self, event):
        """参数输入改变时自动生成预览"""
        self.generate_function()
    
    def generate_function(self):
        """生成Excel函数"""
        selected_func = self.function_var.get()
        functions = self.lang_manager.get_functions()
        
        if selected_func not in functions:
            return
        
        func_info = functions[selected_func]
        template = func_info['template']
        
        # 获取参数值
        params = []
        for entry in self.param_entries:
            value = entry.get().strip()
            if value:
                # 对文本参数自动添加引号（如果需要）
                if not (value.startswith('"') and value.endswith('"')) and \
                   not (value.startswith("'") and value.endswith("'")) and \
                   not self.is_cell_reference(value) and \
                   not value.replace('.', '').replace('-', '').isdigit() and \
                   not value.upper() in ['TRUE', 'FALSE']:
                    if selected_func in ['COUNTIF', 'SUMIF'] and len(params) == 1:
                        # 条件参数需要引号
                        value = f'"{value}"'
                    elif selected_func in ['VLOOKUP', 'HLOOKUP'] and len(params) == 0:
                        # 查找值如果是文本需要引号
                        value = f'"{value}"'
                    elif selected_func == 'IF' and len(params) > 0:
                        # IF函数的返回值如果是文本需要引号
                        value = f'"{value}"'
                    elif selected_func in ['CONCATENATE', 'LEFT', 'RIGHT', 'MID']:
                        # 文本函数的文本参数可能需要引号
                        if len(params) == 0 and not self.is_cell_reference(value):
                            value = f'"{value}"'
            params.append(value)
        
        # 生成函数
        try:
            # 处理可选参数
            while len(params) < template.count('{'):
                params.append('')
            
            # 过滤空参数（除了最后一个必需参数）
            if selected_func in ['SUMIF', 'CONCATENATE']:
                # 这些函数有可选参数
                non_empty_params = []
                for i, param in enumerate(params):
                    if param or i < len(func_info['params']) - (1 if selected_func == 'SUMIF' else 2):
                        non_empty_params.append(param)
                params = non_empty_params
            
            # 格式化函数
            if selected_func == 'SUMIF' and len(params) == 2:
                # SUMIF可以只有两个参数
                function_result = f"=SUMIF({params[0]},{params[1]})"
            elif selected_func == 'CONCATENATE':
                # CONCATENATE可以有可变数量的参数
                non_empty = [p for p in params if p]
                function_result = f"=CONCATENATE({','.join(non_empty)})"
            elif selected_func == 'CHAR_COUNT':
                # 字符计数统计函数的特殊处理
                range_param = params[0] if params[0] else 'A1:J1'
                char_parts = []
                
                # 处理输入的字符，过滤空字符
                chars = [p.strip().strip('"').strip("'") for p in params[1:] if p.strip()]
                
                for char in chars:
                    if char:
                        char_parts.append(f'"{char}"&SUMPRODUCT(LEN({range_param})-LEN(SUBSTITUTE({range_param},"{char}","")))')
                
                if char_parts:
                    function_result = f"={char_parts[0]}"
                    for i in range(1, len(char_parts)):
                        function_result += f"&{char_parts[i]}"
                else:
                    function_result = f'="{self.lang_manager.get_text("please_input_chars")}"'
            else:
                function_result = template.format(*params)
            
            # 显示结果
            self.result_text.delete('1.0', tk.END)
            self.result_text.insert('1.0', function_result)
            
            # 启用复制按钮
            self.copy_btn.config(state="normal")
            
        except Exception as e:
            messagebox.showerror(self.lang_manager.get_text("error"), 
                               f"{self.lang_manager.get_text('generate_error')}{str(e)}")
    
    def is_cell_reference(self, text):
        """检查是否是单元格引用"""
        # 匹配单元格引用模式，如 A1, $A$1, A1:B10, Sheet1!A1 等
        pattern = r'^(\$?[A-Z]+\$?\d+(:?\$?[A-Z]+\$?\d+)?|[A-Za-z0-9_]+!\$?[A-Z]+\$?\d+(:?\$?[A-Z]+\$?\d+)?)$'
        return bool(re.match(pattern, text))
    
    def copy_to_clipboard(self):
        """复制结果到剪贴板"""
        result = self.result_text.get('1.0', tk.END).strip()
        if result:
            pyperclip.copy(result)
            messagebox.showinfo(self.lang_manager.get_text("success"), 
                              self.lang_manager.get_text("copy_success"))
    
    def clear_all(self):
        """清空所有输入和结果"""
        for entry in self.param_entries:
            entry.delete(0, tk.END)
        self.result_text.delete('1.0', tk.END)
        self.copy_btn.config(state="disabled")
    
    def run(self):
        """运行应用程序"""
        self.root.mainloop()


def main():
    """主函数"""
    try:
        app = ExcelFunctionMaker()
        app.run()
    except Exception as e:
        print(f"应用程序启动失败: {e}")
        messagebox.showerror("错误", f"应用程序启动失败: {e}")


if __name__ == "__main__":
    main()
