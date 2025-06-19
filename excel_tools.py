import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd

class ExcelSplitterApp:
    def __init__(self, root):
        """初始化Excel工具界面"""
        self.root = root
        self.root.title("Excel工具")
        
        # 初始化变量
        self.input_file = tk.StringVar()
        self.split_count = tk.IntVar()
        self.output_path = tk.StringVar()
        self.enable_split = tk.BooleanVar(value=True)
        self.enable_column_select = tk.BooleanVar(value=True)
        self.selected_columns = []  # 存储用户选择的列
        self.column_vars = {}  # 存储列选择状态
        self.column_mappings = {}  # 存储列映射关系
        self.mapping_entries = {}  # 新增：存储每列的映射Entry控件
        
        # 创建界面组件
        self.create_widgets()
    
    def create_widgets(self):
        """创建界面布局"""
        # 输入文件选择（放在最上方）
        tk.Label(self.root, text="选择Excel文件:").grid(row=0, column=0, sticky="e", padx=10, pady=5)
        tk.Entry(self.root, textvariable=self.input_file, width=40).grid(row=0, column=1, padx=5, pady=5)
        tk.Button(self.root, text="浏览...", command=self.select_input_file).grid(row=0, column=2, padx=5, pady=5)
        
        # 功能开关区域
        self.features_frame = tk.LabelFrame(self.root, text="功能设置")
        self.features_frame.grid(row=1, column=0, columnspan=3, padx=10, pady=5, sticky="ew")
        
        # 拆分功能开关
        tk.Checkbutton(self.features_frame, text="启用拆分功能", variable=self.enable_split).grid(row=0, column=0, sticky="w", padx=10, pady=5)
        
        # 列选择功能开关
        self.enable_column_select = tk.BooleanVar(value=True)
        tk.Checkbutton(self.features_frame, text="启用列选择", variable=self.enable_column_select, 
                      command=self.toggle_column_selection).grid(row=0, column=1, sticky="w", padx=10, pady=5)
        
        # 列选择框架
        self.columns_frame = tk.LabelFrame(self.root, text="选择要保留的列")
        self.columns_frame.grid(row=2, column=0, columnspan=3, padx=10, pady=5, sticky="ew")
        
        # 表达式输入框
        tk.Label(self.columns_frame, text="表达式(如a,b,c):").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.column_expr = tk.StringVar()
        self.expr_entry = tk.Entry(self.columns_frame, textvariable=self.column_expr, width=30)
        self.expr_entry.grid(row=0, column=1, sticky="w", padx=5, pady=2)
        self.expr_entry.bind('<KeyRelease>', lambda e: self.apply_column_expr())
        tk.Button(self.columns_frame, text="应用", command=self.apply_column_expr).grid(row=0, column=2, padx=5, pady=2)
        
        # 字段映射输入框
        tk.Label(self.columns_frame, text="字段映射(如a:a1,b:b1):").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.column_mapping = tk.StringVar()
        self.mapping_entry = tk.Entry(self.columns_frame, textvariable=self.column_mapping, width=30)
        self.mapping_entry.grid(row=1, column=1, sticky="w", padx=5, pady=2)
        tk.Button(self.columns_frame, text="应用", command=self.apply_column_mapping).grid(row=1, column=2, padx=5, pady=2)
        
        # 列选择区域
        self.checkboxes_frame = tk.Frame(self.columns_frame)
        self.checkboxes_frame.grid(row=2, column=0, columnspan=3, sticky="ew")
        
        # 拆分数量
        tk.Label(self.root, text="拆分数量:").grid(row=3, column=0, sticky="e", padx=10, pady=5)
        tk.Entry(self.root, textvariable=self.split_count, width=10).grid(row=3, column=1, sticky="w", padx=5, pady=5)
        
        # 输出路径
        tk.Label(self.root, text="输出路径:").grid(row=4, column=0, sticky="e", padx=10, pady=5)
        tk.Entry(self.root, textvariable=self.output_path, width=40).grid(row=4, column=1, padx=5, pady=5)
        tk.Button(self.root, text="浏览...", command=self.select_output_path).grid(row=4, column=2, padx=5, pady=5)
        
        # 执行按钮
        tk.Button(self.root, text="执行", command=self.execute, bg="#4CAF50", fg="white").grid(row=5, column=1, pady=10)
    
    def select_input_file(self):
        """选择输入Excel文件并读取表头"""
        file_path = filedialog.askopenfilename(filetypes=[("Excel文件", "*.xlsx")])
        if file_path:
            self.input_file.set(file_path)
            # 设置默认输出路径为输入文件所在目录
            self.output_path.set(os.path.dirname(file_path))
            
            # 读取Excel表头
            try:
                df = pd.read_excel(file_path, nrows=1)
                self.selected_columns = list(df.columns)
                self.column_vars = {}
                self.column_mappings = {col: col for col in df.columns}
                self.mapping_entries = {}  # 新增：初始化mapping_entries
                
                # 清空之前的列选择框
                for widget in self.checkboxes_frame.winfo_children():
                    widget.destroy()
                
                # 创建多选框
                for i, column in enumerate(df.columns):
                    var = tk.BooleanVar(value=True)
                    self.column_vars[column] = var
                    
                    # 创建包含列名和映射输入框的行
                    row_frame = tk.Frame(self.checkboxes_frame)
                    row_frame.grid(row=i, column=0, sticky="w")
                    
                    cb = tk.Checkbutton(row_frame, text=column, variable=var,
                                      command=lambda c=column, v=var: self.update_selected_columns(c, v))
                    cb.grid(row=0, column=0, sticky="w", padx=5, pady=2)
                    
                    tk.Label(row_frame, text="映射为:").grid(row=0, column=1, padx=(10,5))
                    mapping_entry = tk.Entry(row_frame, width=20)
                    mapping_entry.insert(0, column)
                    mapping_entry.grid(row=0, column=2, sticky="w")
                    mapping_entry.bind("<FocusOut>", lambda e, c=column: self.update_column_mapping(c, e.widget.get()))
                    self.mapping_entries[column] = mapping_entry  # 新增：保存Entry控件
                    
            except Exception as e:
                messagebox.showerror("错误", f"读取Excel文件表头失败: {str(e)}")
    
    def select_output_path(self):
        """选择输出目录"""
        dir_path = filedialog.askdirectory()
        if dir_path:
            self.output_path.set(dir_path)
    
    def toggle_column_selection(self):
        """切换列选择功能的显示状态"""
        if self.enable_column_select.get():
            self.columns_frame.grid()
        else:
            self.columns_frame.grid_remove()
    
    def apply_column_expr(self):
        """应用列选择表达式"""
        expr = self.column_expr.get().strip()
        if not expr:
            return
            
        try:
            columns = [col.strip() for col in expr.split(",") if col.strip()]
            for col, var in self.column_vars.items():
                var.set(col in columns)
                self.update_selected_columns(col, var)
        except Exception as e:
            messagebox.showerror("错误", f"解析表达式失败: {str(e)}")
    
    def apply_column_mapping(self):
        """应用字段映射"""
        mapping_str = self.column_mapping.get().strip()
        if not mapping_str:
            return
            
        try:
            mappings = [m.strip() for m in mapping_str.split(",") if m.strip()]
            for mapping in mappings:
                if ":" in mapping:
                    src, dst = mapping.split(":", 1)
                    src = src.strip()
                    dst = dst.strip()
                    if src in self.column_vars:
                        self.column_mappings[src] = dst
            # 新增：同步更新每个Entry控件内容
            for col, entry in self.mapping_entries.items():
                entry.delete(0, tk.END)
                entry.insert(0, self.column_mappings.get(col, col))
        except Exception as e:
            messagebox.showerror("错误", f"解析映射失败: {str(e)}")
    
    def update_column_mapping(self, column, new_name):
        """更新单个列的映射关系"""
        if column in self.column_mappings:
            self.column_mappings[column] = new_name.strip() or column
    
    def update_selected_columns(self, column, var):
        """更新用户选择的列"""
        if var.get() and column not in self.selected_columns:
            self.selected_columns.append(column)
        elif not var.get() and column in self.selected_columns:
            self.selected_columns.remove(column)
    
    def execute(self):
        """执行拆分操作"""
        if not self.enable_split.get():
            messagebox.showinfo("提示", "拆分功能未启用")
            return
            
        try:
            input_path = self.input_file.get()
            if not input_path:
                messagebox.showerror("错误", "请选择Excel文件")
                return
                
            split_count = self.split_count.get()
            if split_count <= 0:
                messagebox.showerror("错误", "拆分数量必须大于0")
                return
                
            output_dir = self.output_path.get() or os.path.dirname(input_path)
            
            # 读取Excel文件
            df = pd.read_excel(input_path)
            
            # 只保留用户选择的列并应用映射
            if hasattr(self, 'selected_columns') and self.selected_columns:
                selected_cols = [col for col in self.selected_columns if col in df.columns]
                df = df[selected_cols]
                
                # 应用列名映射
                if hasattr(self, 'column_mappings'):
                    df = df.rename(columns={k:v for k,v in self.column_mappings.items() 
                                           if k in df.columns and v})
                
            total_rows = len(df)
            chunk_size = total_rows // split_count
            
            # 获取文件名和扩展名
            file_name = os.path.splitext(os.path.basename(input_path))[0]
            file_ext = os.path.splitext(input_path)[1]
            
            # 拆分并保存文件
            for i in range(split_count):
                start = i * chunk_size
                end = (i + 1) * chunk_size if i < split_count - 1 else total_rows
                chunk = df.iloc[start:end]
                
                output_path = os.path.join(output_dir, f"{file_name}_{i+1}{file_ext}")
                chunk.to_excel(output_path, index=False)
            
            messagebox.showinfo("成功", f"Excel文件已成功拆分为{split_count}个文件")
        except Exception as e:
            messagebox.showerror("错误", f"处理过程中出现错误:\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelSplitterApp(root)
    root.mainloop()