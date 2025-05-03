import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import os
import threading
from excel_processor import process_excel_files

class ExcelProcessorUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel数据处理工具")
        self.root.geometry("800x650")
        self.root.minsize(780, 600)
        self.root.resizable(True, True)
        
        # 设置主题样式
        style = ttk.Style()
        style.configure("TFrame", background="#f0f0f0")
        style.configure("TButton", font=("微软雅黑", 10))
        style.configure("TLabel", font=("微软雅黑", 10), background="#f0f0f0")
        style.configure("Header.TLabel", font=("微软雅黑", 12, "bold"), background="#f0f0f0")
        style.configure("Result.TLabel", font=("微软雅黑", 10), foreground="green", background="#f0f0f0")
        
        # 自定义处理按钮样式
        style.configure(
            "Process.TButton", 
            font=("微软雅黑", 12, "bold"), 
            padding=10,
            background="#4CAF50",  # 绿色背景
            foreground="#FFFFFF"   # 白色文字
        )
        style.map(
            "Process.TButton",
            background=[("active", "#45a049"), ("pressed", "#398e3c")],  # 鼠标悬停和按下时的颜色
            foreground=[("active", "#FFFFFF"), ("pressed", "#FFFFFF")]
        )
        
        # 用于存储多个A表文件的信息
        self.a_files = []  # 格式: [(文件路径, 工作表名称), ...]
        self.a_common_sheet = tk.StringVar()  # 用于存储通用工作表名称
        
        self.create_widgets()
    
    def create_widgets(self):
        # 创建一个主滚动框架来容纳所有内容
        # 先创建一个Canvas
        main_canvas = tk.Canvas(self.root, bg="#f0f0f0", width=730)
        main_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 为Canvas添加滚动条
        main_scrollbar = ttk.Scrollbar(self.root, orient=tk.VERTICAL, command=main_canvas.yview)
        main_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 配置Canvas
        main_canvas.configure(yscrollcommand=main_scrollbar.set)
        main_canvas.bind('<Configure>', lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all")))
        
        # 在Canvas上创建一个框架以放置所有控件
        scrollable_frame = ttk.Frame(main_canvas, style="TFrame")
        main_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw", width=710)
        
        # 主框架
        main_frame = ttk.Frame(scrollable_frame, padding="20 20 20 20", style="TFrame")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 绑定鼠标滚轮事件
        def _on_mousewheel(event):
            main_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        main_canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # 标题
        header = ttk.Label(main_frame, text="Excel数据处理工具", style="Header.TLabel")
        header.pack(pady=(0, 20))
        
        # 创建三个选项卡(a表、b表、输出)
        tab_control = ttk.Notebook(main_frame)
        
        # A表选项卡
        a_tab = ttk.Frame(tab_control, style="TFrame")
        tab_control.add(a_tab, text="A表 (源数据)")
        
        # B表选项卡
        b_tab = ttk.Frame(tab_control, style="TFrame")
        tab_control.add(b_tab, text="B表 (对比数据)")
        
        # 输出选项卡
        output_tab = ttk.Frame(tab_control, style="TFrame")
        tab_control.add(output_tab, text="输出设置")
        
        tab_control.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
        
        # A表设置
        self.setup_a_tab(a_tab)
        
        # B表设置
        self.setup_b_tab(b_tab)
        
        # 输出设置
        self.setup_output_tab(output_tab)
        
        # 创建大型处理按钮框架
        process_frame = ttk.Frame(main_frame, style="TFrame")
        process_frame.pack(fill=tk.X, pady=(10, 5))
        
        # 由于ttk样式限制，使用纯tk按钮来实现彩色按钮
        process_button = tk.Button(
            process_frame, 
            text="开始处理数据", 
            command=self.process_data,
            font=("微软雅黑", 12, "bold"),
            bg="#4CAF50",  # 绿色背景
            fg="white",    # 白色文字
            activebackground="#45a049",  # 鼠标悬停时的颜色
            activeforeground="white",
            relief=tk.RAISED,
            bd=1,
            padx=20,
            pady=10,
            cursor="hand2"  # 手型光标
        )
        process_button.pack(fill=tk.X, ipady=5, pady=5)
        
        # 状态标签
        self.status_var = tk.StringVar()
        self.status_var.set("准备就绪")
        status_label = ttk.Label(main_frame, textvariable=self.status_var, style="Result.TLabel")
        status_label.pack(pady=(5, 0))
        
        # 结果显示框
        result_frame = ttk.LabelFrame(main_frame, text="处理结果", padding="10 10 10 10")
        result_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        
        # 使用Text控件显示结果
        self.result_text = tk.Text(result_frame, height=6, wrap=tk.WORD)
        self.result_text.pack(fill=tk.BOTH, expand=True)
        
        # 滚动条
        scrollbar = ttk.Scrollbar(self.result_text, command=self.result_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.result_text.config(yscrollcommand=scrollbar.set)
        
        # 设置默认值
        self.set_default_values()
        
        # 确保Canvas可以滚动所有内容
        scrollable_frame.update_idletasks()
        main_canvas.config(scrollregion=main_canvas.bbox("all"))
    
    def setup_a_tab(self, parent):
        # A表文件列表框架
        files_frame = ttk.LabelFrame(parent, text="A表文件列表", padding="10 10 10 10")
        files_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 文件列表显示区域
        self.files_listbox_frame = ttk.Frame(files_frame)
        self.files_listbox_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 创建Treeview用于显示文件列表
        columns = ("序号", "文件路径", "工作表名称")
        self.files_tree = ttk.Treeview(self.files_listbox_frame, columns=columns, show="headings", height=7)
        
        # 设置列宽和标题
        self.files_tree.column("序号", width=50, anchor="center")
        self.files_tree.column("文件路径", width=450)
        self.files_tree.column("工作表名称", width=150, anchor="center")
        
        self.files_tree.heading("序号", text="序号")
        self.files_tree.heading("文件路径", text="文件路径")
        self.files_tree.heading("工作表名称", text="工作表名称")
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(self.files_listbox_frame, orient="vertical", command=self.files_tree.yview)
        self.files_tree.configure(yscrollcommand=scrollbar.set)
        
        self.files_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 文件操作按钮区域
        files_buttons_frame = ttk.Frame(files_frame)
        files_buttons_frame.pack(fill=tk.X, expand=False, pady=5)
        
        # 添加文件按钮
        add_file_button = ttk.Button(files_buttons_frame, text="添加文件", command=self.add_a_file)
        add_file_button.pack(side=tk.LEFT, padx=5)
        
        # 删除选中文件按钮
        remove_file_button = ttk.Button(files_buttons_frame, text="删除选中", command=self.remove_a_file)
        remove_file_button.pack(side=tk.LEFT, padx=5)
        
        # 清空所有文件按钮
        clear_files_button = ttk.Button(files_buttons_frame, text="清空列表", command=self.clear_a_files)
        clear_files_button.pack(side=tk.LEFT, padx=5)
        
        # 设置工作表名按钮
        set_sheet_button = ttk.Button(files_buttons_frame, text="设置工作表名", command=self.set_sheet_name)
        set_sheet_button.pack(side=tk.LEFT, padx=5)
        
        # 通用工作表名称区域
        common_sheet_frame = ttk.Frame(parent)
        common_sheet_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(common_sheet_frame, text="通用工作表名称:", style="TLabel").pack(side=tk.LEFT)
        
        common_sheet_entry = ttk.Entry(common_sheet_frame, textvariable=self.a_common_sheet, width=20)
        common_sheet_entry.pack(side=tk.LEFT, padx=5)
        
        apply_common_button = ttk.Button(common_sheet_frame, text="应用到所有文件", command=self.apply_common_sheet)
        apply_common_button.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(common_sheet_frame, text="(留空使用各文件的默认活动表)", style="TLabel").pack(side=tk.LEFT)
        
        # A表列选择
        col_frame = ttk.Frame(parent, style="TFrame")
        col_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(col_frame, text="比较列 (如A, B, C或1, 2, 3):", style="TLabel").pack(side=tk.LEFT)
        
        self.a_column = tk.StringVar()
        col_entry = ttk.Entry(col_frame, textvariable=self.a_column, width=5)
        col_entry.pack(side=tk.LEFT, padx=5)
        
        # 说明
        info_frame = ttk.LabelFrame(parent, text="说明", style="TFrame")
        info_frame.pack(fill=tk.BOTH, expand=False, pady=10)
        
        info_text = """A表是源数据表，程序将查找此表中指定列与B表匹配的行。
- 文件列表：可以添加多个Excel文件进行批量处理
- 工作表名称：可以为每个文件单独设置工作表名称，也可以使用通用工作表名应用到所有文件
- 比较列：用于与B表比较的列，可以是字母(A,B,C)或数字(1,2,3)
- 对于含有多个工作表的Excel文件，可以单独指定要处理的工作表"""
        
        ttk.Label(info_frame, text=info_text, style="TLabel", wraplength=650, justify=tk.LEFT).pack(pady=5)
    
    def setup_b_tab(self, parent):
        # B表文件路径
        path_frame = ttk.Frame(parent, style="TFrame")
        path_frame.pack(fill=tk.X, pady=(10, 5))
        
        ttk.Label(path_frame, text="B表文件路径:", style="TLabel").pack(side=tk.LEFT)
        
        self.b_file_path = tk.StringVar()
        path_entry = ttk.Entry(path_frame, textvariable=self.b_file_path, width=60)
        path_entry.pack(side=tk.LEFT, padx=5)
        
        browse_button = ttk.Button(path_frame, text="浏览...", command=self.browse_b_file)
        browse_button.pack(side=tk.LEFT)
        
        # B表工作表名称
        sheet_frame = ttk.Frame(parent, style="TFrame")
        sheet_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(sheet_frame, text="工作表名称:", style="TLabel").pack(side=tk.LEFT)
        
        self.b_sheet_name = tk.StringVar()
        sheet_entry = ttk.Entry(sheet_frame, textvariable=self.b_sheet_name, width=20)
        sheet_entry.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(sheet_frame, text="(留空使用默认活动表)", style="TLabel").pack(side=tk.LEFT)
        
        # B表列选择
        col_frame = ttk.Frame(parent, style="TFrame")
        col_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(col_frame, text="比较列 (如A, B, C或1, 2, 3):", style="TLabel").pack(side=tk.LEFT)
        
        self.b_column = tk.StringVar()
        col_entry = ttk.Entry(col_frame, textvariable=self.b_column, width=5)
        col_entry.pack(side=tk.LEFT, padx=5)
        
        # 说明
        info_frame = ttk.LabelFrame(parent, text="说明", style="TFrame")
        info_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        info_text = """B表是对比数据表，程序将查找A表中与此表指定列匹配的行。
- 文件路径：Excel文件的完整路径
- 工作表名称：要处理的工作表名称，留空将使用默认活动表
- 比较列：用于与A表比较的列，可以是字母(A,B,C)或数字(1,2,3)"""
        
        ttk.Label(info_frame, text=info_text, style="TLabel", wraplength=650, justify=tk.LEFT).pack(pady=5)
    
    def setup_output_tab(self, parent):
        # 输出文件夹路径
        path_frame = ttk.Frame(parent, style="TFrame")
        path_frame.pack(fill=tk.X, pady=(10, 5))
        
        ttk.Label(path_frame, text="输出文件夹:", style="TLabel").pack(side=tk.LEFT)
        
        self.output_folder_path = tk.StringVar()
        path_entry = ttk.Entry(path_frame, textvariable=self.output_folder_path, width=60)
        path_entry.pack(side=tk.LEFT, padx=5)
        
        browse_button = ttk.Button(path_frame, text="浏览...", command=self.browse_output_folder)
        browse_button.pack(side=tk.LEFT)
        
        # 输出文件名
        file_frame = ttk.Frame(parent, style="TFrame")
        file_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(file_frame, text="输出文件名:", style="TLabel").pack(side=tk.LEFT)
        
        self.output_file_name = tk.StringVar()
        self.output_file_name.set("匹配结果.xlsx")
        file_entry = ttk.Entry(file_frame, textvariable=self.output_file_name, width=30)
        file_entry.pack(side=tk.LEFT, padx=5)
        
        # 输出工作表名称
        sheet_frame = ttk.Frame(parent, style="TFrame")
        sheet_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(sheet_frame, text="工作表名称:", style="TLabel").pack(side=tk.LEFT)
        
        self.output_sheet_name = tk.StringVar()
        sheet_entry = ttk.Entry(sheet_frame, textvariable=self.output_sheet_name, width=20)
        sheet_entry.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(sheet_frame, text="(默认为'匹配结果')", style="TLabel").pack(side=tk.LEFT)
        
        # 说明
        info_frame = ttk.LabelFrame(parent, text="说明", style="TFrame")
        info_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        info_text = """输出设置决定匹配结果保存的位置。
- 输出文件夹：选择保存结果文件的文件夹
- 输出文件名：结果文件的名称
- 工作表名称：结果将保存在此工作表中
        
注意：
1. 实际保存的文件名会自动添加时间戳，以避免覆盖现有文件
2. 程序会自动将A表中的第一行作为表头复制到结果文件
3. 如果A表中包含公式，只会保存计算结果，不保存公式本身
4. 程序会自动识别并保持原A表中的合并单元格状态
5. 当处理多个A表文件时，所有匹配的行将合并到一个结果文件中"""
        
        ttk.Label(info_frame, text=info_text, style="TLabel", wraplength=650, justify=tk.LEFT).pack(pady=5)
    
    def add_a_file(self):
        """添加A表文件到列表"""
        filenames = filedialog.askopenfilenames(
            title="选择一个或多个A表文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        
        if not filenames:
            return
            
        # 添加文件到列表
        for filename in filenames:
            # 默认使用通用工作表名
            sheet_name = self.a_common_sheet.get() 
            
            # 添加到内部数据结构
            self.a_files.append((filename, sheet_name))
            
            # 更新UI显示
            self.update_a_files_treeview()
    
    def remove_a_file(self):
        """删除选中的A表文件"""
        selected_items = self.files_tree.selection()
        if not selected_items:
            messagebox.showinfo("提示", "请先选择要删除的文件")
            return
        
        # 获取所有选中项的索引
        indices = [int(self.files_tree.item(item, "values")[0]) - 1 for item in selected_items]
        
        # 从大到小排序索引，以便删除时不影响其他索引
        indices.sort(reverse=True)
        
        # 删除对应的文件
        for idx in indices:
            if 0 <= idx < len(self.a_files):
                self.a_files.pop(idx)
        
        # 更新UI显示
        self.update_a_files_treeview()
    
    def clear_a_files(self):
        """清空A表文件列表"""
        if messagebox.askyesno("确认", "确定要清空所有文件吗？"):
            self.a_files = []
            self.update_a_files_treeview()
    
    def set_sheet_name(self):
        """为选中的文件设置工作表名称"""
        selected_items = self.files_tree.selection()
        if not selected_items:
            messagebox.showinfo("提示", "请先选择要设置工作表名的文件")
            return
        
        # 获取新的工作表名
        sheet_name = simpledialog.askstring("设置工作表名", "请输入工作表名称：",
                                            initialvalue=self.a_common_sheet.get())
        
        if sheet_name is not None:  # 用户点击确定（可能输入空字符串）
            # 更新选中文件的工作表名
            for item in selected_items:
                idx = int(self.files_tree.item(item, "values")[0]) - 1
                if 0 <= idx < len(self.a_files):
                    file_path, _ = self.a_files[idx]
                    self.a_files[idx] = (file_path, sheet_name)
            
            # 更新UI显示
            self.update_a_files_treeview()
    
    def apply_common_sheet(self):
        """将通用工作表名应用到所有文件"""
        common_sheet = self.a_common_sheet.get()
        
        # 更新所有文件的工作表名
        for i in range(len(self.a_files)):
            file_path, _ = self.a_files[i]
            self.a_files[i] = (file_path, common_sheet)
        
        # 更新UI显示
        self.update_a_files_treeview()
    
    def update_a_files_treeview(self):
        """更新文件列表显示"""
        # 清空现有项目
        for item in self.files_tree.get_children():
            self.files_tree.delete(item)
        
        # 添加所有文件
        for i, (file_path, sheet_name) in enumerate(self.a_files):
            self.files_tree.insert("", "end", values=(i+1, file_path, sheet_name or "默认"))
    
    def browse_b_file(self):
        filename = filedialog.askopenfilename(
            title="选择B表文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if filename:
            self.b_file_path.set(filename)
    
    def browse_output_folder(self):
        folder = filedialog.askdirectory(
            title="选择输出文件夹"
        )
        if folder:
            self.output_folder_path.set(folder)
    
    def set_default_values(self):
        # 设置一些默认值
        self.a_column.set("C")
        self.b_column.set("A")
        self.output_sheet_name.set("匹配结果")
        
        # 设置默认输出文件夹为当前工作目录
        self.output_folder_path.set(os.getcwd())
    
    def process_data(self):
        # 获取参数
        a_files = self.a_files
        b_file = self.b_file_path.get().strip()
        output_folder = self.output_folder_path.get().strip()
        output_filename = self.output_file_name.get().strip()
        output_file = os.path.join(output_folder, output_filename)
        
        # 获取默认(通用)工作表名
        default_sheet_a = self.a_common_sheet.get().strip() or None
        
        # 创建文件到工作表的映射
        sheet_a_map = {}
        for file_path, sheet_name in a_files:
            if sheet_name:
                sheet_a_map[file_path] = sheet_name
        
        b_sheet = self.b_sheet_name.get().strip() or None
        output_sheet = self.output_sheet_name.get().strip() or "匹配结果"
        a_col = self.a_column.get().strip()
        b_col = self.b_column.get().strip()
        
        # 参数验证
        if not a_files:
            messagebox.showerror("错误", "请添加至少一个A表文件")
            return
        
        # 检查所有A表文件是否存在
        for file_path, _ in a_files:
            if not os.path.exists(file_path):
                messagebox.showerror("错误", f"A表文件不存在: {file_path}")
                return
            
        if not b_file:
            messagebox.showerror("错误", "请选择B表文件")
            return
        if not os.path.exists(b_file):
            messagebox.showerror("错误", f"B表文件不存在: {b_file}")
            return
            
        if not output_folder:
            messagebox.showerror("错误", "请选择输出文件夹")
            return
        if not os.path.exists(output_folder):
            messagebox.showerror("错误", f"输出文件夹不存在: {output_folder}")
            return
            
        if not output_filename:
            messagebox.showerror("错误", "请指定输出文件名")
            return
        
        if not output_filename.lower().endswith('.xlsx'):
            output_file = output_file + '.xlsx'
            
        if not a_col:
            messagebox.showerror("错误", "请指定A表的比较列")
            return
            
        if not b_col:
            messagebox.showerror("错误", "请指定B表的比较列")
            return
        
        # 清空结果
        self.result_text.delete(1.0, tk.END)
        self.status_var.set("正在处理数据...")
        
        # 提取A表文件路径列表
        a_file_paths = [file_path for file_path, _ in a_files]
        
        # 使用线程进行处理，避免界面卡死
        thread = threading.Thread(target=self.do_process, args=(
            a_file_paths, b_file, output_file, a_col, b_col, default_sheet_a, b_sheet, output_sheet, sheet_a_map
        ))
        thread.daemon = True
        thread.start()
    
    def do_process(self, a_files, b_file, output_file, a_col, b_col, default_sheet_a, b_sheet, output_sheet, sheet_a_map):
        try:
            # 执行处理
            count, saved_path = process_excel_files(
                a_files, b_file, output_file, a_col, b_col,
                sheet_a=default_sheet_a, sheet_b=b_sheet, 
                output_sheet=output_sheet, sheet_a_map=sheet_a_map
            )
            
            # 更新结果
            if count > 0 and saved_path:
                self.root.after(0, lambda: self.status_var.set(f"处理完成，找到 {count} 行匹配数据"))
                result_message = f"处理成功！\n\n共处理了 {len(a_files)} 个文件，找到 {count} 行匹配的数据。\n\n结果已保存到文件:\n{saved_path}"
                self.root.after(0, lambda: self.result_text.insert(tk.END, result_message))
                
                # 询问是否打开文件
                if messagebox.askyesno("处理完成", f"找到 {count} 行匹配数据，已保存到\n{saved_path}\n\n是否打开此文件?"):
                    os.startfile(saved_path)
            else:
                self.root.after(0, lambda: self.status_var.set("处理未完成"))
                self.root.after(0, lambda: self.result_text.insert(tk.END, "未找到匹配的数据或保存文件失败。"))
        except Exception as e:
            error_message = f"处理过程中出错:\n{str(e)}"
            self.root.after(0, lambda: self.status_var.set("处理失败"))
            self.root.after(0, lambda: self.result_text.insert(tk.END, error_message))
            self.root.after(0, lambda: messagebox.showerror("错误", error_message))

def main():
    root = tk.Tk()
    app = ExcelProcessorUI(root)
    root.mainloop()

if __name__ == "__main__":
    main() 