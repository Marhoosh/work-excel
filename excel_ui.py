import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog, font as tkfont
import os
import threading
import sys
import importlib.util
import platform

# 处理PyInstaller打包后的资源路径
def resource_path(relative_path):
    """获取资源的绝对路径，处理PyInstaller打包后的路径问题"""
    try:
        # PyInstaller创建临时文件夹并将资源存储在_MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        # 如果不是打包模式，使用正常的相对路径
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)

# 动态导入excel_processor模块
def import_excel_processor():
    """动态导入excel_processor模块，处理打包后的导入问题"""
    try:
        # 先尝试常规导入
        from excel_processor import process_excel_files
        return process_excel_files
    except ImportError:
        # 如果常规导入失败，尝试从打包路径导入
        try:
            processor_path = resource_path("excel_processor.py")
            if os.path.exists(processor_path):
                spec = importlib.util.spec_from_file_location("excel_processor", processor_path)
                module = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(module)
                return module.process_excel_files
            else:
                raise ImportError("找不到excel_processor.py文件")
        except Exception as e:
            messagebox.showerror("错误", f"导入excel_processor模块失败: {str(e)}")
            sys.exit(1)

# 获取process_excel_files函数
process_excel_files = import_excel_processor()

class ExcelProcessorUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel数据处理工具")
        
        # 设置适应不同平台的窗口大小
        system = platform.system()
        if system == "Darwin":  # macOS
            self.root.geometry("850x650")
            self.root.minsize(820, 600)
        else:
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
        
        # 自定义Notebook选项卡样式
        style.configure(
            "Custom.TNotebook", 
            background="#f0f0f0", 
            borderwidth=1,
            tabmargins=[2, 5, 2, 0]
        )
        
        # 选项卡样式
        style.configure(
            "Custom.TNotebook.Tab",
            font=("微软雅黑", 11),
            padding=[30, 10],
            background="#e8e8e8",  # 淡灰色背景
            foreground="black"     # 黑色文字，确保可见性
        )
        
        # 选中选项卡样式
        style.map(
            "Custom.TNotebook.Tab",
            background=[("selected", "#58b957")],  # 明亮的绿色
            foreground=[("selected", "#FFFFFF")],  # 纯白色文字
            font=[("selected", ("微软雅黑", 11, "bold"))]  # 加粗文字增强可见性
        )
        
        # 创建选项卡内容框架的样式
        style.configure(
            "TabContent.TFrame",
            background="#f8f8f8",  # 稍微淡一点的背景色
            relief="solid",       # 实线边框
            borderwidth=1         # 细边框
        )
        
        # 定义分隔线样式
        style.configure("TSeparator", background="#4CAF50")  # 绿色分隔线
        
        # 用于存储多个A表文件的信息
        self.a_files = []  # 格式: [(文件路径, 工作表名称), ...]
        self.a_common_sheet = tk.StringVar()  # 用于存储通用工作表名称
        
        self.create_widgets()
    
    def create_widgets(self):
        # 创建一个主容器框架
        main_container = ttk.Frame(self.root, style="TFrame")
        main_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 先创建一个Canvas，用于实现滚动效果
        main_canvas = tk.Canvas(main_container, bg="#f0f0f0")
        main_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 为Canvas添加垂直滚动条
        y_scrollbar = ttk.Scrollbar(main_container, orient=tk.VERTICAL, command=main_canvas.yview)
        y_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 添加水平滚动条
        x_scrollbar = ttk.Scrollbar(self.root, orient=tk.HORIZONTAL, command=main_canvas.xview)
        x_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 配置Canvas的滚动
        main_canvas.configure(yscrollcommand=y_scrollbar.set, xscrollcommand=x_scrollbar.set)
        main_canvas.bind('<Configure>', lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all")))
        
        # 在Canvas上创建一个框架以放置所有控件
        scrollable_frame = ttk.Frame(main_canvas, style="TFrame")
        
        # 固定内容宽度，确保水平滚动有效
        content_width = 750 if platform.system() != "Darwin" else 800  # Mac系统使用更宽的尺寸
        main_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw", width=content_width)
        
        # 主框架
        main_frame = ttk.Frame(scrollable_frame, padding="20 20 20 20", style="TFrame")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 绑定鼠标滚轮事件
        def _on_mousewheel(event):
            # 根据平台处理滚轮事件，确保在所有系统上都能工作
            if platform.system() == "Windows":
                main_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            elif platform.system() == "Darwin":  # macOS
                main_canvas.yview_scroll(int(-1*(event.delta)), "units")
            else:  # Linux
                if event.num == 4:
                    main_canvas.yview_scroll(-1, "units")
                elif event.num == 5:
                    main_canvas.yview_scroll(1, "units")
                
        # 绑定滚轮事件，兼容不同平台
        if platform.system() == "Windows" or platform.system() == "Darwin":
            main_canvas.bind_all("<MouseWheel>", _on_mousewheel)
        else:
            # Linux
            main_canvas.bind_all("<Button-4>", _on_mousewheel)
            main_canvas.bind_all("<Button-5>", _on_mousewheel)
        
        # 标题
        header = ttk.Label(main_frame, text="Excel数据处理工具", style="Header.TLabel")
        header.pack(pady=(0, 20))
        
        # 创建自定义选项卡的框架
        tab_frames_container = ttk.Frame(main_frame)
        tab_frames_container.pack(fill=tk.BOTH, expand=True, pady=(0, 20), padx=5)
        
        # 创建选项卡按钮框架
        tab_buttons_frame = ttk.Frame(tab_frames_container)
        tab_buttons_frame.pack(fill=tk.X)
        
        # 创建选项卡内容框架
        content_frame = ttk.Frame(tab_frames_container, style="TabContent.TFrame", borderwidth=1, relief="solid")
        content_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 5))
        
        # 创建A/B/输出设置的框架 - 使用统一的背景色
        a_tab = ttk.Frame(content_frame, padding="10", style="TabContent.TFrame")
        b_tab = ttk.Frame(content_frame, padding="10", style="TabContent.TFrame")
        output_tab = ttk.Frame(content_frame, padding="10", style="TabContent.TFrame")
        
        # 当前选中的选项卡
        self.current_tab = tk.StringVar(value="a_tab")  # 默认显示A表
        
        # 创建选项卡按钮样式 - 更美观的样式
        tab_button_style = {
            "font": ("微软雅黑", 11),
            "padx": 30,
            "pady": 10,
            "bd": 1,
            "relief": "flat",  # 平面风格更美观
            "cursor": "hand2",
            "bg": "#e8e8e8",
            "fg": "#333333",
            "activebackground": "#d0d0d0",
            "activeforeground": "#333333"
        }
        
        # 选项卡按钮点击函数
        def show_tab(tab_name):
            # 隐藏所有选项卡
            a_tab.pack_forget()
            b_tab.pack_forget()
            output_tab.pack_forget()
            
            # 重置所有按钮样式
            a_button.config(bg="#e8e8e8", fg="black", font=("微软雅黑", 11))
            b_button.config(bg="#e8e8e8", fg="black", font=("微软雅黑", 11))
            output_button.config(bg="#e8e8e8", fg="black", font=("微软雅黑", 11))
            
            # 显示选中的选项卡
            if tab_name == "a_tab":
                a_tab.pack(fill=tk.BOTH, expand=True)
                a_button.config(bg="#4CAF50", fg="white", font=("微软雅黑", 11, "bold"))
            elif tab_name == "b_tab":
                b_tab.pack(fill=tk.BOTH, expand=True)
                b_button.config(bg="#4CAF50", fg="white", font=("微软雅黑", 11, "bold"))
            else:  # output_tab
                output_tab.pack(fill=tk.BOTH, expand=True)
                output_button.config(bg="#4CAF50", fg="white", font=("微软雅黑", 11, "bold"))
            
            self.current_tab.set(tab_name)
        
        # 创建选项卡按钮
        a_button = tk.Button(tab_buttons_frame, text="日报表", command=lambda: show_tab("a_tab"), **tab_button_style)
        a_button.pack(side=tk.LEFT)
        
        b_button = tk.Button(tab_buttons_frame, text="患者库", command=lambda: show_tab("b_tab"), **tab_button_style)
        b_button.pack(side=tk.LEFT)
        
        output_button = tk.Button(tab_buttons_frame, text="输出设置", command=lambda: show_tab("output_tab"), **tab_button_style)
        output_button.pack(side=tk.LEFT)
        
        # 默认显示A表选项卡
        show_tab("a_tab")
        
        # A表设置
        self.setup_a_tab(a_tab)
        
        # B表设置
        self.setup_b_tab(b_tab)
        
        # 输出设置
        self.setup_output_tab(output_tab)
        
        # 添加分隔线，增强视觉效果
        separator = ttk.Separator(main_frame, orient="horizontal", style="TSeparator")
        separator.pack(fill=tk.X, pady=(0, 10))
        
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
        result_scrollbar = ttk.Scrollbar(self.result_text, command=self.result_text.yview)
        result_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.result_text.config(yscrollcommand=result_scrollbar.set)
        
        # 设置默认值
        self.set_default_values()
        
        # 确保Canvas可以滚动所有内容
        scrollable_frame.update_idletasks()
        main_canvas.config(scrollregion=main_canvas.bbox("all"))
        
        # 添加调整大小事件处理
        self.root.bind("<Configure>", lambda event: self._on_resize(event, main_canvas, scrollable_frame))
    
    def _on_resize(self, event, canvas, frame):
        """处理窗口大小调整"""
        # 获取当前画布的宽度
        canvas_width = event.width - 30  # 减去滚动条和边距的空间
        
        # 确保最小宽度
        min_width = 750 if platform.system() != "Darwin" else 800
        canvas_width = max(canvas_width, min_width)
        
        # 更新canvas window的宽度
        canvas.itemconfig(canvas.find_withtag("all")[0], width=canvas_width)
        
        # 更新滚动区域
        canvas.configure(scrollregion=canvas.bbox("all"))
    
    def setup_a_tab(self, parent):
        # A表文件列表框架
        files_frame = ttk.LabelFrame(parent, text="日报表文件列表", padding="10 10 10 10")
        files_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 文件列表显示区域
        self.files_listbox_frame = ttk.Frame(files_frame)
        self.files_listbox_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 创建Treeview用于显示文件列表，并使用水平滚动
        tree_container = ttk.Frame(self.files_listbox_frame)
        tree_container.pack(fill=tk.BOTH, expand=True)
        
        # 创建Treeview的垂直滚动条
        tree_vsb = ttk.Scrollbar(tree_container, orient="vertical")
        tree_vsb.pack(side="right", fill="y")
        
        # 创建Treeview的水平滚动条
        tree_hsb = ttk.Scrollbar(tree_container, orient="horizontal")
        tree_hsb.pack(side="bottom", fill="x")
        
        # 创建Treeview
        columns = ("序号", "文件路径", "工作表名称")
        self.files_tree = ttk.Treeview(tree_container, columns=columns, show="headings", 
                                       height=7, yscrollcommand=tree_vsb.set, xscrollcommand=tree_hsb.set)
        
        # 设置滚动条的命令
        tree_vsb.config(command=self.files_tree.yview)
        tree_hsb.config(command=self.files_tree.xview)
        
        # 设置列宽和标题
        self.files_tree.column("序号", width=50, anchor="center", minwidth=50)
        self.files_tree.column("文件路径", width=450, minwidth=100)
        self.files_tree.column("工作表名称", width=150, anchor="center", minwidth=100)
        
        self.files_tree.heading("序号", text="序号")
        self.files_tree.heading("文件路径", text="文件路径")
        self.files_tree.heading("工作表名称", text="工作表名称")
        
        self.files_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 文件操作按钮区域 - 使用流式布局
        files_buttons_frame = ttk.Frame(files_frame)
        files_buttons_frame.pack(fill=tk.X, expand=False, pady=5)
        
        # 为按钮创建一个水平滚动的容器
        button_canvas = tk.Canvas(files_buttons_frame, height=40)
        button_canvas.pack(fill=tk.X, expand=True)
        
        # 在Canvas中创建一个Frame来放置按钮
        button_frame = ttk.Frame(button_canvas)
        button_frame_window = button_canvas.create_window((0,0), window=button_frame, anchor="nw")
        
        # 添加文件按钮
        add_file_button = tk.Button(
            button_frame, 
            text="添加文件", 
            command=self.add_a_file,
            font=("微软雅黑", 10),
            bg="#4CAF50",  # 绿色背景
            fg="white",    # 白色文字
            activebackground="#45a049",  # 鼠标悬停时的颜色
            activeforeground="white",
            relief=tk.RAISED,
            bd=1,
            cursor="hand2"  # 手型光标
        )
        add_file_button.pack(side=tk.LEFT, padx=5)
        
        # 删除选中文件按钮
        remove_file_button = tk.Button(
            button_frame, 
            text="删除选中", 
            command=self.remove_a_file,
            font=("微软雅黑", 10),
            bg="#4CAF50",
            fg="white",
            activebackground="#45a049",
            activeforeground="white",
            relief=tk.RAISED,
            bd=1,
            cursor="hand2"
        )
        remove_file_button.pack(side=tk.LEFT, padx=5)
        
        # 清空所有文件按钮
        clear_files_button = tk.Button(
            button_frame, 
            text="清空列表", 
            command=self.clear_a_files,
            font=("微软雅黑", 10),
            bg="#4CAF50",
            fg="white",
            activebackground="#45a049",
            activeforeground="white",
            relief=tk.RAISED,
            bd=1,
            cursor="hand2"
        )
        clear_files_button.pack(side=tk.LEFT, padx=5)
        
        # 设置工作表名按钮
        set_sheet_button = tk.Button(
            button_frame, 
            text="设置工作表名", 
            command=self.set_sheet_name,
            font=("微软雅黑", 10),
            bg="#4CAF50",
            fg="white",
            activebackground="#45a049",
            activeforeground="white",
            relief=tk.RAISED,
            bd=1,
            cursor="hand2"
        )
        set_sheet_button.pack(side=tk.LEFT, padx=5)
        
        # 更新button_frame的大小
        button_frame.update_idletasks()
        button_canvas.config(scrollregion=button_canvas.bbox("all"))
        
        # 如果按钮太多导致溢出，添加水平滚动条
        if button_frame.winfo_reqwidth() > button_canvas.winfo_width():
            button_hsb = ttk.Scrollbar(files_buttons_frame, orient="horizontal", command=button_canvas.xview)
            button_hsb.pack(fill=tk.X)
            button_canvas.config(xscrollcommand=button_hsb.set)
        
        # 添加Canvas大小调整事件
        button_canvas.bind("<Configure>", lambda e: button_canvas.itemconfig(
            button_frame_window, width=e.width))
        
        # 通用工作表名称区域 - 使用流式布局
        common_sheet_frame = ttk.Frame(parent)
        common_sheet_frame.pack(fill=tk.X, pady=10)
        
        # 创建一个容器
        sheet_wrapper = ttk.Frame(common_sheet_frame)
        sheet_wrapper.pack(fill=tk.X)
        
        ttk.Label(sheet_wrapper, text="通用工作表名称:", style="TLabel").pack(side=tk.LEFT)
        
        common_sheet_entry = ttk.Entry(sheet_wrapper, textvariable=self.a_common_sheet, width=20)
        common_sheet_entry.pack(side=tk.LEFT, padx=5)
        
        apply_common_button = tk.Button(
            sheet_wrapper, 
            text="应用到所有文件", 
            command=self.apply_common_sheet,
            font=("微软雅黑", 10),
            bg="#4CAF50",
            fg="white",
            activebackground="#45a049",
            activeforeground="white",
            relief=tk.RAISED,
            bd=1,
            cursor="hand2"
        )
        apply_common_button.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(sheet_wrapper, text="(留空使用各文件的默认活动表)", style="TLabel").pack(side=tk.LEFT)
        
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
        
        info_text = """日报表是源数据表，程序将查找此表中指定列与患者库匹配的行。
- 文件列表：可以添加多个Excel文件进行批量处理
- 工作表名称：可以为每个文件单独设置工作表名称，也可以使用通用工作表名应用到所有文件
- 比较列：用于与患者库比较的列，可以是字母(A,B,C)或数字(1,2,3)
- 对于含有多个工作表的Excel文件，可以单独指定要处理的工作表
- 系统会自动检测表头中包含"日期"的列，并将其格式化为中文日期格式"""
        
        ttk.Label(info_frame, text=info_text, style="TLabel", wraplength=650, justify=tk.LEFT).pack(pady=5)
    
    def setup_b_tab(self, parent):
        # B表文件路径 - 使用可伸缩布局，确保按钮始终可见
        path_frame = ttk.Frame(parent, style="TFrame")
        path_frame.pack(fill=tk.X, pady=(10, 5))
        
        ttk.Label(path_frame, text="患者库文件路径:", style="TLabel").pack(side=tk.LEFT)
        
        # 创建一个框架来包含输入框和按钮
        entry_button_frame = ttk.Frame(path_frame)
        entry_button_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        self.b_file_path = tk.StringVar()
        # 设置Entry的宽度比例
        path_entry = ttk.Entry(entry_button_frame, textvariable=self.b_file_path)
        path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        # 确保按钮始终可见
        browse_button = tk.Button(
            entry_button_frame, 
            text="浏览...", 
            command=self.browse_b_file,
            font=("微软雅黑", 10),
            bg="#4CAF50",
            fg="white",
            activebackground="#45a049",
            activeforeground="white",
            relief=tk.RAISED,
            bd=1,
            cursor="hand2"
        )
        browse_button.pack(side=tk.RIGHT)
        
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
        
        info_text = """患者库是对比数据表，程序将查找日报表中与此表指定列匹配的行。
- 文件路径：Excel文件的完整路径
- 工作表名称：要处理的工作表名称，留空将使用默认活动表
- 比较列：用于与日报表比较的列，可以是字母(A,B,C)或数字(1,2,3)"""
        
        ttk.Label(info_frame, text=info_text, style="TLabel", wraplength=650, justify=tk.LEFT).pack(pady=5)
    
    def setup_output_tab(self, parent):
        # 输出文件夹路径 - 使用可伸缩布局
        path_frame = ttk.Frame(parent, style="TFrame")
        path_frame.pack(fill=tk.X, pady=(10, 5))
        
        ttk.Label(path_frame, text="输出文件夹:", style="TLabel").pack(side=tk.LEFT)
        
        # 创建一个框架来包含输入框和按钮
        entry_button_frame = ttk.Frame(path_frame)
        entry_button_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        self.output_folder_path = tk.StringVar()
        # 设置Entry的宽度比例
        path_entry = ttk.Entry(entry_button_frame, textvariable=self.output_folder_path)
        path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        # 确保按钮始终可见
        browse_button = tk.Button(
            entry_button_frame, 
            text="浏览...", 
            command=self.browse_output_folder,
            font=("微软雅黑", 10),
            bg="#4CAF50",
            fg="white",
            activebackground="#45a049",
            activeforeground="white",
            relief=tk.RAISED,
            bd=1,
            cursor="hand2"
        )
        browse_button.pack(side=tk.RIGHT)
        
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
2. 程序会自动将日报表中的第一行作为表头复制到结果文件
3. 如果日报表中包含公式，只会保存计算结果，不保存公式本身
4. 程序会自动识别并保持原日报表中的合并单元格状态
5. 当处理多个日报表文件时，所有匹配的行将合并到一个结果文件中
6. 系统会自动检测表头中包含"日期"或"时间"的列，并将其格式化为中文日期格式(如"5月1日")"""
        
        ttk.Label(info_frame, text=info_text, style="TLabel", wraplength=650, justify=tk.LEFT).pack(pady=5)
    
    def add_a_file(self):
        """添加日报表文件到列表"""
        filenames = filedialog.askopenfilenames(
            title="选择一个或多个日报表文件",
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
        """删除选中的日报表文件"""
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
            title="选择患者库文件",
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
            messagebox.showerror("错误", "请添加至少一个日报表文件")
            return
        
        # 检查所有A表文件是否存在
        for file_path, _ in a_files:
            if not os.path.exists(file_path):
                messagebox.showerror("错误", f"日报表文件不存在: {file_path}")
                return
            
        if not b_file:
            messagebox.showerror("错误", "请选择患者库文件")
            return
        if not os.path.exists(b_file):
            messagebox.showerror("错误", f"患者库文件不存在: {b_file}")
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
            messagebox.showerror("错误", "请指定日报表的比较列")
            return
            
        if not b_col:
            messagebox.showerror("错误", "请指定患者库的比较列")
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
                    # 使用全局函数打开文件
                    open_file(saved_path)
            else:
                self.root.after(0, lambda: self.status_var.set("处理未完成"))
                self.root.after(0, lambda: self.result_text.insert(tk.END, "未找到匹配的数据或保存文件失败。"))
        except Exception as e:
            error_message = f"处理过程中出错:\n{str(e)}"
            self.root.after(0, lambda: self.status_var.set("处理失败"))
            self.root.after(0, lambda: self.result_text.insert(tk.END, error_message))
            self.root.after(0, lambda: messagebox.showerror("错误", error_message))

def open_file(file_path):
    """跨平台打开文件的函数"""
    try:
        if platform.system() == "Windows":
            os.startfile(file_path)
        elif platform.system() == "Darwin":  # macOS
            import subprocess
            subprocess.call(["open", file_path])
        else:  # Linux 或其他系统
            import subprocess
            subprocess.call(["xdg-open", file_path])
    except Exception as e:
        print(f"打开文件失败: {e}")
        return False
    return True

def main():
    # 处理高DPI显示的问题
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass  # 非Windows平台或者无法设置DPI感知
    
    root = tk.Tk()
    
    # 检测平台，为Mac调整字体和界面元素
    if platform.system() == "Darwin":  # macOS
        try:
            # 在Mac上应用默认字体
            default_font = tkfont.nametofont("TkDefaultFont")
            default_font.configure(size=12)
            root.option_add("*Font", default_font)
        except Exception as e:
            print(f"设置Mac字体时出错: {e}")
            # 备用方式：直接设置常用字体
            root.option_add("*Font", ("SF Pro", 12))
    
    # 设置应用图标
    try:
        if platform.system() == "Windows":
            icon_path = resource_path("excel_icon.ico")
            if os.path.exists(icon_path):
                root.iconbitmap(icon_path)
        elif platform.system() == "Darwin":  # macOS
            # Mac使用不同的图标设置方式，需要使用.icns文件
            # 这里我们先跳过，通常需要在Mac上专门创建.icns文件
            pass
    except Exception as e:
        print(f"设置图标时出错: {e}")
    
    app = ExcelProcessorUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()

    # todo: 添加日志，能否直接保存到程序里面，还是保存到本地文件
    