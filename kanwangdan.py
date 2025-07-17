import tkinter as tk
from tkinter import ttk, messagebox
import subprocess
import sys
import os
import threading
from datetime import datetime

class VillageScheduleApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("看望单自动生成系统")
        self.root.geometry("700x700")
        self.root.resizable(False, False)
        
        # 设置窗口图标（如果有的话）
        # self.root.iconbitmap("icon.ico")
        
        # 设置样式
        self.setup_styles()
        
        # 创建界面
        self.create_widgets()
        
        # 居中显示窗口
        self.center_window()
    
    def setup_styles(self):
        """设置样式"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # 配置按钮样式
        style.configure('Large.TButton', 
                       font=('Microsoft YaHei', 12, 'bold'),
                       padding=(20, 10))
        
        # 配置标签样式
        style.configure('Title.TLabel', 
                       font=('Microsoft YaHei', 16, 'bold'),
                       foreground='#2c3e50')
        
        style.configure('Info.TLabel', 
                       font=('Microsoft YaHei', 10),
                       foreground='#34495e')
        
        style.configure('Author.TLabel', 
                       font=('Microsoft YaHei', 9),
                       foreground='#7f8c8d')
    
    def create_widgets(self):
        """创建界面组件"""
        # 主容器
        main_frame = ttk.Frame(self.root, padding="30")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        title_label = ttk.Label(main_frame, 
                               text="看望单自动生成系统", 
                               style='Title.TLabel')
        title_label.pack(pady=(0, 20))
        
        # 系统信息
        info_frame = ttk.Frame(main_frame)
        info_frame.pack(pady=(0, 30))
        
        info_text = """本系统由林三化编写，用于自动生成看望人员调度安排
请选择要执行的操作："""
        
        info_label = ttk.Label(info_frame, 
                              text=info_text, 
                              style='Info.TLabel',
                              justify=tk.CENTER)
        info_label.pack()
        
        # 按钮容器
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=20)
        
        # K1按钮 - 生成调度表
        self.k1_button = ttk.Button(button_frame,
                                   text="生成总调度表\n(K1)",
                                   style='Large.TButton',
                                   command=self.run_k1,
                                   width=20)
        self.k1_button.pack(pady=10)
        
        # K2按钮 - 生成个人日程
        self.k2_button = ttk.Button(button_frame,
                                   text="生成同工个人日程\n(K2)",
                                   style='Large.TButton',
                                   command=self.run_k2,
                                   width=20)
        self.k2_button.pack(pady=10)
        
        # 使用说明按钮
        self.help_button = ttk.Button(button_frame,
                                     text="使用说明",
                                     style='Large.TButton',
                                     command=self.show_help,
                                     width=20)
        self.help_button.pack(pady=10)
        
        # 状态显示
        self.status_frame = ttk.Frame(main_frame)
        self.status_frame.pack(pady=(30, 0), fill=tk.X)
        
        self.status_label = ttk.Label(self.status_frame,
                                     text="就绪",
                                     style='Info.TLabel')
        self.status_label.pack()
        
        # 进度条
        self.progress = ttk.Progressbar(self.status_frame,
                                       mode='indeterminate',
                                       length=300)
        self.progress.pack(pady=(10, 0))
        
        # 作者信息
        author_frame = ttk.Frame(main_frame)
        author_frame.pack(side=tk.BOTTOM, fill=tk.X)
        
        separator = ttk.Separator(author_frame, orient=tk.HORIZONTAL)
        separator.pack(fill=tk.X, pady=(20, 10))
        
        author_label = ttk.Label(author_frame,
                                text=f"版本 1.0 | 开发时间：{datetime.now().strftime('%Y年%m月')}",
                                style='Author.TLabel')
        author_label.pack()
        
        # 添加使用说明按钮到底部（移除原来的小按钮）
        # help_button = ttk.Button(author_frame,
        #                        text="使用说明",
        #                        command=self.show_help,
        #                        width=10)
        # help_button.pack(pady=(5, 0))
    
    def center_window(self):
        """居中显示窗口"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def show_help(self):
        """显示使用说明"""
        help_window = tk.Toplevel(self.root)
        help_window.title("使用说明")
        help_window.geometry("600x500")
        help_window.resizable(True, True)
        
        # 居中显示
        help_window.transient(self.root)
        help_window.grab_set()
        
        # 创建文本框和滚动条
        text_frame = ttk.Frame(help_window)
        text_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # 文本框
        text_widget = tk.Text(text_frame, wrap=tk.WORD, font=('Microsoft YaHei', 10))
        scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        
        # 使用说明内容
        help_text = """看望单自动生成系统 - 使用说明

═══════════════════════════════════════════════════════════════

📋 系统功能概述：
本系统包含两个主要功能模块，用于管理教会看望人员调度安排。

═══════════════════════════════════════════════════════════════

🔧 功能1：生成调度表 (K1)

📁 需要准备的文件：
  • 村.csv 或 村.xlsx - 包含村庄信息和接待人员
  • 看望人员.csv 或 看望人员.xlsx - 包含访问人员信息

📝 文件格式要求：
  • 村庄文件：第1列为村名，其他列为接待人员姓名
  • 人员文件：包含"人员1"、"人员2"、"计划出访次数"列

⚙️ 功能说明：
  • 自动获取下一个季度的所有星期六日期
  • 根据人员配置和村庄信息生成访问调度
  • 智能避免同一团队短时间内重复访问同一村庄
  • 优化调度安排，确保访问时间分布均匀

📤 输出结果：
  • 总表.xlsx - 完整的调度安排表格

═══════════════════════════════════════════════════════════════

👤 功能2：生成个人日程 (K2)

📁 需要准备的文件：
  • 总表.xlsx - 由功能1生成的调度表

⚙️ 功能说明：
  • 从总表中提取每个人的个人日程安排
  • 生成个人专属的Excel表格
  • 创建日历提醒文件(.ics格式)

📤 输出结果：
  • 个人日程表/ 文件夹，包含：
    - [姓名].xlsx - 个人日程Excel表格
    - [姓名]_日程提醒.ics - 可导入日历的提醒文件

═══════════════════════════════════════════════════════════════

📋 使用步骤：

第一步：准备数据文件，建议在模板上修改增删
  1. 将村庄信息保存为"村.csv"或"村.xlsx"
  2. 将人员信息保存为"看望人员.csv"或"看望人员.xlsx"
  3. 确保文件与程序在同一目录下

第二步：生成调度表
  1. 点击"生成调度表(K1)"按钮
  2. 耐心等待程序执行完成
  3. 检查生成的"总表.xlsx"文件

第三步：生成个人日程
  1. 点击"生成个人日程(K2)"按钮
  2. 耐心等待程序执行完成
  3. 查看"个人日程表"文件夹中的结果

═══════════════════════════════════════════════════════════════

⚠️ 注意事项：

• 文件位置：所有输入文件必须与程序在同一目录下
• 文件格式：支持.csv和.xlsx格式，使用Excel格式时，需删csv文件
• 编码问题：CSV文件请使用UTF-8或GBK编码保存
• 执行顺序：必须先执行K1生成总表，再执行K2生成个人日程
• 程序运行：执行过程中请勿关闭程序窗口
对总表安排不满意，可以多次点击k1，
也可以自己做个总表，只执行k2生成个人日程
个人日程需导入手机日程才会按时提示
═══════════════════════════════════════════════════════════════

🔍 常见问题解决：

问题1：提示"未找到文件"
  → 检查输入文件是否在正确位置
  → 确认文件名是否正确

问题2：程序执行出错
  → 检查数据文件格式是否正确
  → 确认文件内容是否完整

问题3：生成的日程不正确
  → 检查输入数据是否准确
  → 重新执行K1生成新的总表

═══════════════════════════════════════════════════════════════

📧 技术支持：
如遇到问题，请检查：
1. 文件格式和内容是否正确
2. 程序执行过程中的错误提示
3. 输出文件是否正常生成

版本：1.0  |  更新日期：2025年
"""
        
        # 插入文本
        text_widget.insert(tk.END, help_text)
        text_widget.config(state=tk.DISABLED)  # 设置为只读
        
        # 布局
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 关闭按钮
        close_button = ttk.Button(help_window, text="关闭", command=help_window.destroy)
        close_button.pack(pady=10)
        
        # 居中显示窗口
        help_window.update_idletasks()
        width = help_window.winfo_width()
        height = help_window.winfo_height()
        x = (help_window.winfo_screenwidth() // 2) - (width // 2)
        y = (help_window.winfo_screenheight() // 2) - (height // 2)
        help_window.geometry(f'{width}x{height}+{x}+{y}')
    
    def update_status(self, message):
        """更新状态显示"""
        self.status_label.config(text=message)
        self.root.update_idletasks()
    
    def run_script(self, script_name, button_text):
        """运行Python脚本"""
        try:
            # 禁用所有按钮
            self.k1_button.config(state='disabled')
            self.k2_button.config(state='disabled')
            self.help_button.config(state='disabled')
            
            # 开始进度条
            self.progress.start(10)
            self.update_status(f"正在执行{button_text}...")
            
            # 检查脚本文件是否存在
            if not os.path.exists(script_name):
                raise FileNotFoundError(f"未找到脚本文件：{script_name}")
            
            # 运行脚本 - 修复编码问题
            result = subprocess.run([sys.executable, script_name], 
                                  capture_output=True, 
                                  text=True, 
                                  encoding='gbk',  # Windows系统使用GBK编码
                                  errors='ignore')  # 忽略编码错误
            
            # 停止进度条
            self.progress.stop()
            
            if result.returncode == 0:
                self.update_status(f"{button_text}执行成功！")
                messagebox.showinfo("成功", f"{button_text}执行成功！\n\n输出信息：\n{result.stdout}")
            else:
                self.update_status(f"{button_text}执行失败！")
                messagebox.showerror("错误", f"{button_text}执行失败！\n\n错误信息：\n{result.stderr}")
                
        except FileNotFoundError as e:
            self.progress.stop()
            self.update_status("文件未找到")
            messagebox.showerror("错误", str(e))
        except Exception as e:
            self.progress.stop()
            self.update_status("执行出错")
            messagebox.showerror("错误", f"执行出错：{str(e)}")
        finally:
            # 恢复所有按钮状态
            self.k1_button.config(state='normal')
            self.k2_button.config(state='normal')
            self.help_button.config(state='normal')
            self.progress.stop()
    
    def run_k1(self):
        """运行K1脚本"""
        def run_thread():
            self.run_script("k1.py", "生成调度表")
        
        thread = threading.Thread(target=run_thread)
        thread.daemon = True
        thread.start()
    
    def run_k2(self):
        """运行K2脚本"""
        def run_thread():
            self.run_script("k2.py", "生成个人日程")
        
        thread = threading.Thread(target=run_thread)
        thread.daemon = True
        thread.start()
    
    def run(self):
        """运行应用程序"""
        self.root.mainloop()

if __name__ == "__main__":
    app = VillageScheduleApp()
    app.run()