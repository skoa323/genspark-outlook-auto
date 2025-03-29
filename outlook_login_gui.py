import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import threading
import os
import sys
import tempfile
import logging
from datetime import datetime
import time
import pandas as pd
import csv
import queue
import traceback

# 设置日志
def setup_logger():
    log_dir = "logs"
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = os.path.join(log_dir, f"outlook_automation_{timestamp}.log")
    
    logger = logging.getLogger("OutlookAutomation")
    logger.setLevel(logging.INFO)
    
    # 文件处理器
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setLevel(logging.INFO)
    
    # 控制台处理器
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    
    # 格式化器
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)
    
    # 添加处理器
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    
    return logger, log_file

logger, log_file_path = setup_logger()

class RedirectText:
    def __init__(self, text_widget):
        self.text_widget = text_widget
        self.buffer = ""

    def write(self, string):
        self.buffer += string
        self.text_widget.configure(state="normal")
        self.text_widget.insert(tk.END, string)
        self.text_widget.see(tk.END)
        self.text_widget.configure(state="disabled")

    def flush(self):
        pass

class OutlookAutomationGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Outlook登录自动化工具")
        
        # 增大默认窗口尺寸，确保所有控件可见
        self.root.geometry("800x700")
        
        # 设置更合理的最小窗口尺寸
        self.root.minsize(750, 650)
        
        self.running = False
        self.drivers = {}
        self.automation_thread = None
        self.stdout_backup = None
        
        # 设置GUI组件
        self.setup_ui()
        
        # 默认GENSPARK_URL
        self.genspark_url_var.set("https://www.genspark.ai/invite?invite_code=YTI2MTEzMzVMZmVkZkxkMWI0TGM4MGFMZmU2NDg3NDgyMTk4")
        
        # 默认并发数
        self.concurrent_var.set(5)
        
    def setup_ui(self):
        # 创建主框架 - 添加padding以确保内容不会太靠近窗口边缘
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Genspark URL输入
        url_frame = ttk.Frame(main_frame)
        url_frame.pack(fill=tk.X, pady=(0, 5))  # 减少顶部边距
        
        ttk.Label(url_frame, text="Genspark邀请链接:").pack(side=tk.LEFT, padx=5)
        self.genspark_url_var = tk.StringVar()
        ttk.Entry(url_frame, textvariable=self.genspark_url_var, width=50).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # 最大并发数设置
        concurrent_frame = ttk.Frame(main_frame)
        concurrent_frame.pack(fill=tk.X, pady=(0, 5))  # 减少顶部边距
        
        ttk.Label(concurrent_frame, text="最大并发数量:").pack(side=tk.LEFT, padx=5)
        self.concurrent_var = tk.IntVar()
        concurrent_spinbox = ttk.Spinbox(
            concurrent_frame, 
            from_=1, 
            to=10, 
            textvariable=self.concurrent_var, 
            width=5
        )
        concurrent_spinbox.pack(side=tk.LEFT, padx=5)
        
        # Outlook账号输入 - 减小高度以确保按钮可见
        account_frame = ttk.LabelFrame(main_frame, text="Outlook账号 (格式: email,password，每行一个)")
        account_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 5))  # 减少顶部边距
        
        # 添加账号输入的文本框，减少高度
        self.accounts_text = scrolledtext.ScrolledText(account_frame, height=8)
        self.accounts_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 添加导入CSV按钮
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill=tk.X, pady=(0, 5))  # 减少顶部边距
        
        ttk.Button(
            buttons_frame, 
            text="从CSV导入", 
            command=self.import_from_csv
        ).pack(side=tk.LEFT, padx=5)
        
        # 运行状态
        status_frame = ttk.Frame(main_frame)
        status_frame.pack(fill=tk.X, pady=(0, 5))  # 减少顶部边距
        
        ttk.Label(status_frame, text="状态:").pack(side=tk.LEFT, padx=5)
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(status_frame, textvariable=self.status_var).pack(side=tk.LEFT, padx=5)
        
        # 日志输出区域 - 减小高度以确保按钮可见
        log_frame = ttk.LabelFrame(main_frame, text="运行日志")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 5))  # 减少顶部边距
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=10)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 操作按钮区域 - 确保这是最后一个元素，并且不扩展，固定在底部
        control_frame = ttk.Frame(main_frame)
        control_frame.pack(fill=tk.X, pady=(0, 5), side=tk.BOTTOM)
        
        # 确保按钮使用网格布局，更好地控制位置
        ttk.Button(
            control_frame, 
            text="运行", 
            command=self.start_automation,
            width=10  # 固定宽度
        ).grid(row=0, column=0, padx=5, pady=5)
        
        self.run_button = ttk.Button(
            control_frame, 
            text="运行", 
            command=self.start_automation,
            width=10  # 固定宽度
        )
        self.run_button.grid(row=0, column=0, padx=5, pady=5)
        
        self.stop_button = ttk.Button(
            control_frame, 
            text="停止", 
            command=self.stop_automation,
            state=tk.DISABLED,
            width=10  # 固定宽度
        )
        self.stop_button.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Button(
            control_frame, 
            text="清除日志", 
            command=self.clear_log,
            width=10  # 固定宽度
        ).grid(row=0, column=2, padx=5, pady=5)
        
        ttk.Button(
            control_frame, 
            text="打开日志文件夹", 
            command=self.open_log_folder,
            width=15  # 固定宽度
        ).grid(row=0, column=3, padx=5, pady=5)
        
        # 平衡网格列的权重，使按钮分布均匀
        for i in range(4):
            control_frame.columnconfigure(i, weight=1)
    
    def import_from_csv(self):
        """从CSV文件导入账号"""
        file_path = filedialog.askopenfilename(
            title="选择Outlook账号CSV文件",
            filetypes=[("CSV文件", "*.csv"), ("所有文件", "*.*")]
        )
        
        if not file_path:
            return
            
        try:
            # 读取CSV文件
            df = pd.read_csv(file_path)
            
            # 检查格式是否正确
            if "email" not in df.columns or "password" not in df.columns:
                messagebox.showerror("错误", "CSV文件格式不正确，必须包含email和password列")
                return
                
            # 清空当前文本
            self.accounts_text.delete('1.0', tk.END)
            
            # 添加账号信息到文本框
            for _, row in df.iterrows():
                self.accounts_text.insert(tk.END, f"{row['email']},{row['password']}\n")
                
            messagebox.showinfo("成功", f"已导入 {len(df)} 个账号")
            
        except Exception as e:
            messagebox.showerror("错误", f"导入CSV文件时出错: {str(e)}")
    
    def clear_log(self):
        """清除日志文本框"""
        self.log_text.configure(state="normal")
        self.log_text.delete('1.0', tk.END)
        self.log_text.configure(state="disabled")
    
    def open_log_folder(self):
        """打开日志文件夹"""
        log_dir = os.path.abspath("logs")
        if os.path.exists(log_dir):
            os.startfile(log_dir)
        else:
            messagebox.showinfo("提示", "日志文件夹不存在，将在运行时创建")
    
    def create_accounts_file(self):
        """从GUI文本框创建账号CSV文件"""
        accounts_content = self.accounts_text.get('1.0', tk.END).strip()
        if not accounts_content:
            messagebox.showerror("错误", "请输入Outlook账号信息")
            return None
        
        # 创建临时文件
        temp_dir = tempfile.gettempdir()
        accounts_file = os.path.join(temp_dir, "outlook_accounts_temp.csv")
        
        try:
            with open(accounts_file, 'w', newline='', encoding='utf-8') as f:
                f.write("email,password\n")  # 写入标题行
                for line in accounts_content.split('\n'):
                    line = line.strip()
                    if line:  # 跳过空行
                        f.write(f"{line}\n")
            return accounts_file
        except Exception as e:
            messagebox.showerror("错误", f"创建账号文件时出错: {str(e)}")
            return None
    
    def start_automation(self):
        """启动自动化流程"""
        if self.running:
            messagebox.showinfo("提示", "自动化任务已在运行中")
            return
        
        # 获取用户输入
        genspark_url = self.genspark_url_var.get().strip()
        if not genspark_url:
            messagebox.showerror("错误", "请输入Genspark邀请链接")
            return
        
        # 获取最大并发数
        try:
            max_concurrent = self.concurrent_var.get()
            if max_concurrent <= 0:
                raise ValueError("并发数必须大于0")
        except:
            messagebox.showerror("错误", "请输入有效的最大并发数")
            return
        
        # 创建临时账号文件
        accounts_file = self.create_accounts_file()
        if not accounts_file:
            return
        
        # 更新UI状态
        self.running = True
        self.run_button.configure(state=tk.DISABLED)
        self.stop_button.configure(state=tk.NORMAL)
        self.status_var.set("正在运行...")
        
        # 清空日志
        self.clear_log()
        
        # 重定向stdout到日志文本框
        self.stdout_backup = sys.stdout
        self.redirect = RedirectText(self.log_text)
        sys.stdout = self.redirect
        
        # 在单独的线程中运行自动化
        self.automation_thread = threading.Thread(
            target=self.run_automation_thread, 
            args=(genspark_url, accounts_file, max_concurrent)
        )
        self.automation_thread.daemon = True
        self.automation_thread.start()
    
    def run_automation_thread(self, genspark_url, accounts_file, max_concurrent):
        """运行自动化线程"""
        try:
            # 导入必要的模块
            from outlook_login_automation import main as automation_main
            # 设置全局变量
            import outlook_login_automation
            outlook_login_automation.GENSPARK_URL = genspark_url
            
            # 修改默认输入文件路径
            import os
            original_cwd = os.getcwd()
            outlook_login_automation_dir = os.path.dirname(os.path.abspath(outlook_login_automation.__file__))
            os.chdir(outlook_login_automation_dir)
            
            # 复制临时账号文件到工作目录
            import shutil
            shutil.copy2(accounts_file, "outlook_accounts.csv")
            
            # 运行自动化
            logger.info(f"启动自动化，Genspark URL: {genspark_url}")
            logger.info(f"最大并发数: {max_concurrent}")
            automation_main(max_concurrent=max_concurrent)
            
            # 恢复工作目录
            os.chdir(original_cwd)
            
        except ImportError:
            logger.error("导入outlook_login_automation模块失败")
            logger.error("请确保outlook_login_automation.py文件在同一目录下")
            messagebox.showerror("错误", "找不到自动化模块，请确保outlook_login_automation.py文件在同一目录下")
        except Exception as e:
            logger.error(f"自动化过程出错: {str(e)}")
            logger.error(traceback.format_exc())
            
        finally:
            # 恢复stdout
            if self.stdout_backup:
                sys.stdout = self.stdout_backup
            
            # 删除临时文件
            try:
                if os.path.exists(accounts_file):
                    os.remove(accounts_file)
            except:
                pass
            
            # 更新UI状态
            self.root.after(0, self.update_ui_after_run)
    
    def update_ui_after_run(self):
        """运行完成后更新UI状态"""
        self.running = False
        self.run_button.configure(state=tk.NORMAL)
        self.stop_button.configure(state=tk.DISABLED)
        self.status_var.set("运行完成")
        messagebox.showinfo("完成", "自动化流程已完成，请查看日志获取详细信息")
    
    def stop_automation(self):
        """停止自动化流程"""
        if not self.running:
            return
        
        if messagebox.askyesno("确认", "确定要停止当前运行的自动化流程吗？\n注意：这只会停止启动新任务，已经在运行的浏览器窗口将继续运行。"):
            self.status_var.set("正在停止...")
            messagebox.showinfo("提示", "程序已标记为停止，但当前操作可能需要一段时间才能完成")
            # 在实际实现中，需要有一种方式通知主线程停止创建新的任务
            
    def on_closing(self):
        """窗口关闭事件处理"""
        if self.running:
            if not messagebox.askyesno("确认", "自动化程序正在运行中，确定要退出吗?"):
                return
        self.root.destroy()

if __name__ == "__main__":
    # 设置更好的GUI缩放
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass

    root = tk.Tk()
    app = OutlookAutomationGUI(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop() 