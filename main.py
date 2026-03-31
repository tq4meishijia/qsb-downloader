import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service as EdgeService  # 重命名避免冲突
from selenium.webdriver.edge.options import Options
from selenium.common.exceptions import StaleElementReferenceException, WebDriverException
import time
import base64
from io import BytesIO
from PIL import Image
import os
import threading
import queue
import re
import sys  # 新增sys模块用于检测打包环境

class ExamDownloaderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("箐师帮试卷下载工具")
        self.root.geometry("800x700")
        self.root.configure(bg="#f0f2f5")
        
        # 全局变量
        self.driver = None
        self.download_queue = queue.Queue()
        self.is_downloading = False
        
        # 创建主框架
        main_frame = ttk.Frame(root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        title = ttk.Label(main_frame, text="仅供学习使用，请尊重知识产权", font=("Arial", 16, "bold"))
        title.pack(pady=(0, 15))
        
        # 输入区域
        input_frame = ttk.LabelFrame(main_frame, text="配置参数", padding="15")
        input_frame.pack(fill=tk.X, pady=10)
        
        # 链接输入 - 改为多行文本框，方便粘贴多个链接
        ttk.Label(input_frame, text="试卷链接 (每行一个):").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.url_text = tk.Text(input_frame, height=6, width=60, font=("Arial", 10))
        self.url_text.grid(row=0, column=1, padx=5, pady=5, sticky=tk.EW)
        self.url_text.insert(tk.END, "")
        
        # 操作按钮
        btn_frame = ttk.Frame(input_frame)
        btn_frame.grid(row=1, column=0, columnspan=2, pady=10, sticky=tk.E)
        
        self.login_btn = ttk.Button(btn_frame, text="打开登录页面", command=self.start_login)
        self.login_btn.pack(side=tk.LEFT, padx=5)
        
        self.confirm_login_btn = ttk.Button(btn_frame, text="我已登录", state=tk.DISABLED, command=self.confirm_login)
        self.confirm_login_btn.pack(side=tk.LEFT, padx=5)
        
        self.download_btn = ttk.Button(btn_frame, text="开始批量下载", state=tk.DISABLED, command=self.start_batch_download)
        self.download_btn.pack(side=tk.LEFT, padx=5)
        
        # 状态区域
        status_frame = ttk.LabelFrame(main_frame, text="操作日志", padding="10")
        status_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.log_text = scrolledtext.ScrolledText(status_frame, wrap=tk.WORD, font=("Consolas", 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.config(state=tk.DISABLED)
        
        # 说明文字
        ttk.Label(main_frame, text="提示: 1. 点击'打开登录页面'后，请使用微信扫码登录\n"
                 "2. 登录成功后点击'我已登录'按钮\n"
                 "3. 在上方文本框中粘贴试卷链接，每行一个\n"
                 "4. 点击'开始批量下载'\n"
                 "5. 生成的PDF将自动命名并保存在当前目录", 
                 foreground="#555", font=("Arial", 9)).pack(pady=(0, 10), anchor=tk.W)
        
        # 状态栏
        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        status_bar = ttk.Label(root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def log(self, message, level="info"):
        """添加日志到UI"""
        self.log_text.config(state=tk.NORMAL)
        timestamp = time.strftime("%H:%M:%S")
        if level == "error":
            message = f"[ERROR] {timestamp} - {message}"
            self.log_text.tag_configure("error", foreground="red")
            self.log_text.insert(tk.END, message + "\n", "error")
        else:
            message = f"[INFO] {timestamp} - {message}"
            self.log_text.insert(tk.END, message + "\n")
        
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        
        # 更新状态栏
        self.status_var.set(message.split(" - ")[-1][:50])  # 截断长消息
        self.root.update_idletasks()

    def get_resource_path(self, relative_path):
        """获取资源的绝对路径，适用于打包环境和开发环境"""
        try:
            # PyInstaller打包后的临时目录
            base_path = sys._MEIPASS
        except AttributeError:
            # 正常开发环境
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)

    def get_edgedriver_path(self):
        """获取EdgeDriver的正确路径，兼容打包环境"""
        # 尝试在资源路径中查找
        driver_path = self.get_resource_path("msedgedriver.exe" if os.name == 'nt' else "msedgedriver")
        
        # 检查驱动是否存在
        if not os.path.exists(driver_path):
            # 尝试在当前工作目录查找
            driver_path = os.path.join(os.getcwd(), "msedgedriver.exe" if os.name == 'nt' else "msedgedriver")
            
            if not os.path.exists(driver_path):
                error_msg = (
                    "EdgeDriver 未找到！\n\n"
                    "请执行以下操作之一：\n"
                    "1. 将 msedgedriver.exe (Windows) 或 msedgedriver (Mac/Linux) "
                    "文件放在程序同目录下\n"
                    "2. 从官方下载匹配您Edge浏览器版本的驱动：\n"
                    "   https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/"
                )
                raise FileNotFoundError(error_msg)
        
        self.log(f"EdgeDriver 路径: {driver_path}", "info")
        return driver_path

    def get_edge_binary_path(self):
        """获取Edge浏览器二进制路径"""
        if os.name != 'nt':
            # 非Windows系统通常不需要指定路径
            return None
            
        # Windows系统可能的Edge安装路径
        possible_paths = [
            "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe",
            "C:/Program Files/Microsoft/Edge/Application/msedge.exe",
            os.path.expandvars("%LOCALAPPDATA%/Microsoft/Edge/Application/msedge.exe"),
            os.path.expandvars("%PROGRAMFILES%/Microsoft/Edge/Application/msedge.exe"),
            os.path.expandvars("%PROGRAMFILES(X86)%/Microsoft/Edge/Application/msedge.exe"),
        ]
        
        for path in possible_paths:
            if os.path.exists(path):
                self.log(f"找到Edge浏览器: {path}", "info")
                return path
                
        self.log("未找到Edge浏览器安装路径，将使用系统默认路径", "info")
        return None  # 让Selenium自动查找

    def start_login(self):
        """启动登录流程"""
        self.log("正在打开登录页面...")
        self.login_btn.config(state=tk.DISABLED)
        threading.Thread(target=self.login_thread, daemon=True).start()

    def login_thread(self):
        """登录线程：只负责打开登录页"""
        try:
            driver_path = self.get_edgedriver_path()
            edge_options = Options()
            edge_options.add_argument("--start-maximized")
            edge_options.add_argument("--disable-notifications")
            edge_options.add_experimental_option("excludeSwitches", ["enable-automation"])
            edge_options.add_experimental_option('useAutomationExtension', False)
            
            # 设置Edge浏览器路径
            edge_binary = self.get_edge_binary_path()
            if edge_binary:
                edge_options.binary_location = edge_binary
            
            # 创建服务对象 - 使用重命名的EdgeService
            service = EdgeService(executable_path=driver_path)
            
            # 初始化WebDriver - 使用标准导入方式
            self.driver = webdriver.Edge(service=service, options=edge_options)
            
            # 固定打开首页，等待用户手动登录
            self.driver.get("https://www.jingshibang.com/home/")
            self.log("✅ 已打开登录页面，请在浏览器中使用微信扫码登录...")
            self.root.after(0, lambda: self.confirm_login_btn.config(state=tk.NORMAL))
            
        except Exception as e:
            error_msg = f"打开登录页面失败: {str(e)}"
            self.log(error_msg, "error")
            self.root.after(0, lambda: (
                messagebox.showerror("错误", error_msg),
                self.login_btn.config(state=tk.NORMAL)
            ))
            if self.driver:
                try:
                    self.driver.quit()
                except:
                    pass
            self.driver = None

    def confirm_login(self):
        """手动确认登录状态"""
        if self.driver:
            # 验证是否真的登录成功
            try:
                WebDriverWait(self.driver, 5).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, ".user-avatar, .header-user-info"))
                )
                self.log("✅ 登录状态已确认！")
                self.confirm_login_btn.config(state=tk.DISABLED)
                self.download_btn.config(state=tk.NORMAL)
            except:
                messagebox.showwarning("警告", "检测到未登录状态，请先完成登录！")
                self.log("⚠️ 未检测到登录状态，请先完成微信扫码登录", "error")
        else:
            messagebox.showwarning("警告", "请先打开登录页面！")

    def start_batch_download(self):
        """启动批量下载流程"""
        if not self.driver:
            messagebox.showerror("错误", "请先完成微信登录")
            return

        urls_text = self.url_text.get("1.0", tk.END).strip()
        urls = [url.strip() for url in urls_text.split('\n') if url.strip()]
        
        if not urls:
            messagebox.showwarning("警告", "请至少输入一个试卷链接！")
            return

        self.is_downloading = True
        self.download_btn.config(state=tk.DISABLED)
        self.confirm_login_btn.config(state=tk.DISABLED)
        self.login_btn.config(state=tk.DISABLED)
        
        self.log(f"开始批量下载 {len(urls)} 份试卷...")
        self.download_queue.queue.clear()
        for url in urls:
            self.download_queue.put(url)
        
        threading.Thread(target=self.batch_download_worker, daemon=True).start()

    def batch_download_worker(self):
        """批量下载工作线程"""
        total_tasks = self.download_queue.qsize()
        current_task = 0
        
        while not self.download_queue.empty() and self.is_downloading:
            current_task += 1
            url = self.download_queue.get()
            self.log(f"--- 正在下载第 {current_task}/{total_tasks} 份试卷 ---")
            self.download_single_paper(url)
            self.download_queue.task_done()
        
        self.log("\n🎉 批量下载全部完成！")
        self.is_downloading = False
        self.root.after(0, lambda: (
            self.download_btn.config(state=tk.NORMAL),
            self.login_btn.config(state=tk.NORMAL)
        ))

    def download_single_paper(self, url):
        """下载单份试卷的核心逻辑"""
        try:
            self.log(f"正在加载试卷: {url}")
            self.driver.get(url)
            time.sleep(3)

            # 尝试获取试卷标题 - 专门针对 div.detail-header-title
            self.log("正在尝试获取试卷标题...")
            pdf_filename = "未命名试卷.pdf"
            
            try:
                # 直接查找具有该class的div元素
                title_element = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "div.detail-header-title"))
                )
                
                if title_element.text.strip():
                    raw_title = title_element.text.strip()
                    # 清理文件名中的非法字符
                    safe_title = re.sub(r'[<>:"/\\|?*]', '_', raw_title)
                    # 确保文件名不过长
                    safe_title = safe_title[:100].strip()
                    if safe_title:
                        pdf_filename = f"{safe_title}.pdf"
                    self.log(f"✅ 成功获取试卷标题: {safe_title}")
                else:
                    self.log("⚠️ 未能获取试卷标题，使用默认名称。")

            except Exception as e:
                self.log(f"查找 .detail-header-title 失败: {str(e)}，尝试备选方案")
                
                # 备选方案：查找其他可能的标题元素
                title_selectors = [
                    ".paper-title", ".exam-title", ".detail-title", 
                    ".paper-name", ".title", "h1.title", "h2.title"
                ]
                
                for selector in title_selectors:
                    try:
                        title_element = self.driver.find_element(By.CSS_SELECTOR, selector)
                        if title_element and title_element.text.strip():
                            raw_title = title_element.text.strip()
                            safe_title = re.sub(r'[<>:"/\\|?*]', '_', raw_title)
                            safe_title = safe_title[:100].strip()
                            if safe_title:
                                pdf_filename = f"{safe_title}.pdf"
                                self.log(f"✅ 通过备选方案获取标题: {safe_title}")
                                break
                    except:
                        continue

            # 滚动页面加载所有内容
            self.log("开始滚动页面以加载所有内容...")
            last_height = self.driver.execute_script("return document.body.scrollHeight")
            
            # 滚动3次确保加载完全
            for _ in range(3):
                self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(1.5)
                new_height = self.driver.execute_script("return document.body.scrollHeight")
                if new_height == last_height:
                    break
                last_height = new_height
            self.log("页面滚动完成。")

            # 查找所有Canvas和Img元素
            self.log("正在查找所有图片和Canvas元素...")
            
            # 滚动后重新查找元素
            canvas_elements = self.driver.find_elements(By.TAG_NAME, "canvas")
            img_elements = self.driver.find_elements(By.CSS_SELECTOR, "img[src*='paper'], img[src*='content']")
            
            self.log(f"找到 {len(canvas_elements)} 个 Canvas 元素和 {len(img_elements)} 个相关图片元素")
            
            all_elements = canvas_elements + img_elements
            if not all_elements:
                self.log("❌ 页面上未找到任何Canvas或相关Img元素", "error")
                return

            images = []
            total_elements = len(all_elements)
            for i, element in enumerate(all_elements):
                try:
                    if not element.is_displayed():
                        continue
                    
                    tag_name = element.tag_name
                    img = None
                    
                    if tag_name == 'canvas':
                        # 处理Canvas
                        try:
                            img_data_b64 = self.driver.execute_script("""
                                var canvas = arguments[0];
                                return canvas.toDataURL('image/png');
                            """, element)
                            
                            if img_data_b64 and img_data_b64.startswith('data:image'):
                                img_data = img_data_b64.split(',')[1]
                                img = Image.open(BytesIO(base64.b64decode(img_data)))
                                self.log(f"  ✅ 已处理 Canvas #{i+1}/{total_elements}")
                        except Exception as e:
                            self.log(f"  ⚠️  处理 Canvas #{i+1} 时出错: {str(e)}")
                    
                    elif tag_name == 'img':
                        # 处理Img
                        try:
                            img_src = element.get_attribute('src')
                            if not img_src or 'base64' not in img_src and not img_src.startswith('http'):
                                continue
                            
                            if 'base64' in img_src:
                                # Base64图片
                                img_data = img_src.split(',')[1]
                                img = Image.open(BytesIO(base64.b64decode(img_data)))
                            else:
                                # 网络图片，使用selenium截图
                                element.location_once_scrolled_into_view
                                time.sleep(0.5)
                                img_binary = element.screenshot_as_png
                                img = Image.open(BytesIO(img_binary))
                            
                            self.log(f"  ✅ 已处理 Img #{i+1}/{total_elements}")
                        except Exception as e:
                            self.log(f"  ⚠️  处理 Img #{i+1} 时出错: {str(e)}")
                    
                    # 检查图片是否有效
                    if img and img.width > 100 and img.height > 100:
                        images.append(img)
                    elif img:
                        self.log(f"  ⚠️  跳过小图片 ({img.width}x{img.height})")
                    
                except StaleElementReferenceException:
                    self.log(f"  ⚠️  元素已过时，跳过第{i+1}个元素")
                    continue
                except Exception as e:
                    self.log(f"  ❌ 处理第{i+1}个元素时出错: {str(e)}")
                    continue

            if not images:
                self.log("❌ 没有找到有效的试卷图片内容", "error")
                return

            # 生成PDF
            self.log(f"正在生成PDF文件: {pdf_filename} ({len(images)} 页)")
            
            # 确保所有图片都是RGB模式
            rgb_images = []
            for i, img in enumerate(images):
                try:
                    if img.mode != 'RGB':
                        img = img.convert('RGB')
                    rgb_images.append(img)
                except Exception as e:
                    self.log(f"  ⚠️  转换图片 #{i+1} 时出错: {str(e)}")
            
            if not rgb_images:
                self.log("❌ 没有可转换为RGB的有效图片", "error")
                return
            
            # 保存PDF
            first_img = rgb_images[0]
            if len(rgb_images) > 1:
                first_img.save(pdf_filename, save_all=True, append_images=rgb_images[1:], resolution=100.0)
            else:
                first_img.save(pdf_filename, resolution=100.0)
            
            self.log(f"✅ 试卷已成功保存为: {pdf_filename}")
            
        except WebDriverException as e:
            if "invalid session id" in str(e).lower():
                self.log("❌ 浏览器会话已失效，请重新登录", "error")
                self.root.after(0, lambda: (
                    messagebox.showerror("会话失效", "浏览器会话已失效，请重新登录"),
                    self.confirm_login_btn.config(state=tk.NORMAL),
                    self.download_btn.config(state=tk.DISABLED)
                ))
            else:
                self.log(f"WebDriver错误: {str(e)}", "error")
        except Exception as e:
            self.log(f"下载试卷时出错: {str(e)}", "error")
            import traceback
            traceback.print_exc()

    def on_closing(self):
        """关闭窗口处理"""
        if self.is_downloading:
            if messagebox.askyesno("确认退出", "下载任务正在进行中，确定要退出吗？"):
                self.is_downloading = False
                if self.driver:
                    try:
                        self.driver.quit()
                    except:
                        pass
                self.root.destroy()
        else:
            if self.driver:
                try:
                    self.driver.quit()
                except:
                    pass
            self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = ExamDownloaderApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()