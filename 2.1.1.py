import tkinter as tk
from tkinter import ttk, filedialog, messagebox, Menu, simpledialog
import pandas as pd
import random
import threading
import time
import json
import webbrowser
import os
from datetime import datetime
from PIL import Image, ImageTk
import logging
logging.basicConfig(
    filename='app.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.withdraw()  # 先隐藏主窗口
        
        # 提前初始化语言包和配置
        self.version = "2.1.1"
        self.language_pack = self.load_languages()  # 提前初始化语言包
        self.current_language = "zh_cn"  # 设置默认语言
        
        self.check_agreement()  # 检查用户协议

    def check_agreement(self):
        config = self.load_config()
        if config.get("agreed", False):
            self.continue_init()
            return
            
        self.agreement_window = tk.Toplevel()
        self.agreement_window.title(self.tr("agreement_title"))
        self.agreement_window.resizable(False, False)
        
        # 窗口居中
        window_width = 500
        window_height = 200
        screen_width = self.agreement_window.winfo_screenwidth()
        screen_height = self.agreement_window.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.agreement_window.geometry(f"{window_width}x{window_height}+{x}+{y}")

        main_frame = ttk.Frame(self.agreement_window)
        main_frame.pack(expand=True, fill='both', padx=20, pady=20)

        # 协议文字部分
        text_frame = ttk.Frame(main_frame)
        text_frame.pack(pady=10)
        
        ttk.Label(text_frame, text=self.tr("agreement_prompt_part1")).pack(side='left')
        link = ttk.Label(text_frame, text=self.tr("agreement_link"), 
                        foreground="blue", cursor="hand2")
        link.pack(side='left')
        link.bind("<Button-1>", lambda e: webbrowser.open(
            "https://docs.qq.com/doc/p/3e3bbf3683df96a76422af6730e77390cd665242"))
        ttk.Label(text_frame, text=self.tr("agreement_prompt_part2")).pack(side='left')

        # 按钮部分
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=15)
        
        ttk.Button(btn_frame, text=self.tr("agree"), 
                 command=self.on_agree).pack(side='left', padx=20)
        ttk.Button(btn_frame, text=self.tr("disagree"), 
                 command=self.on_disagree).pack(side='right', padx=20)

        self.agreement_window.protocol("WM_DELETE_WINDOW", self.on_disagree)
        self.agreement_window.grab_set()

    def on_agree(self):
        config = self.load_config()
        config["agreed"] = True
        self.save_config(config)
        self.agreement_window.destroy()
        self.continue_init()

    def on_disagree(self):
        self.agreement_window.destroy()
        self.destroy()

    def continue_init(self):
        self.deiconify()  # 显示主窗口
        self.title("Classroom Q&A System")
        self.state('zoomed')
        self.version = "2.1.1"
        self.current_language = "zh_cn"
        self.language_pack = self.load_languages()
        self.theme_config = self.load_theme("blue")
        self.last_file = self.load_config().get("last_file")
        self.bg_image = None

        self.setup_theme()
        self.create_container()
        self.create_pages()
        config = self.load_config()
        self.current_language = config["language"]
        self.theme_config = self.load_theme(config["theme"])
        self.apply_theme()
        self.bind("<Configure>", self.on_resize)

    def load_languages(self):
        return {
            "zh_cn": {
                "title": "课堂抽问系统",
                "upload": "上传名单文件",
                "select_file": "选择文件",
                "start": "开始抽取",
                "stop": "停止抽取",
                "speed": "速度调节",
                "reset": "重置系统",
                "all_list": "完整名单",
                "unasked": "未提问名单",
                "asked": "已提问名单",
                "about": "关于系统",
                "sponsor": "赞助作者",
                "lang": "语言设置",
                "theme": "主题设置",
                "menu_title": "系统菜单",
                "version": f"版本 {self.version}",
                "author": "开发者：赵钰兴",
                "load_success": "成功加载 {} 名学生",
                "load_error": "文件读取失败：{}",
                "reset_confirm": "已重置所有提问记录",
                "no_file_warn": "请先上传学生名单",
                "reload": "重新选择名单",
                "custom_bg": "自定义背景",
                "error": "错误",
                "bg_error": "无法加载背景图片",
                "stats": "统计信息",
                "total_students": "总人数",
                "asked_students": "已提问",
                "remaining_students": "剩余",
                "bilibili_btn": "访问作者B站主页",
                "all_asked_title": "提问完成",
                "all_asked_message": "所有学生均已完成提问！\n\n请重置系统以重新开始",
                "reset_success": "已重置所有提问记录",
                "agreement_title": "用户协议",
                "agreement_prompt_part1": "请同意",
                "agreement_link": "《课堂抽问系统使用条款》",
                "agreement_prompt_part2": "以继续使用本软件",
                "agree": "同意",
                "disagree": "拒绝"
            },
            "en": {
                "title": "Classroom Q&A System",
                "upload": "Upload File",
                "select_file": "Select File",
                "start": "Start Rolling",
                "stop": "Stop Rolling",
                "speed": "Rolling Speed",
                "reset": "Reset System",
                "all_list": "Full List",
                "unasked": "Unasked List",
                "asked": "Asked List",
                "about": "About",
                "sponsor": "Sponsor",
                "lang": "Language",
                "theme": "Themes",
                "menu_title": "System Menu",
                "version": f"Version {self.version}",
                "author": "Developer: Zhao Yuxing",
                "load_success": "Successfully loaded {} students",
                "load_error": "File error: {}",
                "reset_confirm": "Reset all records",
                "no_file_warn": "Please upload student list first",
                "reload": "Reload List",
                "custom_bg": "Custom Background",
                "error": "Error",
                "bg_error": "Failed to load background image",
                "stats": "Statistics",
                "total_students": "Total Students",
                "asked_students": "Asked Students",
                "remaining_students": "Remaining Students",
                "bilibili_btn": "Visit Bilibili Channel",
                "all_asked_title": "All Asked",
                "all_asked_message": "All students have been asked!\n\nPlease reset system to start over",
                "reset_success": "Reset all records successfully",
                "agreement_title": "User Agreement",
                "agreement_prompt_part1": "Please agree to the",
                "agreement_link": "《Classroom Q&A System Terms of Use》",
                "agreement_prompt_part2": "to continue",
                "agree": "Agree",
                "disagree": "Disagree"
            }
        }

    def setup_theme(self):
        self.style = ttk.Style()
        self.style.theme_use("clam")
        self.style.configure("TButton", font=("Arial", 12), padding=10)
        self.style.configure("Big.TButton", font=("Arial", 24, "bold"), padding=20)
        self.style.configure("TLabel", font=("Arial", 12))

    def create_container(self):
        self.container = ttk.Frame(self)
        self.container.pack(fill="both", expand=True)
        self.container.grid_rowconfigure(0, weight=1)
        self.container.grid_columnconfigure(0, weight=1)

    def create_pages(self):
        self.frames = {}
        for Page in (UploadPage, MainPage):
            frame = Page(self.container, self)
            self.frames[Page] = frame
            frame.grid(row=0, column=0, sticky="nsew")
        if self.last_file:
            self.show_page(MainPage, self.last_file)
        else:
            self.show_page(UploadPage)

    def show_page(self, page_class, file_path=None):
        frame = self.frames[page_class]
        if file_path and hasattr(frame, "load_file"):
            frame.load_file(file_path)
        frame.tkraise()
        self.apply_theme()

    def load_theme(self, theme_name):
        themes = {
            "blue": {"bg": "#E3F2FD", "fg": "#0D47A1", "accent": "#2196F3"},
            "green": {"bg": "#E8F5E9", "fg": "#1B5E20", "accent": "#4CAF50"},
            "orange": {"bg": "#FFF3E0", "fg": "#E65100", "accent": "#FF9800"},
            "purple": {"bg": "#F3E5F5", "fg": "#4A148C", "accent": "#9C27B0"},
            "red": {"bg": "#FFEBEE", "fg": "#B71C1C", "accent": "#F44336"},
            "pink": {"bg": "#FCE4EC", "fg": "#880E4F", "accent": "#E91E63"},
            "cyan": {"bg": "#E0F7FA", "fg": "#006064", "accent": "#00BCD4"},
            "midnight": {"bg": "#2C3E50", "fg": "#ECF0F1", "accent": "#3498DB"},
            "sunset": {"bg": "#F39C12", "fg": "#2C3E50", "accent": "#E74C3C"},
            "forest": {"bg": "#1E824C", "fg": "#ECF0F1", "accent": "#F4D03F"},
            "ocean": {"bg": "#1ABC9C", "fg": "#34495E", "accent": "#F1C40F"},
            "lavender": {"bg": "#9B59B6", "fg": "#ECF0F1", "accent": "#8E44AD"},
            "coral": {"bg": "#FF6F61", "fg": "#2C3E50", "accent": "#FFD166"},
            "slate": {"bg": "#34495E", "fg": "#ECF0F1", "accent": "#95A5A6"},
            "sand": {"bg": "#F4D03F", "fg": "#2C3E50", "accent": "#F5B041"},
            "custom": {"bg": "#FFFFFF", "fg": "#000000", "accent": "#2196F3", "image": None}
        }
        return themes.get(theme_name, themes["blue"])

    def apply_theme(self):
        theme = self.theme_config
        if theme.get("image"):
            self.set_background_image(theme["image"])
        else:
            self.configure(bg=theme["bg"])
            self.style.configure("TFrame", background=theme["bg"])
            self.style.configure("TLabel", background=theme["bg"], foreground=theme["fg"])
            self.style.configure("TButton", background=theme["accent"], foreground="white")
            self.style.map("TButton",
                           background=[("active", theme["accent"]), ("disabled", "#CCCCCC")])

    def load_config(self):
        default_config = {
            "last_file": "",
            "language": "zh_cn",
            "theme": "blue",
            "default_column": "",
            "agreed": False
        }
        
        if not os.path.exists("config.json"):
            with open("config.json", "w", encoding="utf-8") as f:
                json.dump(default_config, f)
            return default_config
            
        try:
            with open("config.json", "r", encoding="utf-8") as f:
                config = json.load(f)
                # 合并缺失的配置项
                for key in default_config:
                    if key not in config:
                        config[key] = default_config[key]
                # 保存更新后的配置
                with open("config.json", "w", encoding="utf-8") as f_write:
                    json.dump(config, f_write)
                return config
        except Exception as e:
            logging.error(f"配置加载失败: {str(e)}")
            return default_config

    def set_background_image(self, image_path):
        try:
            img = Image.open(image_path)
            img = img.resize((self.winfo_width(), self.winfo_height()), Image.Resampling.LANCZOS)
            self.bg_image = ImageTk.PhotoImage(img)
            for frame in self.frames.values():
                frame.configure(image=self.bg_image)
        except Exception as e:
            messagebox.showerror(self.tr("error"), f"{self.tr('bg_error')}: {str(e)}")

    def save_config(self, config_data):
        with open("config.json", "w", encoding="utf-8") as f:
            json.dump(config_data, f, ensure_ascii=False)

    def tr(self, key):
        return self.language_pack[self.current_language][key]

    def on_resize(self, event):
        if self.theme_config.get("image"):
            self.set_background_image(self.theme_config["image"])

class UploadPage(ttk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.setup_ui()

    def setup_ui(self):
        self.label = ttk.Label(self, 
                             text=self.tr("upload"),
                             font=("Arial", 24, "bold"))
        self.btn_upload = ttk.Button(self,
                                   text=self.tr("select_file"),
                                   command=self.upload_file,
                                   style="Big.TButton")
        self.label.pack(pady=100)
        self.btn_upload.pack()

    def tr(self, key):
        return self.controller.language_pack[self.controller.current_language][key]

    def upload_file(self):
        filetypes = [
            ("Excel文件", "*.xlsx"),
            ("CSV文件", "*.csv"),
            ("所有文件", "*.*")
        ]
        file_path = filedialog.askopenfilename(filetypes=filetypes)
        if file_path:
            config = self.controller.load_config()
            config["last_file"] = file_path
            self.controller.save_config(config)
            self.controller.show_page(MainPage, file_path)

class MainPage(ttk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.original_students = []
        self.students = []
        self.asked = []
        self.is_rolling = False
        self.roll_speed = 0.01
        self.setup_ui()
        self.create_menu()

    def load_file(self, file_path):
        try:
            actual_path = str(file_path.get("last_file", "")) if isinstance(file_path, dict) else str(file_path)
            if not actual_path.strip():
                raise ValueError("无效文件路径")

            if actual_path.lower().endswith('.csv'):
                df = pd.read_csv(actual_path, encoding='utf-8')
            else:
                df = pd.read_excel(actual_path)

            columns_list = df.columns.tolist()
            config = self.controller.load_config()
            default_column = config.get("default_column", columns_list if columns_list else "")
            
            selected_col = simpledialog.askstring(
                "选择列",
                f"可用列：{', '.join(columns_list)}\n输入列名（默认使用'{default_column}'）：",
                initialvalue=default_column
            ) or default_column

            config["default_column"] = selected_col
            self.controller.save_config(config)

            self.original_students = list(dict.fromkeys(df[selected_col].astype(str).tolist()))
            self.students = self.original_students.copy()
            random.shuffle(self.students)
            messagebox.showinfo("成功", self.tr("load_success").format(len(self.students)))

        except KeyError as ke:
            messagebox.showerror("列错误", f"无效列名，请选择有效列")
        except pd.errors.EmptyDataError:
            messagebox.showerror("数据错误", "文件内容为空")
        except Exception as e:
            messagebox.showerror("错误", self.tr("load_error").format(str(e)))

    def tr(self, key):
        return self.controller.language_pack[self.controller.current_language][key]

    def setup_ui(self):
        display_font = ("Arial", 72, "bold")
        self.result_display = ttk.Label(self,
                                      font=display_font,
                                      anchor="center",
                                      background="white")
        self.result_display.pack(fill="both", expand=True, padx=20, pady=20)

        control_frame = ttk.Frame(self)
        control_frame.pack(fill="x", padx=20, pady=10)

        left_control = ttk.Frame(control_frame)
        left_control.pack(side="left", padx=(0, 20), fill="x", expand=True)

        self.btn_pick = ttk.Button(left_control,
                                  text=self.tr("start"),
                                  style="Big.TButton",
                                  command=self.toggle_rolling)
        self.btn_pick.pack(side="left", padx=(0, 10))

        speed_frame = ttk.Frame(left_control)
        speed_frame.pack(side="left", fill="x", expand=True)

        ttk.Label(speed_frame, 
                 text=self.tr("speed"),
                 font=("Arial", 10)).pack(side="left")
        self.speed_scale = ttk.Scale(speed_frame,
                                    from_=0.01, to=0.2,
                                    command=self.update_speed)
        self.speed_scale.set(self.roll_speed)
        self.speed_scale.pack(side="left", fill="x", expand=True, padx=10)

        self.menu_btn = ttk.Button(control_frame, 
                                  text=self.tr("menu_title"),
                                  command=self.show_menu)
        self.menu_btn.pack(side="right")

    def toggle_rolling(self):
        if not self.original_students:
            messagebox.showwarning(self.tr("error"), self.tr("no_file_warn"))
            return
            
        if not self.students and not self.is_rolling:
            messagebox.showinfo(
                self.tr("all_asked_title"),
                self.tr("all_asked_message")
            )
            return

        self.is_rolling = not self.is_rolling
        self.btn_pick.config(text=self.tr("stop") if self.is_rolling else self.tr("start"))
        if self.is_rolling:
            threading.Thread(target=self.roll_names, daemon=True).start()
        else:
            self.pick_student()

    def roll_names(self):
        try:
            while self.is_rolling:
                name = random.choice(self.students)
                self.result_display.config(text=name)
                time.sleep(self.roll_speed)
        except Exception as e:
            messagebox.showerror(self.tr("error"), f"Thread error: {str(e)}")

    def pick_student(self):
        if not self.students:
            self.is_rolling = False
            self.btn_pick.config(text=self.tr("start"))
            return
            
        if self.students:
            selected = random.choice(self.students)
            self.students.remove(selected)
            self.asked.append(selected)
            self.result_display.config(text=selected)

            with open("抽中记录.txt", "a", encoding="utf-8") as f:
                f.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} 抽中：{selected}\n")

    def update_speed(self, value):
        self.roll_speed = float(value)

    def create_menu(self):
        self.menu = Menu(self, tearoff=0)
        self.menu.add_command(label=self.tr("reload"), command=self.reload_file)
        self.menu.add_command(label=self.tr("reset"), command=self.reset_system)
        self.menu.add_command(label=self.tr("all_list"), command=self.show_all_students)
        self.menu.add_command(label=self.tr("unasked"), command=self.show_unasked)
        self.menu.add_command(label=self.tr("asked"), command=self.show_asked)
        self.menu.add_command(label=self.tr("stats"), command=self.show_statistics)

        lang_menu = Menu(self.menu, tearoff=0)
        lang_menu.add_command(label="简体中文", command=lambda: self.change_language("zh_cn"))
        lang_menu.add_command(label="English", command=lambda: self.change_language("en"))
        self.menu.add_cascade(label=self.tr("lang"), menu=lang_menu)

        theme_menu = Menu(self.menu, tearoff=0)
        theme_menu.add_command(label="蓝色主题", command=lambda: self.change_theme("blue"))
        theme_menu.add_command(label="绿色主题", command=lambda: self.change_theme("green"))
        theme_menu.add_command(label="橙色主题", command=lambda: self.change_theme("orange"))
        theme_menu.add_command(label="紫色主题", command=lambda: self.change_theme("purple"))
        theme_menu.add_command(label="红色主题", command=lambda: self.change_theme("red"))
        theme_menu.add_command(label="粉色主题", command=lambda: self.change_theme("pink"))
        theme_menu.add_command(label="青色主题", command=lambda: self.change_theme("cyan"))
        theme_menu.add_command(label="午夜主题", command=lambda: self.change_theme("midnight"))
        theme_menu.add_command(label="日落主题", command=lambda: self.change_theme("sunset"))
        theme_menu.add_command(label="森林主题", command=lambda: self.change_theme("forest"))
        theme_menu.add_command(label="海洋主题", command=lambda: self.change_theme("ocean"))
        theme_menu.add_command(label="薰衣草主题", command=lambda: self.change_theme("lavender"))
        theme_menu.add_command(label="珊瑚主题", command=lambda: self.change_theme("coral"))
        theme_menu.add_command(label="石板主题", command=lambda: self.change_theme("slate"))
        theme_menu.add_command(label="沙滩主题", command=lambda: self.change_theme("sand"))
        theme_menu.add_command(label=self.tr("custom_bg"), command=self.set_custom_bg)
        self.menu.add_cascade(label=self.tr("theme"), menu=theme_menu)

        self.menu.add_command(label=self.tr("about"), command=self.show_about)
        self.menu.add_command(label=self.tr("sponsor"),
                             command=lambda: webbrowser.open("https://ifdian.net/a/schoolqa"))

    def set_custom_bg(self):
        filetypes = [("图片文件", "*.jpg *.jpeg *.png")]
        file_path = filedialog.askopenfilename(filetypes=filetypes)
        if file_path:
            self.controller.theme_config = self.controller.load_theme("custom")
            self.controller.theme_config["image"] = file_path
            self.controller.apply_theme()

    def change_language(self, lang):
        self.controller.current_language = lang
        config = self.controller.load_config()
        config["language"] = lang
        self.controller.save_config(config)
        self.update_ui_text()
        self.controller.apply_theme()

    def change_theme(self, theme_name):
        self.controller.theme_config = self.controller.load_theme(theme_name)
        config = self.controller.load_config()
        config["theme"] = theme_name
        self.controller.save_config(config)
        self.controller.apply_theme()

    def update_ui_text(self):
        self.btn_pick.config(text=self.tr("start") if not self.is_rolling else self.tr("stop"))
        self.menu_btn.config(text=self.tr("menu_title"))
        self.menu.entryconfig(0, label=self.tr("reload"))
        self.menu.entryconfig(1, label=self.tr("reset"))
        self.menu.entryconfig(2, label=self.tr("all_list"))
        self.menu.entryconfig(3, label=self.tr("unasked"))
        self.menu.entryconfig(4, label=self.tr("asked"))
        self.menu.entryconfig(5, label=self.tr("stats"))
        self.menu.entryconfig(7, label=self.tr("lang"))
        self.menu.entryconfig(8, label=self.tr("theme"))
        self.menu.entryconfig(10, label=self.tr("about"))
        self.menu.entryconfig(11, label=self.tr("sponsor"))

    def reload_file(self):
        filetypes = [
            ("Excel文件", "*.xlsx"),
            ("CSV文件", "*.csv"),
            ("所有文件", "*.*")
        ]
        file_path = filedialog.askopenfilename(filetypes=filetypes)
        if file_path:
            self.controller.save_config({"last_file": file_path})
            self.load_file(file_path)
            self.reset_system()

    def reset_system(self):
        self.students = self.original_students.copy()
        random.shuffle(self.students)
        self.asked = []
        self.result_display.config(text="")
        messagebox.showinfo(self.tr("reset"), self.tr("reset_success"))

    def show_all_students(self):
        msg = "全部学生名单：\n\n" + "\n".join(self.original_students)
        messagebox.showinfo("完整名单", msg)

    def show_unasked(self):
        if not self.students:
            messagebox.showinfo("未提问名单", "所有学生均已完成提问")
        else:
            msg = f"剩余{len(self.students)}人：\n" + "\n".join(self.students)
            messagebox.showinfo("未提问名单", msg)

    def show_asked(self):
        if not self.asked:
            messagebox.showinfo("已提问名单", "暂无提问记录")
        else:
            msg = f"已提问{len(self.asked)}人：\n" + "\n".join(self.asked)
            messagebox.showinfo("已提问名单", msg)

    def show_statistics(self):
        total = len(self.original_students)
        asked = len(self.asked)
        remaining = total - asked
        msg = (f"{self.tr('total_students')}: {total}\n"
               f"{self.tr('asked_students')}: {asked}\n"
               f"{self.tr('remaining_students')}: {remaining}")
        messagebox.showinfo(self.tr("stats"), msg)

    def show_about(self):
        about_window = tk.Toplevel(self)
        about_window.title(self.tr("about"))
        
        title_font = ("Arial", 14, "bold")
        content_frame = ttk.Frame(about_window)
        content_frame.pack(padx=20, pady=20)

        version_label = ttk.Label(
            content_frame,
            text=f"{self.tr('version')}\n{self.tr('author')}",
            font=title_font,
            justify="center"
        )
        version_label.pack(pady=10)

        bilibili_btn = ttk.Button(
            content_frame,
            text=self.tr("bilibili_btn"),
            command=lambda: webbrowser.open("https://space.bilibili.com/3546648932256226"),
            style="TButton"
        )
        bilibili_btn.pack(pady=5, ipadx=10, ipady=5)

        about_window.update_idletasks()
        width = 320
        height = 200
        x = (about_window.winfo_screenwidth() // 2) - (width // 2)
        y = (about_window.winfo_screenheight() // 2) - (height // 2)
        about_window.geometry(f'{width}x{height}+{x}+{y}')

    def show_menu(self):
        try:
            self.menu.tk_popup(self.menu_btn.winfo_rootx(),
                              self.menu_btn.winfo_rooty() + 30)
        finally:
            self.menu.grab_release()

if __name__ == "__main__":
    app = App()
    app.mainloop()
    app.mainloop()