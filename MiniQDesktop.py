import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import keyboard
import json
from tkinter import simpledialog
from PIL import Image, ImageTk
import pystray
import threading
import ctypes


def get_icon_image(path, size=(32, 32)):
    """提取快捷方式或可执行文件的图标，或为文件夹使用默认图标"""
    if os.path.isdir(path):
        # 使用默认文件夹图标
        icon_path = os.path.join(os.getcwd(), "default_folder_icon.png")
        if not os.path.exists(icon_path):
            # 如果默认图标文件不存在，创建一个简单的占位图标
            icon = Image.new("RGB", size, "gray")
            return ImageTk.PhotoImage(icon)
        else:
            # 使用默认文件夹图标
            icon = Image.open(icon_path)
            icon = icon.resize(size, Image.Resampling.LANCZOS)
            return ImageTk.PhotoImage(icon)
    else:
        # 提取文件的图标
        temp_icon_path = os.path.join(os.getcwd(), "temp_file.png")
        try:
            ctypes.windll.shell32.ExtractIconExW(path, 0, None, ctypes.pointer(ctypes.c_wchar_p(temp_icon_path)), 1)
            icon = Image.open(temp_icon_path)
            icon = icon.resize(size, Image.Resampling.LANCZOS)
            return ImageTk.PhotoImage(icon)
        except Exception as e:
            print(f"无法提取图标: {e}")
            return None
        finally:
            # 删除临时图标文件
            if os.path.exists(temp_icon_path):
                os.remove(temp_icon_path)


class MiniQDesktop:
    def __init__(self):
        # 主窗口设置
        self.root = tk.Tk()
        self.root.title("MiniQ书桌")
        self.root.attributes('-alpha', 0.95)  # 半透明效果
        self.root.attributes('-topmost', True)  # 置顶
        
        # 存储快捷方式数据
        self.shortcuts_file = "shortcuts.json"
        self.shortcuts_data = self.load_shortcuts()
        
        # 初始化UI
        self.setup_ui()
        
        # 系统托盘图标
        self.setup_tray_icon()
        
        # 初始位置(屏幕右下角)
        self.position_window()
        
        # 窗口拖动相关变量
        self._drag_data = {"x": 0, "y": 0}
        
    def position_window(self):
        """将窗口定位在屏幕右下角"""
        self.screen_width = self.root.winfo_screenwidth()
        self.screen_height = self.root.winfo_screenheight()
        window_width = 1000
        window_height = 400
        #x = screen_width - window_width - 50  # 距离右边50像素
        x = 0
        y = self.screen_height - window_height - 60  # 距离底部50像素
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
    
    def setup_ui(self):
        """设置用户界面"""
        # 主容器
        main_frame = ttk.Frame(self.root)
        main_frame.pack(expand=True, fill="both", padx=5, pady=5)
        
        # 标题栏
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill="x")
        
        # 拖动区域
        self.drag_label = ttk.Label(title_frame, text="MiniQ书桌", font=('Arial', 10))
        self.drag_label.pack(side="left", fill="x", expand=True)
        
        # 绑定拖动事件
        # self.drag_label.bind("<ButtonPress-1>", self.start_drag)
        # self.drag_label.bind("<B1-Motion>", self.on_drag)
        
        # 分组容器
        self.group_container = ttk.Frame(main_frame)
        self.group_container.pack(expand=True, fill="both")
        
        # 控制按钮
        control_frame = ttk.Frame(main_frame)
        control_frame.pack(fill="x", pady=(5, 0))
        
        add_btn = ttk.Button(control_frame, text="+", command=self.add_shortcut)
        add_btn.pack(side="left", padx=2)

        add_btn = ttk.Button(control_frame, text="++", command=self.add_shortcut2)
        add_btn.pack(side="left", padx=2)
        
        add_group_btn = ttk.Button(control_frame, text="new group", command=self.add_group)
        add_group_btn.pack(side="left", padx=2)
        
        # 加载已有分组
        self.load_groups()
        
        # 隐藏窗口按钮(模拟开始菜单旁边的唤出按钮)
        self.hidden_btn = tk.Button(self.root, text="MQ", bg="lightblue", command=self.toggle_window, bd=0)
        self.hidden_btn.place(x=-30, y=0, width=30, height=30)
        
        # 绑定窗口事件
        self.root.bind("<Unmap>", lambda e: self.hidden_btn.place(x=-30, y=0))
        self.root.bind("<Map>", lambda e: self.hidden_btn.place_forget())
        
    def setup_tray_icon(self):
        """设置系统托盘图标"""
        icon_path = "icons/icon.png"  # 确保文件路径正确
        if os.path.exists(icon_path):
            image = Image.open(icon_path)
        else:
            # 如果图标文件不存在，使用默认的白色图标
            image = Image.new('RGB', (64, 64), "white")
        self.tray_icon = pystray.Icon("mini_q", image, "MiniQ书桌", 
                                     menu=pystray.Menu(
                                         pystray.MenuItem("Show", self.show_window),
                                         pystray.MenuItem("Quit", self.quit_app)
                                     ))
        
        # 在单独线程中运行系统托盘
        threading.Thread(target=self.tray_icon.run, daemon=True).start()
    
    # def start_drag(self, event):
    #     """开始拖动窗口"""
    #     self._drag_data["x"] = event.x
    #     self._drag_data["y"] = event.y
    
    # def on_drag(self, event):
    #     """拖动窗口"""
    #     x = self.root.winfo_x() + (event.x - self._drag_data["x"])
    #     y = self.root.winfo_y() + (event.y - self._drag_data["y"])
    #     self.root.geometry(f"+{x}+{y}")
    
    
    def show_window(self, *args):
        """显示窗口"""
        self.root.deiconify()
        self.hidden_btn.place_forget()
    
    def hide_window(self, *args):
        """隐藏窗口"""
        self.root.withdraw()
        # 将唤出按钮移动到屏幕左侧
        self.hidden_btn.place(x=0, y=self.screen_height//2, width=30, height=60)
    
    def load_shortcuts(self):
        """加载快捷方式数据"""
        if os.path.exists(self.shortcuts_file):
            with open(self.shortcuts_file, "r", encoding="utf-8") as f:
                return json.load(f)
        return {"groups": {}}
    
    def save_shortcuts(self):
        """保存快捷方式数据"""
        with open(self.shortcuts_file, "w", encoding="utf-8") as f:
            json.dump(self.shortcuts_data, f, ensure_ascii=False, indent=2)
    
    def load_groups(self):
        """加载所有分组"""
        # 清空现有分组
        for widget in self.group_container.winfo_children():
            widget.destroy()
        
        # 加载保存的分组
        for group_name, shortcuts in self.shortcuts_data["groups"].items():
            self.create_group_frame(group_name, shortcuts)
    
    def create_group_frame(self, group_name, shortcuts=None):
        """创建一个新的分组框架"""
        if shortcuts is None:
            shortcuts = []
        
        # 分组框架
        group_frame = ttk.LabelFrame(self.group_container, text=group_name)
        group_frame.pack(fill="x", pady=5, padx=5)
        
        # 快捷方式容器
        shortcuts_frame = ttk.Frame(group_frame)
        shortcuts_frame.pack(fill="x", padx=5, pady=5)
        
        # 添加已有快捷方式
        for shortcut in shortcuts:
            self.create_shortcut_button(shortcuts_frame, shortcut["name"], shortcut["path"])
        
        # 添加"+"按钮
        add_btn = ttk.Button(shortcuts_frame, text="+", width=3,command=lambda: self.add_shortcut_to_group(group_name))
        add_btn.pack(side="left", padx=2)
        add_btn2 = ttk.Button(shortcuts_frame, text="++", width=3,command=lambda: self.add_shortcut_to_group2(group_name))
        add_btn2.pack(side="left", padx=2)
        
        # 删除分组按钮
        del_btn = ttk.Button(group_frame, text="X", command=lambda: self.delete_group(group_name))
        del_btn.pack(side="right", pady=(0, 5))
        
        # 保存分组到数据
        if group_name not in self.shortcuts_data["groups"]:
            self.shortcuts_data["groups"][group_name] = []
            self.save_shortcuts()

    def wrap_text(self,text, max_length):
        return '\n'.join([text[i:i+max_length] for i in range(0, len(text), max_length)])
    
    def create_shortcut_button(self, parent, name, path):
        """创建一个快捷方式按钮"""
        icon_image = get_icon_image(path)
        btn = ttk.Button(parent, text=self.wrap_text(name,8), command=lambda: self.open_shortcut(path), width=8)
        if icon_image:
            btn.config(image=icon_image, compound="top")  # 图标在上，文本在下
            btn.image = icon_image  # 防止图像被垃圾回收
        btn.pack(side="left", padx=2, pady=5)  # 调整外边距
        # btn['padding'] = (5, 10)  # 设置内边距 (左右, 上下)
        
        # 右键菜单
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="打开", command=lambda: self.open_shortcut(path))
        menu.add_command(label="删除", command=lambda: self.delete_shortcut(parent, btn, path))
        
        btn.bind("<Button-3>", lambda e: menu.post(e.x_root, e.y_root))
    
    def open_shortcut(self, path):
        """打开快捷方式"""
        try:
            if os.path.isdir(path):
                os.startfile(path)
            else:
                os.startfile(path)
        except Exception as e:
            messagebox.showerror("错误", f"无法打开: {e}")

        #关闭
        self.hide_window()
    
    def delete_shortcut(self, parent, button, path):
        """删除快捷方式"""
        button.destroy()
        # 从所有分组中删除该快捷方式
        for group in self.shortcuts_data["groups"].values():
            group[:] = [s for s in group if s["path"] != path]
        self.save_shortcuts()
    
    def delete_group(self, group_name):
        """删除整个分组"""
        if messagebox.askyesno("确认", f"确定要删除分组 '{group_name}' 吗?"):
            del self.shortcuts_data["groups"][group_name]
            self.save_shortcuts()
            self.load_groups()

    def add_shortcut2(self):
        """添加新的快捷方式"""
        path = filedialog.askdirectory(title="选择文件夹")
        if not path:
            return
        
        name = os.path.basename(path)
        # 让用户选择分组
        self.ask_for_group(name, path)
    
    def add_shortcut(self):
        """添加新的快捷方式"""
        path = filedialog.askopenfilename(title="选择文件")
        if not path:
            return
        
        name = os.path.basename(path)
        # 让用户选择分组
        self.ask_for_group(name, path)

    def add_shortcut_to_group2(self, group_name):
        """向指定分组添加快捷方式"""
        path = filedialog.askdirectory(title="选择文件夹")
        if not path:
            return
        
        name = os.path.basename(path)
        self.add_shortcut_to_group_with_name(group_name, name, path)
    
    def add_shortcut_to_group(self, group_name):
        """向指定分组添加快捷方式"""
        path = filedialog.askopenfilename(title="选择文件")
        if not path:
            return
        
        name = os.path.basename(path)
        self.add_shortcut_to_group_with_name(group_name, name, path)
    
    def add_shortcut_to_group_with_name(self, group_name, name, path):
        """向指定分组添加快捷方式(指定名称)"""
        # 检查是否已存在
        for group in self.shortcuts_data["groups"].values():
            if any(s["path"] == path for s in group):
                messagebox.showwarning("警告", "该快捷方式已存在!")
                return
        
        if group_name not in self.shortcuts_data["groups"]:
            self.shortcuts_data["groups"][group_name] = []
        # 添加到分组
        self.shortcuts_data["groups"][group_name].append({"name": name, "path": path})
        self.save_shortcuts()
        self.load_groups()
    
    def ask_for_group(self, name, path):
        """询问将快捷方式添加到哪个分组"""
        top = tk.Toplevel(self.root)
        top.title("选择分组")
        top.transient(self.root)
        top.grab_set()
        
        ttk.Label(top, text=f"将 '{name}' 添加到:").pack(pady=5)
        
        # 现有分组选项
        group_var = tk.StringVar()
        for group in self.shortcuts_data["groups"]:
            ttk.Radiobutton(top, text=group, value=group, variable=group_var).pack(anchor="w")
        
        # 新建分组选项
        ttk.Radiobutton(top, text="新建分组:", value="new", variable=group_var).pack(anchor="w")
        new_group_entry = ttk.Entry(top)
        new_group_entry.pack(fill="x", padx=20, pady=(0, 10))
        
        def on_confirm():
            selected_group = group_var.get()
            if selected_group == "new":
                new_group_name = new_group_entry.get().strip()
                if not new_group_name:
                    messagebox.showerror("错误", "请输入新分组名称!")
                    return
                if new_group_name in self.shortcuts_data["groups"]:
                    messagebox.showerror("错误", "分组已存在!")
                    return
                self.shortcuts_data["groups"][new_group_name] = []
                self.add_shortcut_to_group_with_name(new_group_name, name, path)
            else:
                self.add_shortcut_to_group_with_name(selected_group, name, path)
            top.destroy()
        
        ttk.Button(top, text="确定", command=on_confirm).pack(pady=5)
    
    def add_group(self):
        """添加新的分组"""
        name = simpledialog.askstring("新建分组", "输入分组名称:")
        if name and name.strip():
            if name in self.shortcuts_data["groups"]:
                messagebox.showerror("错误", "分组已存在!")
                return
            self.shortcuts_data["groups"][name] = []
            self.save_shortcuts()
            self.load_groups()
    
    def quit_app(self):
        """退出应用程序"""
        self.tray_icon.stop()
        self.root.quit()
        self.root.destroy()


    def setup_global_hotkey(self):
        """设置全局快捷键"""
        # threading.Thread(target=lambda: keyboard.add_hotkey('alt+z', self.toggle_window), daemon=True).start()
        # threading.Thread(target=lambda: keyboard.add_hotkey('alt+z', self.toggle_window), daemon=True).start()
        keyboard.add_hotkey('alt+z', self.toggle_window)

    def toggle_window(self):
        """切换窗口显示状态"""
        current_state = self.root.state()
        if current_state in ("withdrawn", "iconic"):  # 隐藏或最小化状态
            self.show_window()
        else:  # 其他状态（如正常显示）
            self.hide_window()