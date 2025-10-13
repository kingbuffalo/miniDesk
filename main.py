import tkinter as tk
import os
import sys
import traceback
import MiniQDesktop
import os
import psutil  # 需要安装 psutil 库：pip install psutil

def check_single_instance():
    lock_file = "app.lock"
    if os.path.exists(lock_file):
        # 读取锁文件中的 PID
        with open(lock_file, "r") as f:
            pid = int(f.read().strip())
        
        # 检查 PID 是否有效
        if psutil.pid_exists(pid):
            print("程序已经在运行中，不能重复启动。")
            sys.exit(0)  # 退出程序
        else:
            # 如果 PID 无效，删除锁文件
            os.remove(lock_file)

    # 创建新的锁文件，写入当前进程的 PID
    with open(lock_file, "w") as f:
        f.write(str(os.getpid()))

def remove_lock_file():
    lock_file = "app.lock"
    if os.path.exists(lock_file):
        os.remove(lock_file)


if __name__ == "__main__":
    # 检查是否已有实例运行
    check_single_instance()

    try:
        # 获取屏幕高度(用于隐藏按钮定位)
        root = tk.Tk()
        root.withdraw()  # 隐藏临时窗口
        screen_height = root.winfo_screenheight()
        root.update_idletasks()
        root.destroy()  # 销毁临时窗口

        app = MiniQDesktop.MiniQDesktop()
        app.setup_global_hotkey()
        app.root.protocol("WM_DELETE_WINDOW", app.hide_window)
        app.root.mainloop()

    except Exception as e:
        with open("error.log", "a", encoding="utf-8") as f:
            f.write(traceback.format_exc())
    finally:
        # 程序退出时删除锁文件
        remove_lock_file()