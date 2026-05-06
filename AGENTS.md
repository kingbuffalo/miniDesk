# MiniQDesktop (MiniQ书桌) - Agent 开发指南

本文档面向 AI 编程 Agent，介绍本项目的架构、技术栈和开发约定。阅读本文档前，请假设你对本项目一无所知。

---

## 项目概述

MiniQDesktop（MiniQ书桌）是一个基于 Python 和 Tkinter 的 Windows 桌面快捷方式管理工具。项目缘起于腾讯产品"小Q书桌"下架，作者希望延续类似的使用体验。本项目绝大部分代码由 DeepSeek AI 生成，作者仅负责修复运行时错误和优化使用习惯。

**核心功能：**
- 文件和文件夹快捷方式的分组管理
- 分组支持创建、删除、展开/折叠（互斥展开，一次只展开一个分组）
- 系统托盘图标（最小化到托盘，右键菜单显示/退出）
- 全局快捷键 `Alt+Z` 快速唤起/隐藏主窗口
- 单实例运行（通过 `app.lock` 锁文件实现）
- 窗口默认显示在屏幕左下角，宽度 1000 像素，支持鼠标拖动

**运行环境：** Windows 系统（代码中使用了 `ctypes.windll.shell32.ExtractIconExW`、`os.startfile` 等 Windows 专属 API）。

---

## 技术栈

| 组件 | 用途 | 版本 |
|------|------|------|
| Python | 运行时 | 3.11（项目使用 Python 3.11.9 的 venv） |
| tkinter | GUI 框架（Python 内置） | 内置 |
| Pillow | 图像处理（图标加载、缩放） | 9.3.0 |
| pystray | 系统托盘图标 | 0.19.4 |
| keyboard | 全局键盘快捷键监听 | 0.13.5 |
| psutil | 进程管理（单实例检查） | 未在 requirements.txt 中固定版本 |
| ctypes | 调用 Windows Shell API 提取文件图标 | 内置 |

---

## 项目结构

```
miniDesk/
├── main.py              # 主程序入口：单实例检查、全局异常捕获、启动主循环
├── MiniQDesktop.py      # 核心功能模块：MiniQDesktop 类，包含全部 UI 和业务逻辑
├── shortcuts.json       # 用户数据：快捷方式分组配置（运行后自动生成/更新）
├── requirements.txt     # Python 依赖列表
├── run.bat              # Windows 启动脚本（使用 pythonw 无窗口运行）
├── app.lock             # 单实例锁文件（运行时生成，存储 PID）
├── readme.md            # 面向用户的中文项目说明
├── AGENTS.md            # 本文件
├── .venv/               # Python 虚拟环境（Python 3.11）
└── icons/               # 图标资源
    ├── default_folder_icon.png   # 文件夹默认图标
    ├── exe_icon.png              # .exe 文件默认图标
    └── icon.png                  # 系统托盘图标
```

**注意：** 本项目没有 `pyproject.toml`、`setup.py`、`package.json` 或类似的构建配置文件。它就是一个直接的 Python 脚本项目，无需打包构建即可运行。

---

## 模块说明

### main.py（入口文件）

职责：
1. 调用 `check_single_instance()` 检查是否已有实例运行：
   - 读取 `app.lock` 中的 PID，用 `psutil.pid_exists()` 验证
   - 若进程存在，直接退出；否则删除旧锁文件并创建新锁
2. 创建临时 `tk.Tk()` 获取屏幕高度（供主程序定位使用）
3. 实例化 `MiniQDesktop.MiniQDesktop()`，注册全局热键，启动 `mainloop()`
4. 全局异常捕获：写入 `error.log`
5. `finally` 块中删除 `app.lock`

### MiniQDesktop.py（核心模块）

包含一个类 `MiniQDesktop` 和一个工具函数 `get_icon_image()`。

**`get_icon_image(path, size=(32, 32))`**
- 对文件夹：返回 `icons/default_folder_icon.png`
- 对 `.exe` 文件：返回 `icons/exe_icon.png`
- 对其他文件：尝试用 `ctypes.windll.shell32.ExtractIconExW` 提取图标，生成临时文件 `temp_file.png`，加载后删除
- 失败时返回 `None`

**`MiniQDesktop` 类主要方法：**
- `__init__()`：初始化主窗口（半透明、置顶）、加载数据、设置 UI、托盘、窗口位置
- `setup_ui()`：构建界面——标题栏、带滚动条的画布、分组网格容器（每行最多 4 列）、控制按钮（`+` 文件、`++` 文件夹、`new group`）
- `setup_tray_icon()`：启动 `pystray.Icon` 守护线程
- `position_window()`：窗口定位在屏幕左下角（`x=0`, `y=screen_height - 400 - 60`）
- `load_groups()` / `create_group_frame()`：加载分组，每个分组可折叠/展开，展开时自动折叠其他分组
- `create_shortcut_button()`：创建快捷方式按钮（图标在上，文字在下，最多 8 字符换行），支持右键菜单（打开 / goto / 删除）
- `add_shortcut()` / `add_shortcut2()`：弹出文件/文件夹选择对话框，然后让用户选择或新建分组
- `save_shortcuts()` / `load_shortcuts()`：读写 `shortcuts.json`
- `setup_global_hotkey()`：注册 `alt+z` 全局热键，调用 `toggle_window()`
- `hide_window()` / `show_window()`：隐藏窗口时，屏幕左侧会出现一个 `MQ` 唤出按钮

**数据格式（`shortcuts.json`）：**
```json
{
  "groups": {
    "分组名": [
      {"name": "显示名称", "path": "绝对路径"}
    ]
  }
}
```

---

## 运行方式

### 开发/调试运行

```bash
# 安装依赖
pip install -r requirements.txt
pip install psutil

# 直接运行（会弹出控制台窗口）
python main.py
```

### 日常运行（推荐）

```bash
# 使用 run.bat（通过 pythonw 运行，无控制台窗口）
run.bat
```

`run.bat` 内容只有一行：`start "" pythonw main.py`

---

## 开发约定与注意事项

### 平台限制
- **仅支持 Windows**。代码中大量使用了 Windows 专属 API，不要尝试移植到 Linux/macOS。
- `keyboard` 模块监听全局热键**可能需要管理员权限**。

### 代码风格
- 项目规模很小（两个 Python 文件），没有使用代码格式化工具（如 Black、Ruff）或类型检查。
- 注释和 UI 文本均为中文。
- 类方法和函数使用中文 docstring 或行内注释描述职责。
- Tkinter 变量和布局代码比较密集，修改 UI 时注意保持 `pack`/`grid` 布局的一致性。

### 图像引用与内存管理
- `create_shortcut_button()` 中创建的 `ImageTk.PhotoImage` 必须绑定到按钮属性（`btn.image = icon_image`），否则会被 Python 垃圾回收导致图标消失。

### 单实例机制
- `app.lock` 存放当前进程 PID。若程序异常退出未清理锁文件，下次启动时会通过 `psutil.pid_exists()` 自动判断并处理。
- 不要手动删除或修改 `app.lock`，除非确定程序已完全退出。

### 数据持久化
- 所有用户数据（分组和快捷方式）存储在 `shortcuts.json` 中。
- 删除该文件会清空所有数据，程序下次启动时会重新生成空的 `{"groups": {}}`。
- 快捷方式路径存储为绝对路径，跨设备迁移时路径会失效。

### 图标资源
- `icons/` 目录下的三个 PNG 文件是运行时可选项：
  - 若缺失，程序会使用纯色占位图（白色或灰色）。
  - 不要删除 `icon.png`，否则托盘图标会变成一个白色方块。

---

## 测试策略

**本项目目前没有任何自动化测试。**

由于这是一个以 GUI 交互为主的桌面应用，且代码高度耦合在 `MiniQDesktop` 类中，目前也没有为单元测试做解耦设计。如果需要验证修改：

1. **手动运行测试**：直接运行 `python main.py` 或 `run.bat`，在真实环境中操作界面。
2. **关键验证点**：
   - 添加快捷方式（文件 + 文件夹）到现有分组
   - 添加快捷方式到新分组
   - 删除快捷方式和分组
   - 折叠/展开分组（互斥行为）
   - 全局热键 `Alt+Z`
   - 系统托盘菜单（显示 / 退出）
   - 单实例启动（重复启动应提示并退出）
   - 重启后数据是否从 `shortcuts.json` 正确恢复

---

## 部署与发布

本项目没有自动化构建或打包流程。目前的使用方式就是直接克隆/复制源代码并运行。

如果需要打包为独立可执行文件（如 `.exe`），可以考虑使用 **PyInstaller** 或 **nuitka**，但项目中尚未配置。若添加打包配置，建议：
- 将 `icons/`、`shortcuts.json` 作为外部资源或打包为内置资源
- 处理好 `pythonw` 无控制台场景下的异常日志（已写入 `error.log`）

---

## 安全注意事项

1. **管理员权限**：`keyboard` 全局热键可能需要管理员权限，但在普通权限下有时也会工作（视系统安全策略而定）。
2. **锁文件**：`app.lock` 仅包含 PID，不包含敏感信息。
3. **数据文件**：`shortcuts.json` 存储的是用户本地的文件路径，不包含密码或密钥。
4. **外部程序启动**：程序通过 `os.startfile(path)` 启动用户指定的任意文件/程序，这等同于用户双击运行，风险由用户自行承担。
5. **错误日志**：`error.log` 以追加模式写入，可能随时间增长，长期运行建议定期清理。

---

## 常见修改场景指南

| 场景 | 应该修改的文件 | 说明 |
|------|---------------|------|
| 修改界面布局/样式 | `MiniQDesktop.py` | 集中在 `setup_ui()` 和 `create_group_frame()` |
| 修改窗口默认位置/大小 | `MiniQDesktop.py` | `position_window()` 方法 |
| 修改全局热键 | `MiniQDesktop.py` | `setup_global_hotkey()` 中的 `'alt+z'` |
| 修改托盘菜单 | `MiniQDesktop.py` | `setup_tray_icon()` |
| 添加新依赖 | `requirements.txt` | 同时更新本文档的"技术栈"表格 |
| 修改启动方式 | `run.bat` / `main.py` | `run.bat` 仅一行命令 |
| 修改数据格式 | `MiniQDesktop.py` | `load_shortcuts()` / `save_shortcuts()` |

---

## 版本与授权

- 许可证：MIT License
- 代码生成方式：绝大部分由 DeepSeek AI 生成，人工修复运行时错误和优化交互习惯
