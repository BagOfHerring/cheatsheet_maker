# Cheatsheet Maker

作者复习电电时突然冒出的想法，基于antigravity全程vibe coding做的cheatsheet生成器，目前支持读取pdf和ppt文件来快捷生成a4大小的cheatsheet，经测试把一页a4分成四列刚刚好。推荐搭配雨课堂爬虫食用。

赞美伟大的gemini。

## UI 参考与开源声明

当前四栏面板布局和控制面板层级参考了开源项目 [Tabler](https://github.com/tabler/tabler) 的 dashboard/admin panel 设计思路。Tabler 使用 [MIT License](https://github.com/tabler/tabler/blob/dev/LICENSE)。本项目没有复制 Tabler 的源码或素材，仅参考其开源软件面板的信息层级、分栏和控制区组织方式，并使用 Python Tk / CustomTkinter 自行实现。

## 环境要求

- Windows / macOS / Linux 均可运行 PDF 相关功能
- Python 3.x
- Microsoft PowerPoint (仅 Windows 下的 PPT 转 PDF 功能需要)
- 依赖库: `customtkinter`, `PyMuPDF` (fitz), `Pillow`
- Windows PPT 转换依赖: `pywin32`, `comtypes`

## 部署/安装说明

推荐使用项目内 `.venv` 虚拟环境部署，避免依赖安装到全局 Python 环境中。

1. 进入项目目录:

   ```powershell
   cd path\to\cheatsheet_maker
   ```

   macOS 如果使用 Homebrew Python，先确认 Tk 支持可用:

   ```bash
   python -c "import tkinter; print(tkinter.TkVersion)"
   ```

   如果提示 `ModuleNotFoundError: No module named '_tkinter'`，需要安装与 Python 版本匹配的 Tk 包，例如 Python 3.14:

   ```bash
   brew install python-tk@3.14
   ```

2. 创建并激活 `.venv`:

   macOS / Linux / Git Bash:

   ```bash
   python -m venv .venv
   source .venv/bin/activate
   ```

   Windows PowerShell:

   ```powershell
   python -m venv .venv
   .\.venv\Scripts\Activate.ps1
   ```

   如果使用 `cmd`，激活命令改为:

   ```bat
   .\.venv\Scripts\activate.bat
   ```

   如果 PowerShell 提示禁止运行脚本，可在当前终端执行:

   ```powershell
   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
   .\.venv\Scripts\Activate.ps1
   ```

3. 安装项目依赖:

   ```powershell
   python -m pip install --upgrade pip
   python -m pip install -r requirements.txt
   ```

4. 准备素材:
   将你的 PDF 或 PPT 课件/资料放入 `res` 文件夹中 (支持子文件夹分类)。

5. 启动程序:

   ```powershell
   python main.py
   ```

后续再次使用时，只需要重新激活 `.venv` 后启动程序:

macOS / Linux / Git Bash:

```bash
source .venv/bin/activate
python main.py
```

Windows PowerShell:

```powershell
.\.venv\Scripts\Activate.ps1
python main.py
```

## 使用指南

1. **启动程序**:

   ```bash
   python main.py
   ```

2. **选择文件**:
   在左侧侧边栏的文件树中找到你的文件并点击。

3. **查看与添加**:
   - 文件内容会加载到中间的阅读区。
   - 中间阅读区连续显示整份课件，使用鼠标滚轮上下浏览。
   - 使用 **Prev / Next** 或页码输入框跳转到指定页面。
   - 按住 Ctrl 滚动可缩放阅读区，点击 **Fit** 可恢复适宽显示。
   - 点击阅读区中的任意页面，或点击 **Add Page**，即可加入右侧的小抄列表。

4. **管理排版**:
   - 界面分为四栏: 文件夹、课件阅读区、小抄预览区、控制面板。
   - 小抄预览区实时展示 A4 四栏排版预览。
   - 拖动左右分隔条可调整四栏宽度。
   - 在小抄预览区中点击已加入的页面块可以选中该项。
   - 选中后可在 **Item Properties** 中调整缩放比例与矩形裁切范围。
   - 使用 **Remove / Up / Down** 调整已加入页面。
   - 使用 **Undo** 移除刚刚添加的内容。
   - 使用 **Clear** 清空当前画板。

5. **导出成果**:
   - 满意后，点击 **Export PDF**。
   - 选择保存路径，即可获得最终的 PDF 文件。
   - 导出时会按每个页面块的缩放与裁切设置直接嵌入源 PDF 页面，尽量保留原 PDF 中的文字与矢量内容。
   - 导出文件同时，在同一路径下获得白板 PDF 文件，用于后续再操作。

6. **后续调整**：
   - 现有的PDF留白较多，空间利用率低，字体小。
   - 可以使用[**极截速抠**](https://github.com/xiaozy24/ClipMatte/tree/master)将 PDF 内容调整字体大小与排版转移至白板 PDF 中。

## 快捷键

- **Ctrl + 滚轮**: 缩放视图 (同时适用于中间阅读器和小抄预览区)。
- **鼠标左键单击**: 将当前页面添加到小抄。

## 注意事项

- 程序会自动将转换后的 PPT 缓存为 PDF 文件保存在 `res/.cache` 目录中，下次打开同一文件时无需等待转换。
