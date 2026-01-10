# Cheatsheet Maker

作者复习电电时突然冒出的想法，基于antigravity全程vibe coding做的cheatsheet生成器，目前支持读取pdf和ppt文件来快捷生成a4大小的cheatsheet，经测试把一页a4分成四列刚刚好。推荐搭配雨课堂爬虫食用。

赞美伟大的gemini。

## 环境要求

- Windows 操作系统
- Python 3.x
- Microsoft PowerPoint (仅 PPT 转换功能需要)
- 依赖库: `customtkinter`, `PyMuPDF` (fitz), `Pillow`, `pywin32`

## 安装说明

1. 安装项目依赖:

   ```bash
   pip install -r requirements.txt
   ```

2. 准备素材:
   将你的 PDF 或 PPT 课件/资料放入 `res` 文件夹中 (支持子文件夹分类)。

## 使用指南

1. 启动程序:

   ```bash
   python main.py
   ```

2. **选择文件**:
   在左侧侧边栏的文件树中找到你的文件并点击。

3. **查看与添加**:
   - 文件内容会加载到中间的主视图区域。
   - 使用鼠标滚轮浏览，按住 Ctrl 滚动可缩放。
   - **点击** 任意页面图片，它就会自动加入右侧的小抄预览中。

4. **管理排版**:
   - 如果遇到长图放不下的情况，根据弹窗提示选择是否分割。
   - 右侧面板实时展示排版结果。
   - 使用 **Undo (撤销)** 按钮移除刚刚添加的内容。
   - 使用 **Clear (清空)** 按钮清空当前画板。

5. **导出成果**:
   - 满意后，点击 **Export PDF**。
   - 选择保存路径，即可获得最终的 PDF 文件。

## 快捷键

- **Ctrl + 滚轮**: 缩放视图 (同时适用于中间阅读器和右侧预览区)。
- **鼠标左键单击**: 将当前页面添加到小抄。

## 注意事项

- 程序会自动将转换后的 PPT 缓存为 PDF 文件保存在 `res/.cache` 目录中，下次打开同一文件时无需等待转换。
