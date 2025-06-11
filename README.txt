
Word Chapter Deletion Tool / Word章节删除工具
===========================================

This is a Python-based GUI application that allows users to visualize and delete specific chapters (including sub-chapters) from a Microsoft Word `.docx` file based on heading numbering.

这是一个基于 Python 的图形界面应用程序，允许用户根据 Word 文档中的“编号标题”结构，可视化显示章节结构并删除所选章节及其所有子章节。

Features | 功能
---------

- 🧠 Automatically detects numbered headings from a Word document (e.g. "1", "1.2", "1.2.3")
- 👁️ TreeView interface to show the document's heading hierarchy
- 🗑️ Delete selected chapter and its sub-chapters with one click
- ⚠️ Supports visual numbering detection (COM interface, accurate with Word multi-level lists)

- 自动识别 Word 文档中的“编号标题”（如 "1", "1.2", "1.2.3"）
- 使用 TreeView 显示文档章节结构
- 一键删除所选章节及其子章节
- 支持 Word 的视觉编号识别（通过 COM 接口，更准确）

Dependencies | 依赖项
--------------------

- Python 3.7+
- `python-docx`
- `pywin32` (for Word COM automation)
- `tkinter` (GUI, built-in with Python)

安装依赖：

    pip install python-docx pywin32

How to Run | 使用方法
---------------------

1. Place your target `word.docx` file in the same directory.
2. Run the script:

       python EditWordChapter.py

3. In the GUI, you'll see the numbered heading structure.
4. Select any chapter, then click "删除选中章节" (Delete selected chapter).

1. 将你需要操作的 `word.docx` 文件放到程序所在目录；
2. 运行程序：

       python EditWordChapter.py

3. 在图形界面中查看自动生成的章节树；
4. 选中任意章节后点击“删除选中章节”即可。

Notes | 注意事项
----------------

- The program only deletes paragraphs with visual numbering, so normal text won't be affected.
- Only `.docx` files (Word 2007+) are supported.
- Word must be installed on the system (due to COM interface usage).

- 本程序只删除具有“视觉编号”的段落，普通正文内容不会被误删；
- 仅支持 `.docx` 格式（Word 2007及以上）；
- 需要系统已安装 Microsoft Word（使用 COM 接口）；

File Structure | 文件结构
-------------------------

    ├── EditWordChapter.py     # Main Python script
    ├── word.docx              # Word file to be processed
    ├── README.txt             # This file

License | 许可协议
------------------

This project is released under the MIT License.

本项目采用 MIT 协议开源。
