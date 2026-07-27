# 办公工具箱

实用的办公自动化小工具集合。

## 目录结构

```
office-toolbox/
├── 图片插入表格/          # 项目1：Excel 图片批量插入工具
│   ├── excel_image_inserter.py
│   └── requirements.txt
├── 办公工具箱/            # 项目2：办公工具箱 HTML 工具页
│   └── 办公工具箱.html
└── .github/workflows/    # Windows EXE 自动打包工作流
```

## 项目1：Excel 图片批量插入工具（图片插入表格/）

根据 Excel 指定列的单元格内容匹配图片文件名，将图片批量嵌入单元格或插入图片超链接。

**功能特性**

- 图形界面（GUI），无需命令行操作
- 支持精确匹配与模糊匹配（可调阈值）
- 支持嵌入图片到单元格，或插入图片超链接
- 自动标注匹配类型（精确/模糊/未匹配）及匹配到的图片名
- 未匹配数据自动汇总到独立 sheet，便于人工复核
- 支持递归扫描图片文件夹（含子文件夹）

**使用方法**

```bash
pip install -r 图片插入表格/requirements.txt
python 图片插入表格/excel_image_inserter.py
```

**Windows 可执行文件**

在 [Releases](../../releases) 页面下载打包好的 `excel_image_inserter.exe`，双击即可运行，无需安装 Python 环境。

## 项目2：办公工具箱 HTML 工具页（办公工具箱/）

「压坊 · PDF 编辑与办公工具」单文件 HTML 工具页，无需打包，浏览器直接打开 `办公工具箱.html` 即可使用。
