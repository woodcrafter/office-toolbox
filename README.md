# 办公工具箱

实用的办公自动化小工具集合。

## 目录结构

```
office-toolbox/
├── 图片插入表格/          # 项目1：Excel 图片批量插入工具
│   ├── excel_image_inserter.py      # 悬浮图片 / 超链接版
│   ├── excel_cell_image_embedder.py # 真嵌入单元格版（WPS DISPIMG / Excel 365）
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

在 [Releases](../../releases) 页面下载打包好的 exe，双击即可运行，无需安装 Python 环境：

- `excel_image_inserter.exe` — 悬浮图片 / 超链接版
- `excel_cell_image_embedder.exe` — 真嵌入单元格版（WPS / Excel 365）

## 项目1b：Excel 图片真嵌入单元格工具（excel_cell_image_embedder.py）

与上面的悬浮图片版不同，本工具生成的是**真正嵌入单元格**的图片：图片随单元格移动、筛选、排序，同工作簿内可用公式（如 `=Sheet1!D3`、`VLOOKUP`）跨表引用图片。

**两种嵌入格式（GUI 中可选）**

- **WPS（推荐）**：WPS 原生 DISPIMG 格式，用 WPS 打开查看；在 Microsoft Excel 中无法显示图片
- **Excel 365**：微软 richData「放置在单元格中」格式，需要 Microsoft 365 打开；部分 WPS 版本会显示图片破损

**功能特性**

- 图形界面（GUI），无需命令行操作
- 精确/模糊匹配（可调阈值），自动调整行高列宽到指定单元格尺寸
- 完整保留原工作簿的 sheet、样式与合并单元格（WPS 模式直接改原簿）
- 自动标注匹配类型与匹配图片名，未匹配数据汇总到独立 sheet
- webp 等格式自动转 PNG 后嵌入

**使用方法**

```bash
pip install -r 图片插入表格/requirements.txt
python 图片插入表格/excel_cell_image_embedder.py
```

## 项目2：办公工具箱 HTML 工具页（办公工具箱/）

「压坊 · PDF 编辑与办公工具」单文件 HTML 工具页，无需打包，浏览器直接打开 `办公工具箱.html` 即可使用。
