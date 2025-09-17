# Excel 文件批量下载脚本

本项目用于从 Excel 文件中读取 K~T 列的下载链接，批量下载文件（支持图片、docx、pdf、zip 等），并按照指定规则命名和分类存储。适用于跨平台环境（Mac/Windows/Linux），直接使用 Python requests 下载文件。  

脚本功能：

- 遍历 Excel 中 K~T 列的所有链接
- 下载所有类型文件（图片/docx/pdf/zip 等）
- 文件命名格式为 `H列_I列 + 文件后缀`
- 自动分类到 **序号前缀 01~10 的 10 个文件夹**
- 自动处理重复文件名 `(1)`、`(2)`，防止覆盖
- 支持重定向和部分防盗链请求

---

## 环境配置

1. 安装 Python 3.x（推荐 >=3.9）。
2. 建议使用虚拟环境：
```bash
python3 -m venv venv
source venv/bin/activate   # Mac/Linux
venv\Scripts\activate      # Windows

## 文件命名

1.总文件夹：cc_information
2.excel文档：cc_collection.xlsx

## 扩展安装
pip3 install -r requirements.txt

## 使用方法
将文件夹拖入vscode后，安装扩展即可运行./cc_information/new_process.py 使用