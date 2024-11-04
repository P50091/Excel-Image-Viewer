# Excel-Image-Viewer

通过读取 Excel 表格中的 URL 链接，预览表格中的图片，并提取其他关键字段。

## 项目描述

本项目旨在开发一个 Python 脚本，通过读取 Excel 表格中的 URL 链接，预览表格中的图片，并提取其他关键字段。该脚本将使用 `pandas` 库读取 Excel 文件，利用 `requests` 库获取图片，并通过 `IPython.display` 库进行图片预览。该功能特别适用于需要快速浏览和处理包含图片链接的 Excel 数据的场景，如电商产品管理、数据分析和报告生成等。

## 功能概述

1. **读取 Excel 表格**：通过 `pandas` 库读取 Excel 文件，提取包含图片 URL 和其他关键字段的数据。
2. **获取图片**：使用 `requests` 库从 URL 链接中获取图片数据。
3. **预览图片**：利用 `IPython.display` 库在 Jupyter Notebook 或其他支持的环境中预览图片。
4. **提取和显示关键字段**：从 Excel 表格中提取并显示其他关键字段的信息。

## 实现步骤

### 1. 安装所需库

确保安装了以下 Python 库：

```bash
pip install pandas openpyxl requests IPython
```

### 2. 读取 Excel 表格

使用 `pandas` 读取 Excel 文件，并提取包含 URL 和其他关键字段的数据。

```python
import pandas as pd

# 读取 Excel 文件
df = pd.read_excel('your_excel_file.xlsx')

# 假设 URL 列名为 'ImageURL'，其他关键字段为 'Field1' 和 'Field2'
urls = df['ImageURL']
field1 = df['Field1']
field2 = df['Field2']
```

### 3. 获取图片

使用 `requests` 库从 URL 中获取图片数据。

```python
import requests
from IPython.display import Image, display

def get_image_from_url(url):
    response = requests.get(url)
    if response.status_code == 200:
        return Image(response.content)
    else:
        return None
```

### 4. 预览图片和显示关键字段

在 Jupyter Notebook 中预览图片并显示其他关键字段的信息。

```python
for i, url in enumerate(urls):
    img = get_image_from_url(url)
    if img:
        display(img)
    print(f"Field1: {field1[i]}, Field2: {field2[i]}")
```

## 应用场景

- **电商产品管理**：快速预览产品图片并查看相关信息。
- **数据分析**：在数据分析过程中，结合图片和其他字段进行综合分析。
- **报告生成**：生成包含图片和关键字段的报告，便于展示和分享。

## 示例

### Excel 表格

![Excel 表格](\data\excel1.png)

### 界面 1

![界面 1](\data\view2.png)

### 界面 2

![界面 2](\data\view3.png)

## 结论

通过该 Python 脚本，用户可以方便地读取和处理 Excel 表格中的图片 URL 链接，预览图片，并提取和显示其他关键字段信息。这将大大提高数据处理和分析的效率，特别是在需要处理大量图片数据的场景中。