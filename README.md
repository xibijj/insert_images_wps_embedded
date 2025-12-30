# WPS Excel内嵌图片实现（DISPIMG指令）

## 项目概述

该脚本实现了在WPS Excel中使用DISPIMG指令将图片真正内嵌到指定单元格的功能。与传统的漂浮图片不同，使用此方法嵌入的图片会与单元格绑定，无法单独移动，并会随着单元格大小调整而自动缩放。

## 功能特点

- ✅ 支持将多张图片依次内嵌到指定单元格（默认B1, C1, D1...）
- ✅ 图片真正内嵌于单元格，与单元格绑定
- ✅ 图片随单元格大小调整而自动缩放
- ✅ 支持常见图片格式（PNG, JPG, JPEG, BMP, GIF）
- ✅ 自动处理图片ID生成和关系管理
- ✅ 兼容WPS Office格式要求

## 技术实现路线

### 1. 基础Excel文件创建
使用XlsxWriter创建基础Excel文件，设置列宽、行高，并在目标单元格中写入占位公式。

### 2. Excel文件结构解析
将Excel文件解压为XML格式，分析并修改以下关键文件：
- `[Content_Types].xml`：注册cellimages.xml内容类型
- `workbook.xml.rels`：添加cellimages.xml关系
- `cellimages.xml`：定义图片数据和位置信息
- `cellimages.xml.rels`：建立图片文件与cellimages.xml的关系
- `sheet1.xml`：更新单元格公式和属性

### 3. 图片处理与嵌入
- 复制图片到Excel的media目录
- 生成唯一图片ID和关系ID
- 在cellimages.xml中创建图片数据结构
- 更新工作表XML中的公式引用

### 4. 重新压缩Excel文件
将修改后的XML文件重新压缩为标准Excel格式。

## 核心实现细节

### 1. 图片ID和关系管理
```python
# 生成唯一图片ID
image_id = f'ID_{uuid.uuid4().hex}'

# 生成唯一关系ID
r_ids = []
# 从现有关系中提取最大ID
new_r_id = f'rId{max(r_ids) + 1}' if r_ids else 'rId1'
```

### 2. Cellimages.xml结构生成
使用正确的命名空间和嵌套结构创建cellimages.xml：
```python
root = etree.Element(
    '{http://www.wps.cn/officeDocument/2017/etCustomData}cellImages',
    nsmap={
        'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'etc': 'http://www.wps.cn/officeDocument/2017/etCustomData'
    }
)
```

### 3. 图片尺寸转换
将像素尺寸转换为Excel内部单位EMU（1像素=9525 EMU）：
```python
# 将像素转换为EMU（Excel内部单位，1像素=9525 EMU）
width_emu = width * 9525
height_emu = height * 9525
```

### 4. 单元格属性设置
为包含DISPIMG公式的单元格设置正确的属性：
```python
# 设置单元格类型为str
b>cell.set('t', 'str')</b>

# 更新公式和值
f_elem.text = f'_xlfn.DISPIMG("{image_id}",1)'
v_elem.text = f'=DISPIMG("{image_id}",1)'
```

## 能力边界

### 支持的环境
- ✅ WPS Office 2019及以上版本
- ✅ Python 3.7+ 环境
- ✅ Windows, Linux, macOS 操作系统

### 限制
- ❌ 不支持Microsoft Excel（DISPIMG是WPS特有指令）
- ❌ 不支持动态调整图片位置（固定内嵌于单元格）
- ❌ 不支持图片旋转等高级编辑功能
- ❌ 处理大量图片时可能影响性能

## 二次开发扩展

### 1. 支持更多单元格位置
修改`update_worksheet`函数中的单元格地址计算逻辑：
```python
# 当前实现（横向）
column_letter = chr(ord('B') + column_index)
cell_address = f'{column_letter}1'

# 修改为纵向
row_index = 1 + column_index
cell_address = f'B{row_index}'
```

### 2. 自定义图片尺寸
修改`update_cellimages`函数中的尺寸转换逻辑：
```python
# 当前实现
width_emu = width * 9525
height_emu = height * 9525

# 修改为自定义尺寸
max_width_emu = 5000000  # 最大宽度
max_height_emu = 3000000  # 最大高度
width_emu = min(width * 9525, max_width_emu)
height_emu = min(height * 9525, max_height_emu)
```

### 3. 支持批量处理不同目录
修改`IMG_DIR`常量为可配置参数：
```python
# 当前实现
IMG_DIR = './img'

# 修改为命令行参数
import argparse
parser = argparse.ArgumentParser()
parser.add_argument('--img-dir', default='./img', help='图片目录路径')
args = parser.parse_args()
IMG_DIR = args.img_dir
```

## 使用方法

### 1. 准备图片
将需要嵌入的图片放入`./img`目录下。

### 2. 安装依赖
```bash
pip install -r requirements.txt
```

### 3. 运行脚本
```bash
python insert_images_wps_embedded.py
```

### 4. 查看结果
脚本将生成`images_wps_embedded.xlsx`文件，使用WPS Office打开即可查看内嵌图片效果。

## 依赖项

```
xlwt==1.3.0
Pillow==9.0.1
lxml==4.9.1
XlsxWriter==3.0.3
```

## 项目结构

```
├── insert_images_wps_embedded.py  # 主脚本
├── examine_excel.py              # Excel文件分析工具
├── README.md                     # 项目文档
└── img/                          # 图片目录
```

## 故障排除

### 图片显示#REF!错误
1. 确保使用WPS Office打开文件
2. 检查图片格式是否受支持
3. 确保脚本正确执行，没有报错
4. 使用`examine_excel.py`工具分析文件结构

### 图片无法显示
1. 检查cellimages.xml是否正确生成
2. 验证图片文件是否存在于media目录
3. 确认关系文件配置正确

## 示例

运行脚本后，将在Excel中生成如下效果：

| A  | B     | C     | D     | ...
|----|-------|-------|-------|----
| 1  | 图片1 | 图片2 | 图片3 | ...

所有图片都将真正内嵌于对应单元格，无法单独移动，并会随单元格大小调整而缩放。

## 效果展示

以下是脚本运行后生成的Excel文件效果预览：

![内嵌图片效果展示](./img/result.png)

## 注意事项

1. 该脚本使用WPS特有指令，不支持Microsoft Excel
2. 处理大图片时可能需要调整内存设置
3. 建议定期备份原始图片文件
4. 如需批量处理大量图片，建议分批执行脚本

## 许可证

MIT License
