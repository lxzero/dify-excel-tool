## excel-tool

**Author:** lxzero
**Version:** 0.0.1
**Type:** tool

### Description

Excel 工具插件，提供强大的 JSON 转 Excel 功能，支持丰富的格式化选项。

### 功能特性

#### 基础功能

- JSON 数据转 Excel 文件
- 支持数组和对象格式的数据
- 自动生成表头

#### 高级格式化功能

- **单元格宽高设置**：自定义列宽和行高
- **边框样式**：支持多种边框样式（细线、粗线、虚线等）
- **背景颜色**：设置单元格背景颜色
- **文字颜色**：自定义字体颜色
- **字体设置**：字体名称、大小、加粗、斜体
- **对齐方式**：水平对齐、垂直对齐、自动换行
- **合并单元格**：支持多种格式的单元格合并

### 使用方法

#### 简单格式

直接传入数据数组或对象：

```json
[
  { "姓名": "张三", "年龄": 25, "部门": "技术部" },
  { "姓名": "李四", "年龄": 30, "部门": "市场部" }
]
```

#### 增强格式

包含格式配置的完整结构：

```json
{
  "data": [
    { "姓名": "张三", "年龄": 25, "部门": "技术部" },
    { "姓名": "李四", "年龄": 30, "部门": "市场部" }
  ],
  "format": {
    "column_widths": { "A": 15, "B": 10, "C": 20 },
    "row_heights": { "1": 25 },
    "merge_cells": ["A1:C1", { "start": "A2", "end": "B2" }],
    "cells": {
      "1,1": {
        "font": {
          "name": "微软雅黑",
          "size": 14,
          "bold": true,
          "color": "FFFFFF"
        },
        "background_color": "366092",
        "border": {
          "left": "thick",
          "right": "thick",
          "top": "thick",
          "bottom": "thick"
        },
        "alignment": { "horizontal": "center", "vertical": "center" }
      }
    }
  }
}
```

### 详细文档

更多格式配置说明和示例，请参考 [Excel 格式化配置指南](EXCEL_FORMAT_GUIDE.md)。

### 依赖项

- `dify_plugin>=0.2.0,<0.3.0`
- `openpyxl>=3.0.0`
- `pandas>=1.3.0`

### 安装

```bash
pip install -r requirements.txt
```

### 更新日志

#### v0.0.1

- 新增高级格式化功能
- 支持单元格宽高、边框、颜色、字体等设置
- 新增合并单元格功能
- 向后兼容简单格式
- 添加详细的配置文档
