# Excel 格式化 JSON 配置说明

## 概述

增强版的 JSON 转 Excel 工具支持两种格式：

1. **简单格式**：直接传入数据数组或对象
2. **增强格式**：包含数据和格式配置的完整结构

## 列索引格式支持

工具支持两种列索引格式：

### 字母格式 (A, B, C, ...)

- 使用 Excel 标准的字母列标识
- 支持单字母 (A-Z) 和多字母 (AA, AB, ...)
- 例如：`"A"`, `"B"`, `"C"`, `"AA"`, `"AB"`

### 数字格式 (1, 2, 3, ...)

- 使用数字列索引
- 从 1 开始计数
- 例如：`1`, `2`, `3`, `26`, `27`

### 混合使用

在同一个配置中可以混合使用两种格式：

```json
{
  "format": {
    "cells": {
      "1,A": { "font": { "bold": true } }, // 字母格式
      "1,1": { "font": { "bold": true } }, // 数字格式
      "2,B": { "background_color": "FFFF00" }, // 字母格式
      "2,2": { "background_color": "FFFF00" } // 数字格式
    },
    "column_widths": {
      "A": 15, // 字母格式
      "1": 15, // 数字格式
      "B": 10, // 字母格式
      "2": 10 // 数字格式
    }
  }
}
```

## 简单格式

直接传入数据，不包含格式设置：

```json
[
  { "姓名": "张三", "年龄": 25, "部门": "技术部" },
  { "姓名": "李四", "年龄": 30, "部门": "市场部" }
]
```

## 增强格式

包含数据和格式配置的完整结构：

```json
{
  "data": [
    { "姓名": "张三", "年龄": 25, "部门": "技术部" },
    { "姓名": "李四", "年龄": 30, "部门": "市场部" }
  ],
  "format": {
    "column_widths": {
      "A": 15,
      "B": 10,
      "C": 20
    },
    "row_heights": {
      "1": 25,
      "2": 20,
      "3": 20
    },
    "merge_cells": ["A1:C1", { "start": "A2", "end": "B2" }, ["A3", "C3"]],
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
        "alignment": {
          "horizontal": "center",
          "vertical": "center"
        }
      },
      "2,1": {
        "font": {
          "size": 12,
          "color": "000000"
        },
        "background_color": "E6E6E6"
      }
    }
  }
}
```

## 格式配置详解

### 1. 列宽设置 (column_widths)

设置指定列的宽度，支持字母和数字两种格式：

```json
"column_widths": {
  "A": 15,  // A列宽度为15 (字母格式)
  "1": 15,  // 第1列宽度为15 (数字格式)
  "B": 10,  // B列宽度为10 (字母格式)
  "2": 10,  // 第2列宽度为10 (数字格式)
  "C": 20   // C列宽度为20 (字母格式)
}
```

### 2. 行高设置 (row_heights)

设置指定行的高度：

```json
"row_heights": {
  "1": 25,  // 第1行高度为25
  "2": 20,  // 第2行高度为20
  "3": 20   // 第3行高度为20
}
```

### 3. 合并单元格 (merge_cells)

支持多种格式的合并单元格配置：

```json
"merge_cells": [
  "A1:C1",                    // 字符串格式：合并A1到C1
  { "start": "A2", "end": "B2" },  // 对象格式：合并A2到B2
  ["A3", "C3"]               // 数组格式：合并A3到C3
]
```

合并单元格支持三种格式：

1. **字符串格式**：`"A1:B2"` - 直接指定合并范围
2. **对象格式**：`{"start": "A1", "end": "B2"}` - 分别指定起始和结束单元格
3. **数组格式**：`["A1", "B2"]` - 数组形式指定起始和结束单元格

### 4. 单元格格式 (cells)

使用 `"行号,列索引"` 的格式指定单元格位置，支持字母和数字两种列索引格式：

```json
"cells": {
  "1,1": {  // 第1行第1列 (数字格式)
    "font": { ... },
    "background_color": "366092",
    "border": { ... },
    "alignment": { ... }
  },
  "1,A": {  // 第1行A列 (字母格式)
    "font": { ... },
    "background_color": "366092"
  },
  "2,3": {  // 第2行第3列 (数字格式)
    "font": { ... }
  },
  "2,C": {  // 第2行C列 (字母格式)
    "font": { ... }
  }
}
```

#### 字体设置 (font)

```json
"font": {
  "name": "微软雅黑",     // 字体名称
  "size": 14,           // 字体大小
  "bold": true,         // 是否加粗
  "italic": false,      // 是否斜体
  "color": "FFFFFF"     // 字体颜色（十六进制，不含#）
}
```

#### 背景颜色 (background_color)

```json
"background_color": "366092"  // 十六进制颜色值，不含#
```

#### 边框设置 (border)

```json
"border": {
  "left": "thin",      // 左边框样式
  "right": "thin",     // 右边框样式
  "top": "thin",       // 上边框样式
  "bottom": "thin"     // 下边框样式
}
```

边框样式选项：

- `"thin"` - 细线
- `"medium"` - 中等线
- `"thick"` - 粗线
- `"dashed"` - 虚线
- `"dotted"` - 点线
- `"double"` - 双线

#### 对齐方式 (alignment)

```json
"alignment": {
  "horizontal": "center",  // 水平对齐
  "vertical": "center",    // 垂直对齐
  "wrap_text": false       // 是否自动换行
}
```

对齐选项：

- 水平对齐：`"left"`, `"center"`, `"right"`, `"justify"`
- 垂直对齐：`"top"`, `"center"`, `"bottom"`, `"justify"`

## 完整示例

### 示例 1：带表头的格式化表格

```json
{
  "data": [
    { "姓名": "张三", "年龄": 25, "部门": "技术部", "薪资": 8000 },
    { "姓名": "李四", "年龄": 30, "部门": "市场部", "薪资": 9000 },
    { "姓名": "王五", "年龄": 28, "部门": "人事部", "薪资": 7500 }
  ],
  "format": {
    "column_widths": {
      "A": 12,
      "B": 8,
      "C": 15,
      "D": 12
    },
    "row_heights": {
      "1": 30
    },
    "merge_cells": ["A1:D1"],
    "cells": {
      "1,1": {
        "font": {
          "name": "微软雅黑",
          "size": 16,
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
        "alignment": {
          "horizontal": "center",
          "vertical": "center"
        }
      }
    }
  }
}
```

### 示例 2：复杂表格布局（带合并单元格）

```json
{
  "data": [
    { "部门": "技术部", "姓名": "张三", "年龄": 25, "薪资": 8000 },
    { "部门": "技术部", "姓名": "李四", "年龄": 28, "薪资": 8500 },
    { "部门": "市场部", "姓名": "王五", "年龄": 30, "薪资": 9000 },
    { "部门": "市场部", "姓名": "赵六", "年龄": 27, "薪资": 8800 }
  ],
  "format": {
    "column_widths": {
      "A": 15,
      "B": 12,
      "C": 8,
      "D": 12
    },
    "row_heights": {
      "1": 30,
      "2": 25,
      "3": 25,
      "4": 25,
      "5": 25
    },
    "merge_cells": ["A1:A2", "A3:A4", "A5:A5"],
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
        "alignment": {
          "horizontal": "center",
          "vertical": "center"
        }
      },
      "1,2": {
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
        "alignment": {
          "horizontal": "center",
          "vertical": "center"
        }
      },
      "1,3": {
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
        "alignment": {
          "horizontal": "center",
          "vertical": "center"
        }
      },
      "1,4": {
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
        "alignment": {
          "horizontal": "center",
          "vertical": "center"
        }
      },
      "2,1": {
        "font": {
          "name": "微软雅黑",
          "size": 12,
          "bold": true,
          "color": "000000"
        },
        "background_color": "E6E6E6",
        "border": {
          "left": "thin",
          "right": "thin",
          "top": "thin",
          "bottom": "thin"
        },
        "alignment": {
          "horizontal": "center",
          "vertical": "center"
        }
      },
      "3,1": {
        "font": {
          "name": "微软雅黑",
          "size": 12,
          "bold": true,
          "color": "000000"
        },
        "background_color": "E6E6E6",
        "border": {
          "left": "thin",
          "right": "thin",
          "top": "thin",
          "bottom": "thin"
        },
        "alignment": {
          "horizontal": "center",
          "vertical": "center"
        }
      }
    }
  }
}
```

### 示例 3：条件格式（突出显示高薪资）

```json
{
  "data": [
    { "姓名": "张三", "年龄": 25, "部门": "技术部", "薪资": 8000 },
    { "姓名": "李四", "年龄": 30, "部门": "市场部", "薪资": 9000 },
    { "姓名": "王五", "年龄": 28, "部门": "人事部", "薪资": 7500 }
  ],
  "format": {
    "merge_cells": ["A1:D1"],
    "cells": {
      "1,1": {
        "font": {
          "name": "微软雅黑",
          "size": 16,
          "bold": true,
          "color": "FFFFFF"
        },
        "background_color": "366092",
        "alignment": {
          "horizontal": "center",
          "vertical": "center"
        }
      },
      "2,4": {
        "background_color": "FFE6E6",
        "font": {
          "color": "CC0000",
          "bold": true
        }
      },
      "3,4": {
        "background_color": "FFE6E6",
        "font": {
          "color": "CC0000",
          "bold": true
        }
      }
    }
  }
}
```

## 注意事项

1. **颜色格式**：所有颜色值都使用十六进制格式，不包含 `#` 符号
2. **单元格位置**：使用 `"行号,列索引"` 的格式，从 1 开始计数
3. **列标识**：使用字母标识（A, B, C...）
4. **合并单元格**：支持多种格式，建议使用字符串格式 `"A1:B2"`
5. **兼容性**：简单格式仍然支持，向后兼容
6. **默认值**：未指定的格式将使用 Excel 默认样式
7. **合并顺序**：合并单元格操作在数据写入和格式应用之后进行

## 常用颜色代码

- 白色：`FFFFFF`
- 黑色：`000000`
- 红色：`FF0000`
- 绿色：`00FF00`
- 蓝色：`0000FF`
- 黄色：`FFFF00`
- 灰色：`808080`
- 深蓝：`366092`
- 浅灰：`E6E6E6`
