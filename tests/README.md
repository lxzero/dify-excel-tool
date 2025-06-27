# 测试文档

## 概述

本项目包含完整的单元测试和集成测试，确保 Excel 工具的功能正确性和稳定性。

## 测试结构

```
tests/
├── __init__.py              # 测试包初始化文件
├── conftest.py              # pytest 配置文件
├── test_write_excel.py      # unittest 风格的测试
├── test_write_excel_pytest.py  # pytest 风格的测试
└── README.md                # 测试文档
```

## 测试类型

### 1. 单元测试 (Unit Tests)

测试各个独立方法的功能：

- **列索引标准化测试**

  - `test_normalize_cell_key_letter_format`: 测试字母格式的单元格键标准化
  - `test_normalize_cell_key_number_format`: 测试数字格式的单元格键标准化
  - `test_normalize_column_index_letter_format`: 测试字母格式的列索引标准化
  - `test_normalize_column_index_number_format`: 测试数字格式的列索引标准化

- **格式应用测试**

  - `test_apply_cell_format_with_letter_index`: 测试使用字母索引应用单元格格式
  - `test_apply_cell_format_with_number_index`: 测试使用数字索引应用单元格格式
  - `test_apply_column_width_with_letter_index`: 测试使用字母索引应用列宽设置
  - `test_apply_column_width_with_number_index`: 测试使用数字索引应用列宽设置
  - `test_apply_row_height`: 测试应用行高设置

- **合并单元格测试**
  - `test_apply_merge_cells_string_format`: 测试字符串格式的合并单元格
  - `test_apply_merge_cells_dict_format`: 测试字典格式的合并单元格
  - `test_apply_merge_cells_list_format`: 测试列表格式的合并单元格

### 2. 集成测试 (Integration Tests)

测试完整的工作流程：

- `test_invoke_simple_format`: 测试简单格式的 JSON 数据处理
- `test_invoke_enhanced_format`: 测试增强格式的 JSON 数据处理
- `test_full_workflow_with_letter_indexes`: 测试使用字母索引的完整工作流程
- `test_full_workflow_with_number_indexes`: 测试使用数字索引的完整工作流程
- `test_full_workflow_with_mixed_indexes`: 测试混合索引的完整工作流程

### 3. 边界情况测试

- `test_invoke_invalid_json`: 测试无效 JSON 数据的处理
- `test_invoke_dict_data`: 测试字典数据的处理
- `test_invoke_empty_data`: 测试空数据的处理
- `test_filename_processing`: 测试文件名处理
- `test_default_filename`: 测试默认文件名

## 运行测试

### 1. 使用测试运行脚本

```bash
# 运行所有测试
python run_tests.py

# 只运行 unittest 测试
python run_tests.py --test-type unittest

# 只运行 pytest 测试
python run_tests.py --test-type pytest

# 运行覆盖率测试
python run_tests.py --test-type coverage

# 运行特定测试
python run_tests.py --test-name TestWriteExcelToolPytest::test_normalize_cell_key_letter_format
```

### 2. 直接使用 pytest

```bash
# 运行所有测试
pytest

# 运行特定测试文件
pytest tests/test_write_excel_pytest.py

# 运行特定测试类
pytest tests/test_write_excel_pytest.py::TestWriteExcelToolPytest

# 运行特定测试方法
pytest tests/test_write_excel_pytest.py::TestWriteExcelToolPytest::test_normalize_cell_key_letter_format

# 运行标记的测试
pytest -m unit          # 只运行单元测试
pytest -m integration   # 只运行集成测试

# 运行覆盖率测试
pytest --cov=tools --cov-report=html --cov-report=term-missing
```

### 3. 直接使用 unittest

```bash
# 运行所有 unittest 测试
python -m unittest tests.test_write_excel

# 运行特定测试类
python -m unittest tests.test_write_excel.TestWriteExcelTool

# 运行特定测试方法
python -m unittest tests.test_write_excel.TestWriteExcelTool.test_normalize_cell_key_letter_format
```

## 测试标记

- `@pytest.mark.unit`: 单元测试
- `@pytest.mark.integration`: 集成测试
- `@pytest.mark.slow`: 慢速测试

## 测试覆盖率

运行覆盖率测试会生成 HTML 报告，可以在 `htmlcov/` 目录中查看详细的覆盖率信息。

## 测试数据

测试使用以下类型的测试数据：

1. **简单数据**: 基本的数组格式数据
2. **增强数据**: 包含格式配置的完整数据结构
3. **边界数据**: 空数据、无效数据等边界情况

## 模拟对象

测试中使用了以下模拟对象：

- `mock_runtime`: 模拟运行时对象
- `mock_session`: 模拟会话对象
- `mock_create_text_message`: 模拟文本消息创建
- `mock_create_blob_message`: 模拟二进制消息创建

## 注意事项

1. 测试需要安装所有依赖包
2. 某些测试需要模拟 Dify 插件的运行时环境
3. 覆盖率测试需要安装 `pytest-cov` 包
4. 测试文件中的路径设置确保可以正确导入模块

## 持续集成

建议在 CI/CD 流程中包含以下测试步骤：

1. 安装依赖: `pip install -r requirements.txt`
2. 运行单元测试: `pytest -m unit`
3. 运行集成测试: `pytest -m integration`
4. 生成覆盖率报告: `pytest --cov=tools --cov-report=xml`
