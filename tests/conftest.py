#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pytest
import sys
import os

# 添加项目根目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# pytest配置
def pytest_configure(config):
    """pytest配置"""
    config.addinivalue_line(
        "markers", "unit: 标记为单元测试"
    )
    config.addinivalue_line(
        "markers", "integration: 标记为集成测试"
    )
    config.addinivalue_line(
        "markers", "slow: 标记为慢速测试"
    )

@pytest.fixture
def mock_runtime():
    """模拟运行时对象"""
    from unittest.mock import Mock
    return Mock()

@pytest.fixture
def mock_session():
    """模拟会话对象"""
    from unittest.mock import Mock
    return Mock()

@pytest.fixture
def simple_data():
    """简单测试数据"""
    return [
        {"姓名": "张三", "年龄": 25, "部门": "技术部"},
        {"姓名": "李四", "年龄": 30, "部门": "市场部"}
    ]

@pytest.fixture
def enhanced_data(simple_data):
    """增强格式测试数据"""
    return {
        "data": simple_data,
        "format": {
            "column_widths": {"A": 15, "B": 10, "C": 20},
            "row_heights": {"1": 25, "2": 20},
            "merge_cells": ["A1:C1"],
            "cells": {
                "1,1": {
                    "font": {"bold": True, "size": 14},
                    "background_color": "FFFF00"
                },
                "2,2": {
                    "font": {"italic": True},
                    "alignment": {"horizontal": "center"}
                }
            }
        }
    } 