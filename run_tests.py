#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
测试运行脚本
支持运行 unittest 和 pytest 测试
"""

import sys
import os
import subprocess
import argparse

def run_unittest_tests():
    """运行 unittest 测试"""
    print("=" * 60)
    print("运行 unittest 测试")
    print("=" * 60)
    
    # 添加项目根目录到Python路径
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    
    # 导入并运行测试
    import unittest
    from tests.test_write_excel import TestWriteExcelTool, TestWriteExcelToolIntegration
    
    # 创建测试套件
    test_suite = unittest.TestSuite()
    
    # 添加单元测试
    test_suite.addTest(unittest.makeSuite(TestWriteExcelTool))
    test_suite.addTest(unittest.makeSuite(TestWriteExcelToolIntegration))
    
    # 运行测试
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(test_suite)
    
    return result.wasSuccessful()

def run_pytest_tests():
    """运行 pytest 测试"""
    print("=" * 60)
    print("运行 pytest 测试")
    print("=" * 60)
    
    try:
        # 运行 pytest
        result = subprocess.run([
            sys.executable, "-m", "pytest", 
            "tests/", 
            "-v", 
            "--tb=short"
        ], capture_output=True, text=True)
        
        print(result.stdout)
        if result.stderr:
            print("错误输出:")
            print(result.stderr)
        
        return result.returncode == 0
    except Exception as e:
        print(f"运行 pytest 时出错: {e}")
        return False

def run_specific_test(test_name):
    """运行特定的测试"""
    print(f"=" * 60)
    print(f"运行特定测试: {test_name}")
    print("=" * 60)
    
    try:
        result = subprocess.run([
            sys.executable, "-m", "pytest", 
            f"tests/test_write_excel_pytest.py::{test_name}", 
            "-v", 
            "--tb=short"
        ], capture_output=True, text=True)
        
        print(result.stdout)
        if result.stderr:
            print("错误输出:")
            print(result.stderr)
        
        return result.returncode == 0
    except Exception as e:
        print(f"运行特定测试时出错: {e}")
        return False

def run_coverage_tests():
    """运行覆盖率测试"""
    print("=" * 60)
    print("运行覆盖率测试")
    print("=" * 60)
    
    try:
        # 运行覆盖率测试
        result = subprocess.run([
            sys.executable, "-m", "pytest", 
            "tests/", 
            "--cov=tools", 
            "--cov-report=html", 
            "--cov-report=term-missing",
            "-v"
        ], capture_output=True, text=True)
        
        print(result.stdout)
        if result.stderr:
            print("错误输出:")
            print(result.stderr)
        
        return result.returncode == 0
    except Exception as e:
        print(f"运行覆盖率测试时出错: {e}")
        return False

def main():
    """主函数"""
    parser = argparse.ArgumentParser(description="运行 Excel 工具测试")
    parser.add_argument(
        "--test-type", 
        choices=["unittest", "pytest", "all", "coverage"], 
        default="all",
        help="选择测试类型 (默认: all)"
    )
    parser.add_argument(
        "--test-name",
        help="运行特定的测试方法名 (仅适用于 pytest)"
    )
    
    args = parser.parse_args()
    
    success = True
    
    if args.test_name:
        success = run_specific_test(args.test_name)
    elif args.test_type == "unittest":
        success = run_unittest_tests()
    elif args.test_type == "pytest":
        success = run_pytest_tests()
    elif args.test_type == "coverage":
        success = run_coverage_tests()
    elif args.test_type == "all":
        print("运行所有测试...")
        success_unittest = run_unittest_tests()
        print("\n")
        success_pytest = run_pytest_tests()
        success = success_unittest and success_pytest
    
    print("=" * 60)
    if success:
        print("✅ 所有测试通过!")
        sys.exit(0)
    else:
        print("❌ 部分测试失败!")
        sys.exit(1)

if __name__ == "__main__":
    main() 