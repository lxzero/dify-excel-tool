import sys
import os
import importlib.util
import json

# 动态导入 tools/writeExcel.py
spec = importlib.util.spec_from_file_location("write_excel_module", os.path.join("tools", "writeExcel.py"))
write_excel_module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(write_excel_module)

# 构造 mock runtime 和 session
class Dummy:
    def __getattr__(self, name):
        return lambda *args, **kwargs: None
runtime = Dummy()
session = Dummy()

# 读取 test_example.json 文件
with open("test_example.json", "r", encoding="utf-8") as f:
    json_data = json.load(f)

# 实例化工具
excel_tool_instance = write_excel_module.WriteExcelTool(runtime, session)

# 调用 _invoke 方法，debug=True
params = {
    "json_str": json.dumps(json_data),
    "filename": "test_debug_blob",
    "debug": True
}

for msg in excel_tool_instance._invoke(params):
    if hasattr(msg, 'blob') and msg.blob:
        print("已生成 blob 消息（内容略）")
    elif hasattr(msg, 'content'):
        print("消息内容:", msg.content)
    else:
        print("消息类型:", type(msg)) 