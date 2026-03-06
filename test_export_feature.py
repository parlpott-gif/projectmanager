#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
导出功能测试脚本
验证导出记录管理函数是否正常工作
"""

import os
import sys
import json
from datetime import datetime

# 添加当前目录到 Python 路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from server import (
    load_export_records,
    save_export_records,
    add_export_record,
    EXPORT_RECORDS_PATH
)


def test_export_functions():
    """测试导出记录函数"""
    print("\n" + "="*60)
    print("导出功能测试")
    print("="*60 + "\n")

    # 测试 1: 加载空记录
    print("[TEST 1] 加载导出记录...")
    records = load_export_records()
    assert isinstance(records, dict), "返回值应该是字典"
    assert "records" in records, "应该包含 'records' 键"
    print(f"[PASS] 当前记录数: {len(records.get('records', []))}\n")

    # 测试 2: 添加记录
    print("[TEST 2] 添加导出记录...")
    add_export_record(
        "LED控制器",
        ["LED控制器+PCB制板文件.zip", "LED控制器+原理图.pdf"],
        "硬件素材"
    )
    records = load_export_records()
    assert len(records["records"]) > 0, "应该至少有一条记录"
    first_record = records["records"][0]
    assert first_record["project_name"] == "LED控制器", "项目名应该是 'LED控制器'"
    assert first_record["file_count"] == 2, "文件数应该是 2"
    assert first_record["sub_folder"] == "硬件素材", "文件夹应该是 '硬件素材'"
    print(f"[PASS] 成功添加记录，ID: {first_record['id']}\n")

    # 测试 3: 验证时间戳
    print("[TEST 3] 验证时间戳格式...")
    assert "id" in first_record, "记录应该包含 'id' 字段"
    assert "export_time" in first_record, "记录应该包含 'export_time' 字段"
    assert "timestamp" in first_record, "记录应该包含 'timestamp' 字段"
    print(f"[PASS] 导出时间: {first_record['export_time']}\n")

    # 测试 4: 添加多条记录
    print("[TEST 4] 添加多条记录...")
    add_export_record(
        "温湿度传感器",
        ["温湿度传感器_展示_01.jpg", "温湿度传感器+功能演示.mp4"],
        "演示"
    )
    records = load_export_records()
    assert len(records["records"]) >= 2, "应该至少有两条记录"
    print(f"[PASS] 当前记录数: {len(records['records'])}\n")

    # 测试 5: 验证记录顺序（最新的在前）
    print("[TEST 5] 验证记录顺序...")
    assert records["records"][0]["project_name"] == "温湿度传感器", "最新的记录应该在前面"
    assert records["records"][1]["project_name"] == "LED控制器", "旧记录应该在后面"
    print("[PASS] 记录顺序正确（最新的在前面）\n")

    # 测试 6: 验证导出记录文件
    print("[TEST 6] 验证数据文件...")
    assert os.path.exists(EXPORT_RECORDS_PATH), "导出记录文件应该存在"
    with open(EXPORT_RECORDS_PATH, 'r', encoding='utf-8') as f:
        data = json.load(f)
    assert isinstance(data, dict), "文件内容应该是字典"
    assert "records" in data, "文件应该包含 'records' 键"
    print(f"[PASS] 文件路径: {EXPORT_RECORDS_PATH}\n")

    # 测试 7: 手动保存和加载
    print("[TEST 7] 测试保存和加载...")
    test_data = {
        "records": [
            {
                "id": "test_20260302_120000",
                "project_name": "测试项目",
                "export_time": "2026-03-02 12:00:00",
                "timestamp": 1772000000,
                "files": ["test1.pdf", "test2.zip"],
                "file_count": 2,
                "sub_folder": "硬件素材"
            }
        ]
    }
    save_export_records(test_data)
    loaded = load_export_records()
    assert len(loaded["records"]) > 0, "应该加载到数据"
    print("[PASS] 保存和加载成功\n")

    # 清理测试数据
    print("[CLEANUP] 清理测试数据...")
    if os.path.exists(EXPORT_RECORDS_PATH):
        os.remove(EXPORT_RECORDS_PATH)
        print("[OK] 测试数据已清理\n")

    # 总结
    print("="*60)
    print("[OK] 所有导出功能测试通过！")
    print("="*60 + "\n")

    return True


if __name__ == '__main__':
    try:
        success = test_export_functions()
        sys.exit(0 if success else 1)
    except Exception as e:
        print(f"\n[ERROR] 测试失败: {e}\n")
        import traceback
        traceback.print_exc()
        sys.exit(1)
