"""
文件自动重命名规则测试脚本

这个脚本测试 get_auto_renamed_file() 和 get_numbered_image_name() 函数
的正确性，无需启动完整的 Flask 服务器。
"""

import os
import sys
import tempfile

# 导入 server 模块中的函数
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from server import get_auto_renamed_file, get_numbered_image_name


def print_test_result(test_name, expected, actual, passed):
    """打印测试结果"""
    status = "[PASS]" if passed else "[FAIL]"
    print(f"{status} {test_name}")
    if not passed:
        print(f"     Expected: {expected}")
        print(f"     Actual:   {actual}")
    print()


def test_hardware_materials_stage():
    """测试硬件素材阶段的重命名规则"""
    print("\n" + "="*60)
    print("Test: Hardware Materials Stage (Draft, Waiting for PCB)")
    print("="*60 + "\n")

    test_cases = [
        # (filename, project_name, expected)
        ("pcb_design.zip", "LED控制器", "LED控制器+PCB制板文件.zip"),
        ("DESIGN.ZIP", "温湿度传感器", "温湿度传感器+PCB制板文件.ZIP"),
        ("project.epro2", "电源管理模块", "电源管理模块+eda工程.epro2"),
        ("CIRCUIT.EPRO2", "通信模块", "通信模块+eda工程.EPRO2"),
        ("schematic.pdf", "马达控制器", "马达控制器+原理图.pdf"),
        ("diagram.PDF", "显示屏驱动", "显示屏驱动+原理图.PDF"),
        ("photo.jpg", "LED控制器", "LED控制器+图片展示.jpg"),
        ("image.PNG", "传感器模块", "传感器模块+图片展示.PNG"),
        ("circuit.png", "电路板", "电路板+图片展示.png"),
        ("design.gif", "UI设计", "UI设计+图片展示.gif"),
        ("demo.bmp", "演示器", "演示器+图片展示.bmp"),
        # 应该保持原名的文件
        ("bom_list.xlsx", "LED控制器", "bom_list.xlsx"),
        ("notes.doc", "传感器", "notes.doc"),
        ("readme.txt", "项目", "readme.txt"),
    ]

    passed = 0
    for filename, project_name, expected in test_cases:
        result = get_auto_renamed_file(filename, project_name, '硬件素材')
        is_passed = result == expected
        if is_passed:
            passed += 1
        print_test_result(f"{filename} → {expected}", expected, result, is_passed)

    print(f"\nHardware Materials: {passed}/{len(test_cases)} tests passed\n")
    return passed == len(test_cases)


def test_demo_stage_non_image():
    """测试演示阶段非图片文件的重命名规则"""
    print("\n" + "="*60)
    print("Test: Demo Stage - Non-Image Files")
    print("="*60 + "\n")

    test_cases = [
        # (filename, project_name, expected)
        ("demo.mp4", "LED控制器", "LED控制器+功能演示.mp4"),
        ("VIDEO.MOV", "传感器", "传感器+功能演示.MOV"),
        ("test.avi", "电路板", "电路板+功能演示.avi"),
        ("record.mkv", "显示屏", "显示屏+功能演示.mkv"),
        # 应该保持原名的文件
        ("README.md", "项目", "README.md"),
        ("document.pdf", "说明", "document.pdf"),
        ("notes.txt", "记录", "notes.txt"),
    ]

    passed = 0
    for filename, project_name, expected in test_cases:
        result = get_auto_renamed_file(filename, project_name, '演示')
        is_passed = result == expected
        if is_passed:
            passed += 1
        print_test_result(f"{filename} → {expected}", expected, result, is_passed)

    print(f"\nDemo Stage (Non-Image): {passed}/{len(test_cases)} tests passed\n")
    return passed == len(test_cases)


def test_demo_stage_images():
    """测试演示阶段图片的自动编号"""
    print("\n" + "="*60)
    print("Test: Demo Stage - Image Auto-Numbering")
    print("="*60 + "\n")

    # 创建临时目录模拟项目文件夹
    with tempfile.TemporaryDirectory() as temp_dir:
        demo_dir = os.path.join(temp_dir, "演示")
        os.makedirs(demo_dir, exist_ok=True)

        project_name = "LED控制器"

        # 测试连续上传多张图片，验证编号递增
        # 重要：必须先创建占位符文件，以便下一个编号正确递增
        test_files = [
            ("photo1.jpg", "LED控制器_展示_01.jpg"),
            ("photo2.png", "LED控制器_展示_02.png"),
            ("photo3.gif", "LED控制器_展示_03.gif"),
            ("photo4.bmp", "LED控制器_展示_04.bmp"),
        ]

        passed = 0
        for i, (filename, expected) in enumerate(test_files):
            # 创建前面文件的占位符（在调用 get_numbered_image_name 之前）
            for j in range(i):
                placeholder_file = test_files[j][1]
                placeholder_path = os.path.join(demo_dir, placeholder_file)
                if not os.path.exists(placeholder_path):
                    open(placeholder_path, 'w').close()

            # 获取重命名建议（演示图片返回元组）
            result = get_auto_renamed_file(filename, project_name, '演示')

            # 如果是图片，应该返回元组
            if isinstance(result, tuple):
                # 调用 get_numbered_image_name 生成最终名称
                final_name = get_numbered_image_name(result[1], demo_dir, result[2])
            else:
                final_name = result

            is_passed = final_name == expected
            if is_passed:
                passed += 1

            print_test_result(f"{filename} -> {expected}", expected, final_name, is_passed)

        print(f"\nDemo Stage (Image Numbering): {passed}/{len(test_files)} tests passed\n")
        return passed == len(test_files)


def test_edge_cases():
    """测试边界情况"""
    print("\n" + "="*60)
    print("Test: Edge Cases")
    print("="*60 + "\n")

    # Test case-insensitive handling
    result = get_auto_renamed_file("MyFile.ZIP", "TestProject", "硬件素材")
    expected = "TestProject+PCB制板文件.ZIP"
    is_passed = result == expected
    print_test_result("Case-insensitive ZIP", expected, result, is_passed)

    # Test file without extension
    result = get_auto_renamed_file("README", "Project", "硬件素材")
    expected = "README"
    is_passed = result == expected
    print_test_result("File without extension", expected, result, is_passed)

    # Test multiple dots in filename
    result = get_auto_renamed_file("my.design.v1.zip", "Project", "硬件素材")
    expected = "Project+PCB制板文件.zip"
    is_passed = result == expected
    print_test_result("Multiple dots in filename", expected, result, is_passed)

    return True


def run_all_tests():
    """运行所有测试"""
    print("\n" + "="*60)
    print("File Auto-Rename Rules Test Suite")
    print("="*60)

    results = []
    results.append(test_hardware_materials_stage())
    results.append(test_demo_stage_non_image())
    results.append(test_demo_stage_images())
    results.append(test_edge_cases())

    # Summary
    print("\n" + "="*60)
    print("Test Summary")
    print("="*60)
    all_passed = all(results)
    if all_passed:
        print("[OK] All tests passed!")
        return 0
    else:
        print("[FAILED] Some tests failed. Please check implementation.")
        return 1


if __name__ == '__main__':
    sys.exit(run_all_tests())
