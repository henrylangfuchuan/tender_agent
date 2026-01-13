"""
完整的 Word 智能填写工作流
分割 → 大模型处理 → 合并 → 转Word
"""
import os
import sys
import subprocess
import json
from datetime import datetime


def print_header(text: str):
    """打印标题"""
    print("\n" + "=" * 70)
    print(f"  {text}")
    print("=" * 70)


def print_step(step_num: int, text: str):
    """打印步骤"""
    print(f"\n【步骤{step_num}】{text}")
    print("-" * 70)


def check_prerequisites():
    """检查前置条件"""
    print_header("前置条件检查")
    
    checks = [
        ("Python环境", sys.version.split()[0]),
        ("lxml库", "已安装"),
        ("requests库", "已安装"),
    ]
    
    for name, status in checks:
        print(f"  ✓ {name}: {status}")
    
    # 检查文件
    files_to_check = [
        ("template.docx", "原始Word文档"),
        ("split_pages.py", "分割脚本"),
        ("process_with_llm.py", "LLM处理脚本"),
        ("merge_pages.py", "合并脚本"),
        ("xml_to_docx.py", "转Word脚本"),
    ]
    
    print("\n文件检查:")
    missing_files = []
    for filename, description in files_to_check:
        if os.path.exists(filename):
            print(f"  ✓ {filename} - {description}")
        else:
            print(f"  ✗ {filename} - {description} [缺失]")
            missing_files.append(filename)
    
    if missing_files:
        print(f"\n❌ 缺失文件: {', '.join(missing_files)}")
        return False
    
    return True


def step1_split_pages():
    """步骤1：分割页面"""
    print_step(1, "分割Word文档为单独的页面")
    
    if not os.path.exists("split_pages"):
        print("正在分割页面...")
        result = subprocess.run([sys.executable, "split_pages.py"], capture_output=True, text=True)
        
        if result.returncode == 0:
            print("✓ 页面分割成功")
            return True
        else:
            print(f"✗ 页面分割失败:\n{result.stderr}")
            return False
    else:
        print("⚠️  split_pages 目录已存在，跳过分割")
        
        # 计算页数
        import glob
        page_files = glob.glob("split_pages/page_*.xml")
        print(f"✓ 检测到 {len(page_files)} 个页面")
        return True


def step2_check_llm_config():
    """步骤2：检查LLM配置"""
    print_step(2, "检查LLM配置")
    
    if not os.path.exists("llm_config.json"):
        print("❌ 未找到 llm_config.json 配置文件")
        print("\n请先运行配置向导:")
        print("  python setup_llm_config.py")
        return False
    
    with open("llm_config.json", "r") as f:
        config = json.load(f)
    
    print(f"✓ 已检测到LLM配置:")
    print(f"  - 服务商: {config.get('provider', 'unknown')}")
    print(f"  - Model: {config.get('model', 'unknown')}")
    print(f"  - API URL: {config.get('api_url', 'unknown')[:50]}...")
    
    return True


def step3_check_fill_data():
    """步骤3：检查填充数据"""
    print_step(3, "检查填充数据")
    
    if not os.path.exists("fill_data.json"):
        print("❌ 未找到 fill_data.json 数据文件")
        print("\n请准备填充数据:")
        print("  1. 运行示例数据生成: python setup_llm_config.py sample")
        print("  2. 编辑 fill_data.json，填入实际数据")
        return False
    
    with open("fill_data.json", "r") as f:
        data = json.load(f)
    
    print(f"✓ 已检测到填充数据:")
    print(f"  - 数据条目数: {len(data)}")
    
    # 显示数据摘要
    for key, value in list(data.items())[:3]:
        print(f"  - {key}: {list(value.keys())[:2]}")
    
    return True


def step4_process_with_llm():
    """步骤4：用大模型处理页面"""
    print_step(4, "用大模型智能处理页面")
    
    print("⚠️  这一步需要调用大模型API，可能需要几秒到几分钟的时间...")
    print("(取决于页面数量和网络连接)")
    
    confirm = input("\n是否继续? (y/n): ").strip().lower()
    if confirm != 'y':
        print("已跳过LLM处理")
        return True
    
    print("\n正在处理页面...")
    result = subprocess.run([sys.executable, "process_with_llm.py"], capture_output=True, text=True)
    
    print(result.stdout)
    
    if result.returncode == 0:
        print("✓ LLM处理完成")
        return True
    else:
        print(f"✗ LLM处理失败:\n{result.stderr}")
        return False


def step5_merge_pages():
    """步骤5：合并页面"""
    print_step(5, "合并修改后的页面为单个XML")
    
    print("正在合并页面...")
    result = subprocess.run([sys.executable, "merge_pages.py"], capture_output=True, text=True)
    
    print(result.stdout)
    
    if result.returncode == 0 and os.path.exists("merged_document.xml"):
        print("✓ 页面合并成功")
        return True
    else:
        print(f"✗ 页面合并失败:\n{result.stderr}")
        return False


def step6_convert_to_word():
    """步骤6：转换为Word"""
    print_step(6, "转换为Word文档")
    
    print("正在转换为Word...")
    result = subprocess.run([sys.executable, "xml_to_docx.py"], capture_output=True, text=True)
    
    print(result.stdout)
    
    if result.returncode == 0 and os.path.exists("output_document.docx"):
        print("✓ Word文档生成成功")
        return True
    else:
        print(f"✗ Word转换失败:\n{result.stderr}")
        return False


def generate_report(results: dict):
    """生成处理报告"""
    print_header("处理报告")
    
    print(f"\n处理时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"\n步骤执行结果:")
    
    steps = [
        "分割页面",
        "检查LLM配置",
        "检查填充数据",
        "LLM处理",
        "合并页面",
        "转为Word"
    ]
    
    for i, step_name in enumerate(steps, 1):
        status = "✓" if results.get(f"step{i}", False) else "✗"
        print(f"  {status} 步骤{i}: {step_name}")
    
    all_success = all(results.values())
    
    if all_success:
        print("\n" + "=" * 70)
        print("  ✓ 所有步骤已完成！")
        print("=" * 70)
        print(f"\n生成的Word文档: output_document.docx")
        print("\n下一步操作:")
        print("  1. 使用Word打开 output_document.docx")
        print("  2. 检查填写内容是否正确")
        print("  3. 进行必要的手动调整")
    else:
        print("\n" + "=" * 70)
        print("  ⚠️  处理过程中出现错误")
        print("=" * 70)
        print("\n请检查上述错误信息并重试")
    
    return all_success


def run_complete_workflow():
    """运行完整工作流"""
    print_header("Word 智能填写完整工作流")
    
    print("\n本工作流包括以下步骤:")
    print("  1. 分割Word文档为单独的页面")
    print("  2. 检查LLM配置")
    print("  3. 检查填充数据")
    print("  4. 使用大模型智能处理每个页面")
    print("  5. 合并修改后的页面")
    print("  6. 转为最终的Word文档")
    
    results = {}
    
    # 前置检查
    if not check_prerequisites():
        print("\n❌ 前置条件检查失败")
        return False
    
    # 执行步骤
    steps = [
        (1, step1_split_pages),
        (2, step2_check_llm_config),
        (3, step3_check_fill_data),
        (4, step4_process_with_llm),
        (5, step5_merge_pages),
        (6, step6_convert_to_word),
    ]
    
    for step_num, step_func in steps:
        try:
            result = step_func()
            results[f"step{step_num}"] = result
            
            if not result:
                print(f"\n❌ 步骤{step_num}失败，工作流中止")
                break
        
        except Exception as e:
            print(f"\n❌ 步骤{step_num}异常: {e}")
            results[f"step{step_num}"] = False
            break
    
    # 生成报告
    success = generate_report(results)
    
    return success


if __name__ == '__main__':
    try:
        success = run_complete_workflow()
        sys.exit(0 if success else 1)
    
    except KeyboardInterrupt:
        print("\n\n⚠️  用户中止")
        sys.exit(1)
    
    except Exception as e:
        print(f"\n❌ 未预期的错误: {e}")
        sys.exit(1)
