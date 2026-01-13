
# -*- coding: utf-8 -*-
"""
新工作流测试：
Word → XML → 定位字段 → LLM填充 → 回填XML → Word
"""

import json
import traceback
from pathlib import Path

from requests import request
import requests
from lxml import etree

from llm_connector import LLMConnector, load_config_from_file
from process_with_llm import XMLPageAnalyzer
from merge_pages import merge_pages
from xml_to_docx import xml_to_docx

# 提供的数据
PROVIDED_DATA = {
    "项目名称": "测试项目",
    "包编号": "PKG-001",
    "投标人": "上海星辰科技有限公司",
    "时间": "2026年1月13日",
    "联系人": "李明",
    "电话": "021-56789012",
    "传真": "021-56789013",
    "法定代表人": "王强",
    "技术负责人": "赵丽"
}

def extract_field_context(xml_file, max_context_len=200):
    """
    Step 1: XML解析 + Step 2: 定位待填字段
    提取所有待填字段及其上下文
    """
    analyzer = XMLPageAnalyzer(xml_file)
    
    # 获取所有待填字段
    blank_fields = analyzer.find_blank_fields()
    placeholder_fields = analyzer.find_placeholder_fields()
    
    fields_info = []
    
    # 处理占位符字段
    for placeholder in placeholder_fields:
        fields_info.append({
            'id': placeholder,
            'type': 'placeholder',
            'context': f'字段: {placeholder}',
            'description': f'需要填充的占位符: {placeholder}'
        })
    
    # 处理空白字段
    for blank in blank_fields:
        fields_info.append({
            'id': blank,
            'type': 'blank',
            'context': f'空白字段: {blank}',
            'description': f'需要填充的空白字段: {blank}'
        })
    
    return fields_info, analyzer

def build_llm_prompt(page_num, fields_info, provided_data):
    """
    Step 3: 只把字段描述+少量上下文交给大模型
    构建LLM提示词
    """
    fields_description = []
    for i, field in enumerate(fields_info, 1):
        fields_description.append(
            f"{i}. [{field['id']}] {field['description']}"
        )
    
    provided_data_str = json.dumps(provided_data, ensure_ascii=False, indent=2)
    
    prompt = f"""你是一个专业的文件填充助手。
    
需要填充的字段列表（第{page_num}页）：
{chr(10).join(fields_description)}

可用的数据：
{provided_data_str}

任务：
1. 分析每个字段的含义
2. 从提供的数据中选择最合适的值来填充这些字段
3. 返回JSON格式的填充结果

返回格式必须是有效的JSON，格式如下：
{{
    "field_id_1": "填充值1",
    "field_id_2": "填充值2"
}}

注意：
- 只返回JSON，不要有其他文字
- 如果某个字段在数据中找不到对应值，返回空字符串
- 确保返回的是有效的JSON格式
"""
    
    return prompt

def parse_llm_response(response_text):
    """
    Step 4: 大模型返回JSON填充值
    解析LLM返回的JSON
    """
    try:
        # 尝试提取JSON部分
        if '{' in response_text and '}' in response_text:
            start = response_text.find('{')
            end = response_text.rfind('}') + 1
            json_str = response_text[start:end]
            return json.loads(json_str)
        else:
            return {}
    except json.JSONDecodeError as e:
        print(f"JSON解析失败: {e}")
        return {}

def apply_fills_to_xml(analyzer, fill_values):
    """
    Step 5: 程序回填XML
    将填充值应用到XML中
    """
    tree = analyzer.tree
    root = tree.getroot()
    ns = analyzer.ns
    
    applied_count = 0
    
    # 回填占位符
    for text_elem in root.findall('.//ns0:t', ns):
        if text_elem.text:
            for field_id, fill_value in fill_values.items():
                placeholder = f"${{{field_id}}}"
                if placeholder in text_elem.text:
                    text_elem.text = text_elem.text.replace(placeholder, fill_value)
                    applied_count += 1
    
    return applied_count

def test():
    xml_content = ""
    with open('split_pages/page_8.xml', 'r', encoding='utf-8') as f:
        xml_content = f.read()
    messages = []
    system_prompt =  """
    你是一个精通 Word XML（docx document.xml）的文档结构分析专家。

    你的任务不是修改 XML 内容，而是：

    1. 识别文档中所有“需要人工填写的字段”
    2. 判断该字段在 XML 中的精确位置
    3. 返回结构化 JSON 数据，供程序后续自动回填

    你必须严格遵守以下规则：
    - ❌ 不修改 XML
    - ❌ 不生成示例数据
    - ❌ 不输出解释性文字
    - ✅ 只输出 JSON 数组
    - ✅ 使用 XPath 精确定位
    """

    prompt = """
    下面是一段 Word 的 document.xml 内容（可能是段落或表格的一部分）。

    请你完成以下任务：

    一、识别所有“需要填写的字段”，判断依据包括但不限于：
    - 字段后面为空
    - 出现“：”但没有内容
    - 出现明显的业务字段名称（如：投标人名称、联系人、电话等）

    二、对每一个需要填写的字段，返回以下信息：
    - field_id：英文标识，使用 snake_case
    - label：字段中文名称（原文）
    - xpath：字段值所在节点的 XPath（用于填写内容）
    - hint：一句简短填写提示（如：填写公司全称）

    三、返回格式要求：
    - 返回一个 JSON 数组
    - 不要输出任何解释说明
    - 不要包裹 markdown
    - XPath 必须是相对于 document.xml 的绝对或稳定路径

    示例输出格式（仅示例，内容请以实际 XML 为准）：

    [
    {
        "field_id": "bidder_name",
        "label": "投标人名称",
        "xpath": "/w:document/w:body/w:tbl[1]/w:tr[2]/w:tc[2]/w:p/w:r/w:t",
        "hint": "填写投标人公司全称"
    }
    ]

    下面是需要分析的 XML 内容：
    ========================
    {{xml_content}}
    ========================
    """
    print(system_prompt)
    prompt = prompt.replace("{{xml_content}}", xml_content)
    print(prompt)
    # system_prompt 官方示例没有 system，但是你可以自己加成 user 消息或者 role=system
    messages.append({"Role": "system", "Content": system_prompt})
    messages.append({"Role": "user", "Content": prompt})

    payload = {
        "MaxTokens": 4096,
        "Temperature": 0.6,
        "Model": "deepseek-v3-0324",
        "Messages": messages,
        "Stream": False   # 流式输出设 False，方便直接返回
    }

    response = requests.post(
        "https://api.lkeap.cloud.tencent.com/v1/chat/completions",
        json=payload,
        headers={
            "Authorization": "Bearer sk-VcqUMssTr1tzmkyEZ64N5Xba8k5jVCVm0ZYVMo2AqJKxvFIC",
            "Content-Type": "application/json",
            "X-TC-Action": "ChatCompletions"   # 官方要求
        },
        timeout=120
    )
    response.raise_for_status()
    result = response.json()
    return json.dumps(result['choices'][0]['message']['content'].replace('```json ', '').replace('```', '').strip(), ensure_ascii=False, indent=2)

def main():
    print("="*70)
    print("  新工作流测试: Word → XML → 字段定位 → LLM填充 → 回填XML → Word")
    print("="*70)
    
    try:
        # Step 0: 加载LLM配置
        print("\n[Step 0] 加载LLM配置")
        config = load_config_from_file('llm_config.json')
        print(f"✓ LLM提供商: {config.provider.value}")
        print(f"✓ 模型: {config.model}")
        
        connector = LLMConnector(config)
        
        # ===== 处理 Page 2 =====
        print("\n" + "="*70)
        print("处理 Page 2 (商务部分)")
        print("="*70)
        
        # Step 1 + 2: XML解析 + 定位字段
        print("\n[Step 1-2] XML解析 + 定位待填字段")
        xml_file_page2 = 'split_pages/page_2.xml'
        fields_info_page2, analyzer_page2 = extract_field_context(xml_file_page2)
        
        print(f"✓ 检测到 {len(fields_info_page2)} 个待填字段")
        for field in fields_info_page2[:5]:
            print(f"  - {field['id']} ({field['type']})")
        if len(fields_info_page2) > 5:
            print(f"  ... 等等 {len(fields_info_page2) - 5} 个字段")
        
        # Step 3: 构建LLM提示词
        print("\n[Step 3] 构建字段描述 + 上下文")
        prompt_page2 = build_llm_prompt(2, fields_info_page2, PROVIDED_DATA)
        print(f"✓ 提示词长度: {len(prompt_page2)} 字符")
        print("提示词内容预览:")
        print(prompt_page2[:300] + "...\n")
        
        # Step 4: 调用LLM获取填充值
        print("[Step 4] 调用LLM获取填充值JSON")
        print("正在请求LLM... (可能需要10-30秒)")
        
        response_page2 = connector.call(
            prompt=prompt_page2,
            system_prompt="你是专业的文件填充助手。只返回有效的JSON格式，不要有其他文字。"
        )
        
        print("✓ 收到LLM响应")
        print("LLM返回内容:")
        print(response_page2[:500])
        
        # 解析JSON
        fill_values_page2 = parse_llm_response(response_page2)
        print(f"\n✓ 解析到 {len(fill_values_page2)} 个填充值")
        for key, value in list(fill_values_page2.items())[:3]:
            print(f"  - {key}: {value}")
        if len(fill_values_page2) > 3:
            print(f"  ... 等等 {len(fill_values_page2) - 3} 个")
        
        # Step 5: 回填XML
        print("\n[Step 5] 回填XML")
        applied_count_page2 = apply_fills_to_xml(analyzer_page2, fill_values_page2)
        print(f"✓ 成功填充 {applied_count_page2} 个位置")
        
        # 保存修改后的Page 2
        analyzer_page2.save()
        print("✓ Page 2 已保存")
        
        # ===== 处理 Page 8 =====
        print("\n" + "="*70)
        print("处理 Page 8 (企业基本情况表)")
        print("="*70)
        
        # Step 1 + 2: XML解析 + 定位字段
        print("\n[Step 1-2] XML解析 + 定位待填字段")
        xml_file_page8 = 'split_pages/page_8.xml'
        fields_info_page8, analyzer_page8 = extract_field_context(xml_file_page8)
        
        print(f"✓ 检测到 {len(fields_info_page8)} 个待填字段")
        for field in fields_info_page8[:5]:
            print(f"  - {field['id']} ({field['type']})")
        if len(fields_info_page8) > 5:
            print(f"  ... 等等 {len(fields_info_page8) - 5} 个字段")
        
        # Step 3: 构建LLM提示词
        print("\n[Step 3] 构建字段描述 + 上下文")
        prompt_page8 = build_llm_prompt(8, fields_info_page8, PROVIDED_DATA)
        print(f"✓ 提示词长度: {len(prompt_page8)} 字符")
        
        # Step 4: 调用LLM获取填充值
        print("\n[Step 4] 调用LLM获取填充值JSON")
        print("正在请求LLM... (可能需要10-30秒)")
        
        response_page8 = connector.call(
            prompt=prompt_page8,
            system_prompt="你是专业的文件填充助手。只返回有效的JSON格式，不要有其他文字。"
        )
        
        print("✓ 收到LLM响应")
        
        # 解析JSON
        fill_values_page8 = parse_llm_response(response_page8)
        print(f"✓ 解析到 {len(fill_values_page8)} 个填充值")
        
        # Step 5: 回填XML
        print("\n[Step 5] 回填XML")
        applied_count_page8 = apply_fills_to_xml(analyzer_page8, fill_values_page8)
        print(f"✓ 成功填充 {applied_count_page8} 个位置")
        
        # 保存修改后的Page 8
        analyzer_page8.save()
        print("✓ Page 8 已保存")
        
        # Step 6: 合并XML
        print("\n" + "="*70)
        print("Step 6: 合并XML")
        print("="*70)
        
        import shutil
        import os
        
        test_split_dir = 'test_split_pages_llm'
        os.makedirs(test_split_dir, exist_ok=True)
        
        shutil.copy('split_pages/page_2.xml', f'{test_split_dir}/page_1.xml')
        shutil.copy('split_pages/page_8.xml', f'{test_split_dir}/page_2.xml')
        
        print("✓ 已复制修改后的页面文件")
        
        output_xml = 'test_merged_llm.xml'
        merge_pages(test_split_dir, output_xml)
        print(f"✓ 已合并为: {output_xml}")
        
        # Step 7: 转为Word
        print("\n" + "="*70)
        print("Step 7: 转换为Word")
        print("="*70)
        
        output_docx = 'test_output_llm_workflow.docx'
        xml_to_docx(output_xml, output_docx, 'template.docx')
        print(f"✓ 已转换为Word: {output_docx}")
        
        # 清理临时文件
        shutil.rmtree(test_split_dir)
        
        print("\n" + "="*70)
        print("✓ 流程完成！")
        print("="*70)
        print(f"\n最终输出文件: {output_docx}")
        print("\n工作流总结:")
        print(f"  Page 2: 检测{len(fields_info_page2)}字段 → 填充{applied_count_page2}位置")
        print(f"  Page 8: 检测{len(fields_info_page8)}字段 → 填充{applied_count_page8}位置")
        print(f"  总计: {len(fields_info_page2) + len(fields_info_page8)}字段 → {applied_count_page2 + applied_count_page8}位置")
        
    except Exception as e:
        print(f"\n❌ 处理失败: {e}")
        traceback.print_exc()

if __name__ == '__main__':
    # result = test()
    # json_result = json.loads(result)
    # print("Extracted JSON:")
    # print(json_result)
    xml_content = ""
    with open('split_pages/page_8.xml', 'r', encoding='utf-8') as f:
        xml_content = f.read()

    print(xml_content)