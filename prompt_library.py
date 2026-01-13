"""
提示词模板库 - 为不同场景提供完善的提示词
"""
from typing import Dict, List
from enum import Enum


class PromptTemplate:
    """提示词模板"""
    
    def __init__(self, name: str, system_prompt: str, user_prompt_template: str):
        self.name = name
        self.system_prompt = system_prompt
        self.user_prompt_template = user_prompt_template
    
    def format(self, **kwargs) -> tuple:
        """
        格式化提示词
        
        :param kwargs: 模板变量
        :return: (系统提示词, 用户提示词)
        """
        system = self.system_prompt.format(**kwargs)
        user = self.user_prompt_template.format(**kwargs)
        return system, user


class PromptLibrary:
    """提示词库"""
    
    # ============ 招标文件填写模板 ============
    
    TENDER_FORM_FILLING = PromptTemplate(
        name="招标文件表单填写",
        system_prompt="""你是一个专业的招标文件智能填写助手。你的任务是根据提供的数据，
智能填写招标文件中的空白字段。

你需要遵循以下原则：
1. 【准确性】所有填写内容必须与提供的数据完全匹配，不能猜测或编造
2. 【格式一致】保持原文档的格式和结构不变
3. 【智能理解】根据字段名和上下文，理解应该填写什么内容
4. 【完整性】如果某个字段有多个相关数据，确保全部填写
5. 【合规性】填写内容应符合招标规范和法律要求""",
        
        user_prompt_template="""请根据以下数据，填写Word文档中的相应字段。

【页面信息】
第{page_num}页 - {page_title}

【页面内容分析】
当前页面包含以下需要填写的字段（用[  ]表示）：
{fields_to_fill}

【提供的数据】
{provided_data}

【任务要求】
1. 分析当前页面中所有空白字段
2. 根据提供的数据，智能匹配和填写每个字段
3. 输出修改后的XML内容，保持原始格式不变
4. 对于无法确定的字段，保留原样

【输出格式】
以JSON格式返回：
{{
    "fields_filled": [
        {{"field_name": "字段名", "original_value": "原值", "new_value": "新值", "confidence": 0.95}},
        ...
    ],
    "unfilled_fields": ["字段名1", "字段名2"],
    "notes": "任何特殊说明",
    "xml_updates": [
        {{"xpath": "XML路径", "old_content": "旧内容", "new_content": "新内容"}},
        ...
    ]
}}"""
    )
    
    # ============ 合同条款填写模板 ============
    
    CONTRACT_CLAUSE_FILLING = PromptTemplate(
        name="合同条款智能填写",
        system_prompt="""你是一个资深的合同起草和填写专家。你的职责是根据业务数据，
精确填写合同中的各项条款和金额。

核心要求：
1. 【数字准确】所有金额、数量、日期必须精确，不能有错误
2. 【法律合规】填写内容必须符合合同法和相关行业规范
3. 【逻辑一致】不同条款之间的数据必须相互一致（如总价 = 单价 × 数量）
4. 【语法规范】保持法律文件的正式语气和结构
5. 【完整性检查】所有必填项都要检查和填写""",
        
        user_prompt_template="""请根据以下合同数据，填写合同模板中的相应内容。

【合同信息】
合同类型: {contract_type}
当前页面: 第{page_num}页

【提取的空白项】
{blank_items}

【业务数据】
{business_data}

【特殊说明】
{special_notes}

【输出要求】
1. 逐个分析每个空白项
2. 从业务数据中精确匹配相应信息
3. 检查数据的逻辑一致性
4. 输出JSON格式的修改列表"""
    )
    
    # ============ 表格数据填写模板 ============
    
    TABLE_DATA_FILLING = PromptTemplate(
        name="表格数据智能填充",
        system_prompt="""你是表格数据填写专家。你的任务是根据提供的数据，
正确填写Word文档中的表格。

关键原则：
1. 【行列匹配】确保数据填在正确的行和列
2. 【数据验证】验证数据类型是否正确（数字、日期、文本）
3. 【格式保留】保持表格的原始格式和结构
4. 【完整填充】填写所有相关数据
5. 【错误检查】发现并报告数据异常""",
        
        user_prompt_template="""请根据提供的数据，填写表格。

【表格信息】
表格名称: {table_name}
表格位置: 第{page_num}页
表格结构: 
{table_structure}

【表格数据】
{table_data}

【填充规则】
{filling_rules}

【输出格式】
返回修改后的表格XML和数据映射信息。"""
    )
    
    # ============ 自由文本生成模板 ============
    
    FREE_TEXT_GENERATION = PromptTemplate(
        name="自由文本智能生成",
        system_prompt="""你是一个优秀的文案和文本创作专家。你的任务是根据提供的信息，
生成符合上下文和格式要求的文本内容。

标准：
1. 【上下文匹配】生成的文本应与周围内容相辅相成
2. 【风格统一】保持文档的整体风格和语气
3. 【长度适当】生成的文本不应过长或过短
4. 【内容准确】基于提供的数据生成准确的内容
5. 【可读性】确保生成的文本易于理解""",
        
        user_prompt_template="""请根据以下信息，生成相应的文本内容。

【页面上下文】
页面标题: {page_title}
前一页内容摘要: {previous_content}
当前页主题: {page_topic}

【生成数据】
{generation_data}

【格式要求】
- 长度范围: {length_range}
- 语言风格: {language_style}
- 排版格式: {format_requirements}

【输出要求】
返回生成的文本内容，确保与整个文档风格一致。"""
    )
    
    # ============ XML格式验证和修复模板 ============
    
    XML_VALIDATION_FIX = PromptTemplate(
        name="XML格式验证和修复",
        system_prompt="""你是XML文档处理专家。你的任务是分析和修复Word文档的XML内容。

目标：
1. 【完整性检查】确保所有必需的XML标签完整
2. 【属性验证】检查XML属性是否正确
3. 【格式修复】修复不规范的XML结构
4. 【内容保护】确保修改过程中内容不丢失
5. 【性能优化】确保修复后的XML可被Word正确打开""",
        
        user_prompt_template="""请分析和修复以下XML片段。

【原始XML】
{original_xml}

【问题描述】
{issues_description}

【修复要求】
{fix_requirements}

【输出要求】
1. 分析问题根源
2. 修复XML结构
3. 验证修复结果
4. 输出修复后的XML"""
    )
    
    # ============ 数据提取和理解模板 ============
    
    DATA_EXTRACTION = PromptTemplate(
        name="页面数据提取和理解",
        system_prompt="""你是数据分析和提取专家。你的任务是从Word文档的XML中
提取和理解关键数据。

职责：
1. 【字段识别】识别页面中的所有可填写字段
2. 【类型判断】确定每个字段的数据类型
3. 【关系分析】分析字段之间的关系和依赖
4. 【优先级】确定填写的优先级和依赖顺序
5. 【验证规则】提出数据验证规则""",
        
        user_prompt_template="""请分析以下页面内容，提取关键字段信息。

【页面XML】
{page_xml}

【页面标题】
{page_title}

【分析要求】
1. 提取所有空白字段
2. 分析字段的数据类型
3. 判断字段的必填性
4. 找出字段之间的关联

【输出格式】
{{
    "page_number": {page_num},
    "page_title": "{page_title}",
    "fields": [
        {{
            "field_id": "字段ID",
            "field_name": "字段名称",
            "data_type": "数据类型",
            "required": true/false,
            "related_fields": ["字段1", "字段2"],
            "xpath": "XML路径"
        }},
        ...
    ],
    "validation_rules": ["规则1", "规则2"],
    "notes": "特殊说明"
}}"""
    )
    
    @classmethod
    def get_template(cls, template_name: str) -> PromptTemplate:
        """获取指定的提示词模板"""
        templates = {
            'tender_form': cls.TENDER_FORM_FILLING,
            'contract_clause': cls.CONTRACT_CLAUSE_FILLING,
            'table_data': cls.TABLE_DATA_FILLING,
            'free_text': cls.FREE_TEXT_GENERATION,
            'xml_validation': cls.XML_VALIDATION_FIX,
            'data_extraction': cls.DATA_EXTRACTION,
        }
        
        if template_name not in templates:
            raise ValueError(f"未知的模板: {template_name}. 可用模板: {list(templates.keys())}")
        
        return templates[template_name]
    
    @classmethod
    def list_templates(cls) -> List[str]:
        """列出所有可用的模板"""
        return [
            'tender_form',
            'contract_clause',
            'table_data',
            'free_text',
            'xml_validation',
            'data_extraction',
        ]


if __name__ == '__main__':
    print("可用的提示词模板:")
    for template_name in PromptLibrary.list_templates():
        print(f"  - {template_name}")
