import logging
import json
from analyze_paragraphs import WordSemanticParser
from lxml import etree
from parse_docx import docx_to_xml
from xml_to_docx import xml_to_docx

# --- 配置日志 ---
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

class XMLDataFiller:
    def __init__(self, source_xml_path):
        self.source_xml_path = source_xml_path
        self.namespaces = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'w14': 'http://schemas.microsoft.com/office/word/2010/wordml'
        }
        self.tree = etree.parse(self.source_xml_path)

    def fill(self, fill_tasks, output_path):
        """
        执行填充任务
        :param fill_tasks: 之前步骤生成的结构化结果，并附带了 mock 值的 list
        :param output_path: 保存路径
        """
        logger.info("开始执行数据填充任务...")
        
        fill_count = 0
        for group in fill_tasks:
            group_name = group.get("group_name")
            for slot in group.get("slots", []):
                xpath = slot.get("xpath")
                fill_value = slot.get("mock_value") # 我们需要在填入前把 mock 数据放进这个字段
                
                if not fill_value:
                    logger.warning(f"跳过空值: {slot.get('label')}")
                    continue

                # 通过 XPath 找到对应节点
                nodes = self.tree.xpath(xpath, namespaces=self.namespaces)
                if nodes:
                    target_node = nodes[0]
                    # 更新文本内容
                    # 注意：Word 可能会使用 xml:space="preserve" 属性，这里直接赋值 text 会保留它
                    target_node.text = str(fill_value)
                    fill_count += 1
                    logger.info(f"已填充 [{group_name}-{slot.get('label')}]: {fill_value}")
                else:
                    logger.error(f"无法定位 XPath 节点: {xpath}")

        # 保存结果
        self.tree.write(output_path, encoding='utf-8', xml_declaration=True, standalone=True)
        logger.info(f"填充完成！共填充 {fill_count} 处，结果已保存至: {output_path}")

# --- 模拟业务逻辑：生成 Mock 数据 ---
def prepare_mock_data(semantic_results):
    """
    根据语义标签自动生成假数据
    """
    # 这里定义你的 mock 规则
    mock_registry = {
        "项目名称": "2026年智慧城市建设项目",
        "包编号": "A-001",
        "投标人": "科技有限公司",
        "投标人名称": "科技有限公司",
        "姓名": "张三",
        "性别": "男",
        "年龄": "35",
        "职务": "技术总监",
        "年": "2026",
        "月": "01",
        "日": "15",
        "授权委托人姓名": "李四"
    }

    # 将 mock 数据注入到之前的解析结果中
    for group in semantic_results:
        for slot in group["slots"]:
            label = slot["label"]
            # 优先匹配 sub_label，再匹配 group_name
            val = mock_registry.get(label) or mock_registry.get(group["group_name"], "测试数据")
            slot["mock_value"] = val
            
    return semantic_results

# --- 主流程 ---
if __name__ == "__main__":
    semantic_output = []

    file_path = "template2.docx"
    docx_to_xml(file_path)

    # 初始化解析器
    parser = WordSemanticParser("template.xml")
    
    # 1. 规则提取
    raw_data = parser.get_raw_candidates()
    
    # 2. 语义提纯（调用 DeepSeek）
    if raw_data:
        semantic_results = parser.call_llm_service(raw_data)
        
        # 3. 结果组装
        final_output = parser.final_assemble(semantic_results)
        
        # 打印最终 JSON
        print("\n=== 最终语义解析结果 ===")
        print(json.dumps(final_output, indent=2, ensure_ascii=False))
        semantic_output = final_output
    # 2. 注入 Mock 数据
    data_to_fill = prepare_mock_data(semantic_output)

    # 3. 执行填充并保存
    filler = XMLDataFiller("template.xml")
    filler.fill(data_to_fill, "filled_result.xml")

    success = xml_to_docx('filled_result.xml', 'output_document.docx', 'template2.docx')
    if success:
        print("\n现在你可以用Word打开 output_document.docx 了！")