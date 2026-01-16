import logging
import json
from lxml import etree

# --- 配置日志 ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class WordTableParser:
    def __init__(self, xml_path):
        self.xml_path = xml_path
        self.namespaces = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'w14': 'http://schemas.microsoft.com/office/word/2010/wordml'
        }
        self.tree = etree.parse(self.xml_path)

    def extract_table_candidates(self):
        """
        提取表格中的潜在填写位，并记录其行列坐标
        """
        tables = self.tree.xpath("//w:tbl", namespaces=self.namespaces)
        logger.info(f"发现表格数量: {len(tables)}")
        
        table_candidates = []

        for t_idx, tbl in enumerate(tables):
            rows = tbl.xpath("./w:tr", namespaces=self.namespaces)
            table_data = {"table_index": t_idx, "rows": []}
            
            for r_idx, tr in enumerate(rows):
                row_data = {"row_index": r_idx, "cells": []}
                cells = tr.xpath("./w:tc", namespaces=self.namespaces)
                
                for c_idx, tc in enumerate(cells):
                    # 获取单元格内所有文本
                    cell_text = "".join(tc.xpath(".//w:t/text()", namespaces=self.namespaces))
                    
                    # 判定该单元格是否包含填写位（下划线或占位符）
                    has_u = len(tc.xpath(".//w:rPr/w:u", namespaces=self.namespaces)) > 0
                    has_placeholder = "${" in cell_text
                    
                    cell_info = {
                        "col_index": c_idx,
                        "text": cell_text,
                        "is_fillable": has_u or has_placeholder,
                        # 记录单元格内段落的 pid，用于精确回填
                        "paragraphs": []
                    }
                    
                    if cell_info["is_fillable"]:
                        ps = tc.xpath(".//w:p", namespaces=self.namespaces)
                        for p in ps:
                            pid = p.get(f"{{{self.namespaces['w14']}}}paraId")
                            cell_info["paragraphs"].append({"pid": pid})
                            
                    row_data["cells"].append(cell_info)
                table_data["rows"].append(row_data)
            table_candidates.append(table_data)
            
        return table_candidates

    def build_table_prompt(self, table_candidates):
        """
        构建针对表格的 LLM Prompt
        """
        prompt = f"""
你是一个专业的文档解析助手。请分析以下从 Word 表格中提取的结构化数据。
这些数据以表格、行、列的形式排列。

**任务要求：**
1. 识别填写位：找到 `is_fillable` 为 true 的单元格。
2. 语义关联：根据该单元格在表格中的位置（左侧单元格的内容或所属列的表头内容），给出准确的 logical_label。
3. 过滤：忽略那些看起来只是表格边框或无关紧要的装饰位。

**输入数据：**
{json.dumps(table_candidates, ensure_ascii=False)}

**返回格式（严格 JSON 数组）：**
[
  {{
    "table_index": 0,
    "row_index": 1,
    "col_index": 2,
    "logical_label": "法定代表人电话",
    "pids": ["段落ID"],
    "reason": "左侧单元格内容为'联系电话'"
  }}
]
"""
        return prompt

    def fill_table_mock(self, llm_results, mock_registry):
        """
        根据 LLM 结果和 Mock 注册表生成回填任务
        """
        fill_tasks = []
        for item in llm_results:
            label = item["logical_label"]
            mock_val = mock_registry.get(label, "表格测试数据")
            
            # 针对该单元格内的每个段落生成 XPath
            for pid in item["pids"]:
                # 注意：表格内的 XPath 定位。为简单起见，我们直接用 pid 定位，这在表格内依然有效
                xpath = f"//w:p[@w14:paraId='{pid}']//w:r[descendant::w:u or contains(., '${{')]//w:t"
                fill_tasks.append({
                    "label": label,
                    "xpath": xpath,
                    "mock_value": mock_val
                })
        return fill_tasks

# --- 使用示例 ---
if __name__ == "__main__":
    parser = WordTableParser("template.xml")
    
    # 1. 提取表格候选
    tables_data = parser.extract_table_candidates()
    
    # 2. 生成 Prompt (发送给 DeepSeek)
    table_prompt = parser.build_table_prompt(tables_data)
    
    # 3. 假设 LLM 返回了如下结果
    mock_llm_table_res = [
        {
            "table_index": 0,
            "row_index": 2,
            "col_index": 1,
            "logical_label": "注册资本",
            "pids": ["5A2B3C4D"],
            "reason": "位于'注册资本'行对应的数值列"
        }
    ]
    
    # 4. 准备 Mock 数据并回填
    mock_registry = {"注册资本": "500万元"}
    fill_tasks = parser.fill_table_mock(mock_llm_table_res, mock_registry)
    
    # 5. 调用之前的 XMLDataFiller 进行回填即可
    print(f"生成的表格回填任务: {json.dumps(fill_tasks, indent=2, ensure_ascii=False)}")