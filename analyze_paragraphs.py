import logging
import json
import re
from lxml import etree
from openai import OpenAI

# --- 日志配置 ---
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# --- 模型配置信息 ---
LLM_CONFIG = {
    "api_key": "sk-VcqUMssTr1tzmkyEZ64N5Xba8k5jVCVm0ZYVMo2AqJKxvFIC",
    "base_url": "https://api.lkeap.cloud.tencent.com/v1",
    "model": "deepseek-v3-0324",
    "temperature": 0.6,
    "max_tokens": 8192
}

class WordSemanticParser:
    def __init__(self, xml_path):
        self.xml_path = xml_path
        self.namespaces = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'w14': 'http://schemas.microsoft.com/office/word/2010/wordml'
        }
        # 初始化 OpenAI 客户端
        self.client = OpenAI(
            api_key=LLM_CONFIG["api_key"],
            base_url=LLM_CONFIG["base_url"]
        )

    def get_raw_candidates(self):
        """第一阶段：从 XML 中提取所有包含下划线或占位符的候选段落"""
        try:
            tree = etree.parse(self.xml_path)
            logger.info(f"解析 XML 成功: {self.xml_path}")
        except Exception as e:
            logger.error(f"解析 XML 失败: {e}")
            return []

        candidates = []
        paragraphs = tree.xpath("//w:p", namespaces=self.namespaces)
        
        for p in paragraphs:
            pid = p.get(f"{{{self.namespaces['w14']}}}paraId")
            runs = p.xpath("./w:r", namespaces=self.namespaces)
            
            p_data = {"pid": pid, "full_text": "", "structure": []}
            has_feature = False
            text_parts = []

            for i, r in enumerate(runs):
                t_node = r.find("w:t", namespaces=self.namespaces)
                text = t_node.text if t_node is not None else ""
                text_parts.append(text)

                is_u = len(r.xpath("./w:rPr/w:u", namespaces=self.namespaces)) > 0
                is_placeholder = "${" in text
                
                if is_u or is_placeholder:
                    has_feature = True
                
                p_data["structure"].append({
                    "index": i + 1,
                    "text": text,
                    "type": "underline" if is_u else ("placeholder" if is_placeholder else "none")
                })

            if has_feature:
                p_data["full_text"] = "".join(text_parts)
                candidates.append(p_data)
        
        logger.info(f"提取到 {len(candidates)} 个潜在填写段落")
        return candidates

    def call_llm_service(self, candidates):
        """第二阶段：利用 DeepSeek 进行语义识别与结构化转换"""
        if not candidates:
            return []

        prompt = f"""
你是一个专业的文档解析专家。请分析以下从 Word XML 中提取的候选段落数据。
这些段落中包含用户需要填写的区域（由下划线或占位符表示）。

**任务要求：**
1. 修正 Label：根据 'full_text' 的语义，为每个填写位赋予准确的业务名称。
2. 逻辑分组：将同一行内的关联字段（如：年/月/日，姓名/性别/职务）归纳到同一个 logical_label 下。
3. 过滤：删除那些纯粹作为装饰线的下划线。
4. 务必保留原始的 pid 和 structure 中的 index，以便我回填 XPath。

**输入数据：**
{json.dumps(candidates, ensure_ascii=False)}

**必须返回且仅返回 JSON 格式数组，格式如下：**
[
  {{
    "pid": "段落ID",
    "logical_label": "该组字段的总称",
    "fields": [
      {{ "index": 对应structure中的index, "sub_label": "具体字段名", "type": "类型" }}
    ]
  }}
]
"""
        logger.info(f"发起模型请求 (Model: {LLM_CONFIG['model']})...")
        try:
            response = self.client.chat.completions.create(
                model=LLM_CONFIG["model"],
                messages=[
                    {"role": "system", "content": "你是一个只输出 JSON 的文档解析助手。"},
                    {"role": "user", "content": prompt}
                ],
                temperature=LLM_CONFIG["temperature"],
                max_tokens=LLM_CONFIG["max_tokens"]
            )
            
            content = response.choices[0].message.content
            # 清洗模型可能返回的 Markdown 代码块标签
            clean_json = re.sub(r'```json\s*|\s*```', '', content).strip()
            result = json.loads(clean_json)
            logger.info("模型响应解析成功")
            return result

        except Exception as e:
            logger.error(f"调用 LLM 失败: {e}")
            return []

    def final_assemble(self, llm_results):
        """第三阶段：将语义结果与物理 XPath 缝合"""
        final_data = []
        for item in llm_results:
            pid = item.get("pid")
            logical_label = item.get("logical_label")
            
            group_info = {
                "pid": pid,
                "group_name": logical_label,
                "slots": []
            }
            
            for f in item.get("fields", []):
                idx = f["index"]
                # 构造准确的 XPath
                xpath = f"/w:document/w:body/w:p[@w14:paraId='{pid}']/w:r[{idx}]/w:t"
                group_info["slots"].append({
                    "label": f["sub_label"],
                    "xpath": xpath,
                    "type": f["type"]
                })
            final_data.append(group_info)
        
        return final_data

# --- 执行主程序 ---
if __name__ == "__main__":
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
    else:
        logger.warning("未发现任何潜在的填写内容。")