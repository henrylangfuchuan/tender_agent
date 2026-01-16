from lxml import etree
from openai import OpenAI
import json
import re

# ------------------------------
# 解析空单元格 + 获取上下文 label
# ------------------------------
def find_empty_cells_with_context(xml_file):
    """
    解析 Word XML 表格，获取空单元格及上下文 label。
    返回格式：
    [
        {
            "table_index": 0,
            "row_index": 1,
            "cell_index": 1,
            "label": "投标人名称",
            "full_label": "投标人名称",  # 可根据上下文生成完整 label
            "text": ""
        },
        ...
    ]
    """
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    tree = etree.parse(xml_file)
    root = tree.getroot()

    results = []

    tables = root.xpath("//w:tbl", namespaces=ns)
    for t_index, table in enumerate(tables):
        rows = table.xpath(".//w:tr", namespaces=ns)
        for r_index, row in enumerate(rows):
            cells = row.xpath(".//w:tc", namespaces=ns)
            row_labels = []
            for c_index, cell in enumerate(cells):
                texts = cell.xpath(".//w:t", namespaces=ns)
                cell_text = "".join([t.text for t in texts if t.text]).strip()
                row_labels.append(cell_text)

            # 再遍历一次填充空单元格信息
            for c_index, cell in enumerate(cells):
                texts = cell.xpath(".//w:t", namespaces=ns)
                cell_text = "".join([t.text for t in texts if t.text]).strip()

                if cell_text == "" and c_index > 0:
                    # 左邻单元格文本
                    label = row_labels[c_index - 1]
                    # 上下文完整 label，可加入整行信息区分同名字段
                    context_label = "_".join([l for l in row_labels[:c_index] if l.strip()])
                    results.append({
                        "table_index": t_index,
                        "row_index": r_index,
                        "cell_index": c_index,
                        "label": label,
                        "full_label": context_label,
                        "text": cell_text
                    })
    return results

# ------------------------------
# 填充单元格函数
# ------------------------------
def fill_cell(cell, value):
    ns_uri = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    
    # 找 <w:p>
    p = cell.find(".//w:p", namespaces={'w': ns_uri})
    if p is None:
        p = etree.SubElement(cell, f"{{{ns_uri}}}p")
    
    # 找 <w:r>
    r = p.find(".//w:r", namespaces={'w': ns_uri})
    if r is None:
        r = etree.SubElement(p, f"{{{ns_uri}}}r")
    
    # 找 <w:t>
    t = r.find(".//w:t", namespaces={'w': ns_uri})
    if t is None:
        t = etree.SubElement(r, f"{{{ns_uri}}}t")
    
    t.text = str(value)

# ------------------------------
# 填充表格函数（使用 full_label）
# ------------------------------
def fill_table_cells_by_label(xml_file, fill_dict, output_file=None):
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    tree = etree.parse(xml_file)
    root = tree.getroot()

    tables = root.xpath("//w:tbl", namespaces=ns)
    for t_index, table in enumerate(tables):
        rows = table.xpath(".//w:tr", namespaces=ns)
        for r_index, row in enumerate(rows):
            cells = row.xpath(".//w:tc", namespaces=ns)
            row_labels = []
            for c_index, cell in enumerate(cells):
                texts = cell.xpath(".//w:t", namespaces=ns)
                cell_text = "".join([t.text for t in texts if t.text]).strip()
                row_labels.append(cell_text)

            for c_index, cell in enumerate(cells):
                texts = cell.xpath(".//w:t", namespaces=ns)
                cell_text = "".join([t.text for t in texts if t.text]).strip()

                if cell_text == "" and c_index > 0:
                    # 左邻单元格
                    label = row_labels[c_index - 1]
                    # 上下文完整 label
                    full_label = "_".join([l for l in row_labels[:c_index] if l.strip()])
                    if full_label in fill_dict:
                        fill_cell(cell, fill_dict[full_label])

    if output_file is None:
        output_file = xml_file
    tree.write(output_file, encoding="utf-8", xml_declaration=True, pretty_print=True)
    return tree

# ------------------------------
# LLM 映射 full_label
# ------------------------------
def normalize_labels_with_llm_openai(table_full_labels, user_fields, api_key, base_url, model="gpt-4"):
    """
    使用 OpenAI LLM 将用户字段归一化到表格 full_label（上下文 label）
    """
    client = OpenAI(api_key=api_key, base_url=base_url)

    user_field_keys = list(user_fields.keys())

    prompt = f"""
我有以下表格字段（带上下文）：{table_full_labels}
用户提供的数据字段：{user_field_keys}
请帮我生成一个映射字典，key 为表格 full_label，value 为用户字段名。
如果表格字段没有对应用户字段可以不映射。
请返回严格 JSON，例如：
{{"联系方式_电话": "联系人电话", "法定代表人_电话": "法人电话"}}
"""

    response = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": "你是一个数据匹配助手，将用户字段映射到表格字段。"},
            {"role": "user", "content": prompt}
        ],
        temperature=0.5
    )

    text = response.choices[0].message.content.strip()
    match = re.search(r"```json\s*(\{.*\})\s*```", text, re.DOTALL)
    json_str = match.group(1) if match else text

    try:
        mapping = json.loads(json_str)
    except Exception as e:
        print("⚠️ 解析 LLM 输出 JSON 出错:", e)
        print("输出内容:", text)
        mapping = {}

    return mapping

# ------------------------------
# 示例流程
# ------------------------------
if __name__ == '__main__':
    xml_path = "template.xml"

    # 1️⃣ 获取空单元格及上下文 full_label
    empty_cells = find_empty_cells_with_context(xml_path)
    for cell in empty_cells:
        print(f"表格 {cell['table_index']} 行 {cell['row_index']} 列 {cell['cell_index']} 为空，标签: {cell['full_label']}")
    table_full_labels = list(set([cell['full_label'] for cell in empty_cells]))

    # 2️⃣ 用户数据
    user_data = {
        "投标人名称": "张三公司",
        "联系方式_联系人": "李四",
        "联系方式_电话": "123456789",
        "法定代表人_电话": "987654321",
        "技术负责人_电话": "1122334455",
        "注册资本": "500万",
        "高级职称人员": "王五"
    }

    # 3️⃣ 使用 LLM 生成 full_label 映射
    mapping = normalize_labels_with_llm_openai(
        table_full_labels,
        user_data,
        api_key="你的OpenAI API Key",
        base_url="你的OpenAI Base URL",
        model="deepseek-v3-0324"
    )

    # 4️⃣ 构建填充字典
    fill_dict = {}
    for full_label, user_field in mapping.items():
        if user_field in user_data:
            fill_dict[full_label] = user_data[user_field]

    # 5️⃣ 填充表格
    fill_table_cells_by_label(xml_path, fill_dict, output_file="filled_template.xml")
    print("填充完成！输出文件: filled_template.xml")
