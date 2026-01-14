import xml.etree.ElementTree as ET
import os
from lxml import etree

def split_word_by_pages(xml_file, output_dir='split_pages'):
    """
    将Word XML按页分割，保留原始格式（包括表格）。
    每一页都是一个包含完整XML声明和命名空间的独立文件。
    
    :param xml_file: 输入的XML文件路径
    :param output_dir: 输出目录
    """
    
    # 创建输出目录
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # 使用lxml保留完整格式和命名空间
    tree = etree.parse(xml_file)
    root = tree.getroot()
    
    # 定义命名空间
    ns = {
        'ns0': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'ns2': 'http://schemas.microsoft.com/office/word/2010/wordml'
    }
    
    # 找到body元素
    body = root.find('ns0:body', ns)
    if body is None:
        print("未找到body元素")
        return
    
    # 获取body中的所有直接子元素（段落和表格都在这里）
    all_elements = list(body)
    paragraphs = body.findall('ns0:p', ns)
    tables = body.findall('ns0:tbl', ns)
    
    print(f"总共找到 {len(paragraphs)} 个段落，{len(tables)} 个表格，共 {len(all_elements)} 个元素")
    
    # 找到页分割点（包含sectPr的段落索引）
    page_boundaries = []
    for idx, elem in enumerate(all_elements):
        # 检查是否是段落
        if elem.tag == '{%s}p' % ns['ns0']:
            pPr = elem.find('ns0:pPr', ns)
            if pPr is not None:
                sectPr = pPr.find('ns0:sectPr', ns)
                if sectPr is not None:
                    page_boundaries.append(idx)
    
    print(f"找到 {len(page_boundaries)} 个页面分割点（sectPr）")
    
    # 按页分割
    page_num = 1
    start_idx = 0
    
    for boundary_idx in page_boundaries:
        # 为当前页创建新的XML文档
        new_root = etree.Element(root.tag, nsmap=root.nsmap)
        new_body = etree.SubElement(new_root, '{%s}body' % ns['ns0'])
        
        # 复制这一页的所有元素（段落和表格）
        for idx in range(start_idx, boundary_idx + 1):
            elem = all_elements[idx]
            # 深度复制元素
            elem_copy = etree.fromstring(etree.tostring(elem))
            new_body.append(elem_copy)
        
        # 保存为文件
        output_file = os.path.join(output_dir, f'page_{page_num}.xml')
        new_tree = etree.ElementTree(new_root)
        new_tree.write(output_file, encoding='utf-8', xml_declaration=True, pretty_print=True)
        
        print(f"✓ 第 {page_num} 页已保存: {output_file} ({boundary_idx - start_idx + 1} 个元素)")
        
        page_num += 1
        start_idx = boundary_idx + 1
    
    # 处理最后一页（如果最后一段没有sectPr）
    if start_idx < len(all_elements):
        new_root = etree.Element(root.tag, nsmap=root.nsmap)
        new_body = etree.SubElement(new_root, '{%s}body' % ns['ns0'])
        
        for idx in range(start_idx, len(all_elements)):
            elem = all_elements[idx]
            # 深度复制元素
            elem_copy = etree.fromstring(etree.tostring(elem))
            new_body.append(elem_copy)
        
        output_file = os.path.join(output_dir, f'page_{page_num}.xml')
        new_tree = etree.ElementTree(new_root)
        new_tree.write(output_file, encoding='utf-8', xml_declaration=True, pretty_print=True)
        
        print(f"✓ 第 {page_num} 页已保存: {output_file} ({len(all_elements) - start_idx} 个元素)")


if __name__ == '__main__':
    split_word_by_pages('template.xml')
    print("\n✓ 页面分割完成！")
