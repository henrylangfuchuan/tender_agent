"""
合并所有分割的页面XML文件为一个完整的XML文档
"""
import os
from lxml import etree
import glob

def merge_pages(input_dir='split_pages', output_file='merged_document.xml'):
    """
    将所有页面XML合并为一个文档
    
    :param input_dir: 包含分割页面的目录
    :param output_file: 输出合并后的XML文件名
    """
    
    # 获取所有页面文件，按页码排序
    page_files = sorted(glob.glob(os.path.join(input_dir, 'page_*.xml')),
                       key=lambda x: int(x.split('page_')[1].split('.')[0]))
    
    if not page_files:
        print(f"未找到页面文件在 {input_dir}")
        return
    
    print(f"找到 {len(page_files)} 个页面文件")
    
    # 读取第一个页面作为模板
    first_tree = etree.parse(page_files[0])
    root = first_tree.getroot()
    
    # 定义命名空间
    ns = {
        'ns0': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    }
    
    # 获取body元素
    body = root.find('ns0:body', ns)
    
    # 清空body中的内容
    for child in body:
        body.remove(child)
    
    print("合并页面内容...")
    
    # 遍历所有页面文件，合并内容
    for idx, page_file in enumerate(page_files, 1):
        tree = etree.parse(page_file)
        page_root = tree.getroot()
        page_body = page_root.find('ns0:body', ns)
        
        if page_body is None:
            print(f"⚠ 第 {idx} 页没有body元素，跳过")
            continue
        
        # 复制这个页面的所有段落到主body
        for para in page_body.findall('ns0:p', ns):
            # 深度复制元素
            para_copy = etree.fromstring(etree.tostring(para))
            body.append(para_copy)
        
        print(f"✓ 第 {idx} 页已合并")
    
    # 保存合并后的文档
    first_tree.write(output_file, encoding='utf-8', xml_declaration=True, pretty_print=False)
    print(f"\n✓ 合并完成！已保存到: {output_file}")
    
    return output_file


if __name__ == '__main__':
    merge_pages()
