"""
修改分割的页面内容的示例脚本
使用这个脚本作为模板来修改你的页面内容
"""
import os
from lxml import etree
import glob

def modify_pages_content(input_dir='split_pages', modifications_func=None):
    """
    遍历所有页面并修改其内容
    
    :param input_dir: 输入目录
    :param modifications_func: 自定义修改函数，接收 (page_num, root) 并修改root
    
    示例修改函数:
    def my_modifications(page_num, root):
        ns = {'ns0': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        # 在第2页找到所有文本并替换
        if page_num == 2:
            for t in root.findall('.//ns0:t', ns):
                if t.text:
                    t.text = t.text.replace('原文本', '新文本')
    """
    
    # 获取所有页面文件
    page_files = sorted(glob.glob(os.path.join(input_dir, 'page_*.xml')),
                       key=lambda x: int(x.split('page_')[1].split('.')[0]))
    
    ns = {
        'ns0': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'ns2': 'http://schemas.microsoft.com/office/word/2010/wordml'
    }
    
    for page_file in page_files:
        page_num = int(page_file.split('page_')[1].split('.')[0])
        
        print(f"处理第 {page_num} 页...")
        
        tree = etree.parse(page_file)
        root = tree.getroot()
        
        # 如果提供了自定义修改函数，则使用它
        if modifications_func:
            modifications_func(page_num, root, ns)
        else:
            # 默认修改示例
            default_modifications(page_num, root, ns)
        
        # 保存修改后的页面
        tree.write(page_file, encoding='utf-8', xml_declaration=True)
        print(f"✓ 第 {page_num} 页已保存")
    
    print("\n✓ 所有页面修改完成！")


def default_modifications(page_num, root, ns):
    """
    默认的修改示例
    根据你的需求修改这个函数
    """
    
    # 示例1：在所有文本前面添加页码标记
    for para in root.findall('.//ns0:p', ns):
        texts = para.findall('.//ns0:t', ns)
        if texts and texts[0].text:
            # 在第一个文本前添加"[第X页]"标记
            # texts[0].text = f"[第{page_num}页] " + texts[0].text
            pass
    
    # 示例2：替换特定的占位符
    for t in root.findall('.//ns0:t', ns):
        if t.text:
            # 替换占位符
            if '${' in t.text:
                # 这是一个占位符，可以进行替换
                # t.text = t.text.replace('${项目名称}', '我的项目')
                pass
    
    # 示例3：根据页码修改特定页面的内容
    if page_num == 1:
        # 修改首页内容
        for t in root.findall('.//ns0:t', ns):
            if t.text and '第六章' in t.text:
                # 修改标题
                # t.text = '修改后的标题'
                pass
    
    # 示例4：添加新段落
    # if page_num == 2:
    #     body = root.find('ns0:body', ns)
    #     # 这里可以添加新的段落...
    #     pass


def custom_replace_text(input_dir='split_pages', search_text='', replace_text=''):
    """
    便捷函数：替换所有页面中的特定文本
    
    :param input_dir: 输入目录
    :param search_text: 要查找的文本
    :param replace_text: 要替换的文本
    """
    
    def replacer(page_num, root, ns):
        for t in root.findall('.//ns0:t', ns):
            if t.text and search_text in t.text:
                t.text = t.text.replace(search_text, replace_text)
                print(f"  ✓ 第 {page_num} 页: '{search_text}' → '{replace_text}'")
    
    modify_pages_content(input_dir, replacer)


def extract_all_text(input_dir='split_pages'):
    """
    提取所有页面的文本内容（用于查看和修改规划）
    """
    
    page_files = sorted(glob.glob(os.path.join(input_dir, 'page_*.xml')),
                       key=lambda x: int(x.split('page_')[1].split('.')[0]))
    
    ns = {'ns0': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    
    print("\n=== 文档内容预览 ===\n")
    
    for page_file in page_files:
        page_num = int(page_file.split('page_')[1].split('.')[0])
        
        tree = etree.parse(page_file)
        root = tree.getroot()
        
        texts = []
        for t in root.findall('.//ns0:t', ns):
            if t.text and t.text.strip():
                texts.append(t.text)
        
        if texts:
            print(f"【第 {page_num} 页】")
            print(' '.join(texts)[:100] + ('...' if len(' '.join(texts)) > 100 else ''))
            print()


if __name__ == '__main__':
    print("修改页面内容脚本\n")
    
    # 选项1：查看所有文本
    print("1. 预览文档内容...")
    extract_all_text()
    
    # 选项2：使用默认修改（需要自己编写default_modifications函数）
    # modify_pages_content()
    
    # 选项3：快速替换文本
    # 例如：custom_replace_text('split_pages', '原文本', '新文本')
    
    print("\n提示：")
    print("1. 编辑 default_modifications() 函数来自定义修改")
    print("2. 或调用 custom_replace_text() 进行快速文本替换")
    print("3. 修改完成后，运行 merge_pages.py 合并")
    print("4. 最后运行 xml_to_docx.py 转为Word文档")
