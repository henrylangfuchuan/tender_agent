"""
将XML文档转换为Word (.docx) 文件
"""
import zipfile
import shutil
import os
from lxml import etree

def xml_to_docx(xml_file, output_docx, template_docx='template.docx'):
    """
    将修改后的XML转换为Word文档。
    使用原始template.docx作为模板，替换其中的document.xml。
    
    :param xml_file: 要转换的XML文件（merged_document.xml）
    :param output_docx: 输出的Word文件名
    :param template_docx: 模板Word文件（用来提取其他资源）
    """
    
    if not os.path.exists(template_docx):
        print(f"错误：未找到模板文件 {template_docx}")
        return False
    
    if not os.path.exists(xml_file):
        print(f"错误：未找到XML文件 {xml_file}")
        return False
    
    try:
        # 创建临时目录
        temp_dir = '_temp_docx_'
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
        os.makedirs(temp_dir)
        
        print("解析模板Word文档...")
        # 解压模板docx文件
        with zipfile.ZipFile(template_docx, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        
        print("替换document.xml内容...")
        # 读取新的XML文档
        new_tree = etree.parse(xml_file)
        
        # 保存XML到解压的文件
        word_document_path = os.path.join(temp_dir, 'word', 'document.xml')
        new_tree.write(word_document_path, encoding='utf-8', xml_declaration=True)
        
        print("重新打包Word文档...")
        # 重新创建docx文件（zipfile）
        with zipfile.ZipFile(output_docx, 'w', zipfile.ZIP_DEFLATED) as docx_zip:
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, temp_dir)
                    docx_zip.write(file_path, arcname)
        
        # 清理临时文件
        shutil.rmtree(temp_dir)
        
        print(f"✓ 转换成功！已保存为: {output_docx}")
        return True
        
    except Exception as e:
        print(f"✗ 转换失败: {e}")
        # 清理临时文件
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
        return False


if __name__ == '__main__':
    # 假设你已经合并了XML并修改了内容
    success = xml_to_docx('filled_result.xml', 'output_document.docx', 'template.docx')
    if success:
        print("\n现在你可以用Word打开 output_document.docx 了！")
