import zipfile

def docx_to_xml(path, output_xml='template.xml'):
    xml_content = ""
    with zipfile.ZipFile(path) as docx:
        # 读取 Word 主文档内容 XML
        xml_content = docx.read('word/document.xml').decode('utf-8')
    with open(output_xml, "w", encoding="utf-8") as f:
        f.write(xml_content)

file_path = "template2.docx"
xml_data = docx_to_xml(file_path)