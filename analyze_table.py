import xml.etree.ElementTree as ET
import json
from lxml import etree
import os

class TableAnalyzer:
    def __init__(self, xml_file):
        self.xml_file = xml_file
        self.ns = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'w14': 'http://schemas.microsoft.com/office/word/2010/wordml'
        }
        self.tree = etree.parse(xml_file)
        self.root = self.tree.getroot()
        
    def get_cell_text(self, cell):
        """获取单元格中的所有文本"""
        texts = []
        for t in cell.findall('.//w:t', self.ns):
            if t.text:
                texts.append(t.text)
        return ''.join(texts).strip()
    
    def analyze_table(self, table_index=0):
        """分析表格并返回所有单元格的坐标信息"""
        tables = self.root.findall('.//w:tbl', self.ns)
        
        if not tables or table_index >= len(tables):
            return None
            
        table = tables[table_index]
        table_xpath = f"/w:document/w:body/w:tbl[{table_index + 1}]"
        
        cells_data = []
        rows = table.findall('w:tr', self.ns)
        
        for row_idx, row in enumerate(rows):
            cells = row.findall('w:tc', self.ns)
            
            for col_idx, cell in enumerate(cells):
                cell_text = self.get_cell_text(cell)
                
                # 构建XPath - 不同方式获取
                cell_xpath = f"{table_xpath}/w:tr[{row_idx + 1}]/w:tc[{col_idx + 1}]"
                
                # 获取gridSpan和vMerge信息
                tc_pr = cell.find('w:tcPr', self.ns)
                grid_span = 1
                v_merge = None
                
                if tc_pr is not None:
                    gs = tc_pr.find('w:gridSpan', self.ns)
                    if gs is not None:
                        grid_span = int(gs.get('{%s}val' % self.ns['w'], '1'))
                    
                    vm = tc_pr.find('w:vMerge', self.ns)
                    if vm is not None:
                        v_merge = vm.get('{%s}val' % self.ns['w'], 'continue')
                
                cell_info = {
                    "cell_id": f"cell_{row_idx}_{col_idx}",
                    "row": row_idx,
                    "col": col_idx,
                    "grid_span": grid_span,
                    "v_merge": v_merge,
                    "xpath": cell_xpath,
                    "label": cell_text,
                    "is_empty": len(cell_text) == 0,
                    "row_display": f"第{row_idx + 1}行",
                    "col_display": f"第{col_idx + 1}列",
                    "position": f"({row_idx}, {col_idx})"
                }
                
                cells_data.append(cell_info)
        
        return {
            "table_index": table_index,
            "table_name": "(一)基本情况表",
            "total_rows": len(rows),
            "total_cols": len(rows[0].findall('w:tc', self.ns)) if rows else 0,
            "cells": cells_data
        }
    
    def save_to_json(self, output_file, table_data):
        """保存为JSON文件"""
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(table_data, f, ensure_ascii=False, indent=2)
        print(f"✓ 表格坐标数据已保存到: {output_file}")


def main():
    xml_file = 'split_pages/page_8.xml'
    output_file = 'table_coordinates.json'
    
    if not os.path.exists(xml_file):
        print(f"❌ 文件不存在: {xml_file}")
        return
    
    analyzer = TableAnalyzer(xml_file)
    table_data = analyzer.analyze_table(table_index=0)
    
    if table_data:
        print(f"\n📊 表格分析结果:")
        print(f"  表名称: {table_data['table_name']}")
        print(f"  行数: {table_data['total_rows']}")
        print(f"  列数: {table_data['total_cols']}")
        print(f"  总单元格数: {len(table_data['cells'])}")
        print(f"  空白单元格数: {sum(1 for c in table_data['cells'] if c['is_empty'])}")
        
        print(f"\n📍 单元格坐标信息示例:")
        for i, cell in enumerate(table_data['cells'][:5]):
            print(f"\n  {i+1}. {cell['cell_id']}")
            print(f"     位置: {cell['position']}")
            print(f"     XPath: {cell['xpath']}")
            print(f"     标签/内容: {cell['label']}")
            print(f"     是否空白: {cell['is_empty']}")
        
        print(f"\n  ... (共 {len(table_data['cells'])} 个单元格)\n")
        
        analyzer.save_to_json(output_file, table_data)
        
        # 输出关键单元格的详细坐标
        print(f"📋 关键字段的坐标:")
        for cell in table_data['cells']:
            if cell['label'] and not cell['is_empty']:
                print(f"  {cell['label']:15} -> {cell['position']:12} -> {cell['xpath']}")
    else:
        print("❌ 未找到表格")


if __name__ == '__main__':
    main()
