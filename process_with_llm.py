"""
XML解析和大模型处理脚本
从分割的页面XML中提取内容 → 调用大模型 → 更新XML
"""
import json
import os
import glob
from lxml import etree
from typing import Dict, List, Any, Optional
from llm_connector import LLMConnector, LLMConfig
from prompt_library import PromptLibrary


class XMLPageAnalyzer:
    """页面XML分析器"""
    
    def __init__(self, page_xml_path: str):
        self.page_path = page_xml_path
        self.tree = etree.parse(page_xml_path)
        self.root = self.tree.getroot()
        self.ns = {
            'ns0': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'ns2': 'http://schemas.microsoft.com/office/word/2010/wordml'
        }
    
    def extract_text_content(self) -> str:
        """提取页面的全部文本内容"""
        texts = []
        for t in self.root.findall('.//ns0:t', self.ns):
            if t.text and t.text.strip():
                texts.append(t.text)
        return ' '.join(texts)
    
    def find_blank_fields(self) -> List[Dict[str, Any]]:
        """找出页面中的空白字段"""
        blank_fields = []
        
        for idx, t in enumerate(self.root.findall('.//ns0:t', self.ns)):
            if t.text:
                # 检查是否是空白字段（如下划线、虚线、方括号等）
                text = t.text.strip()
                if text in ['_', '__', '___', '____', '—', '、', '[]', '[ ]', '【】', '【  】']:
                    para = t.getparent().getparent().getparent()
                    para_text = self._get_paragraph_context(para)
                    
                    blank_fields.append({
                        'field_id': f'blank_{idx}',
                        'original_value': text,
                        'context': para_text,
                        'xpath': self._get_xpath(t),
                        'element_ref': t
                    })
        
        return blank_fields
    
    def find_placeholder_fields(self) -> List[Dict[str, Any]]:
        """找出页面中的占位符字段 (${...})"""
        placeholder_fields = []
        
        for idx, t in enumerate(self.root.findall('.//ns0:t', self.ns)):
            if t.text and '${' in t.text and '}' in t.text:
                import re
                matches = re.findall(r'\$\{([^}]+)\}', t.text)
                
                for match in matches:
                    placeholder_fields.append({
                        'field_id': f'placeholder_{idx}',
                        'placeholder_name': match,
                        'original_value': f'${{{match}}}',
                        'context': self._get_paragraph_context(t.getparent().getparent()),
                        'xpath': self._get_xpath(t),
                        'element_ref': t,
                        'full_text': t.text
                    })
        
        return placeholder_fields
    
    def get_page_info(self) -> Dict[str, Any]:
        """获取页面信息"""
        return {
            'text_content': self.extract_text_content(),
            'blank_fields': self.find_blank_fields(),
            'placeholder_fields': self.find_placeholder_fields(),
            'page_xml': etree.tostring(self.root, encoding='utf-8', pretty_print=True).decode('utf-8')[:2000]  # 前2000字符
        }
    
    def _get_paragraph_context(self, para_elem) -> str:
        """获取段落的上下文文本"""
        texts = []
        for t in para_elem.findall('.//ns0:t', self.ns):
            if t.text:
                texts.append(t.text)
        return ''.join(texts)
    
    def _get_xpath(self, elem) -> str:
        """获取元素的XPath"""
        try:
            return self.tree.getpath(elem)
        except:
            return "unknown"
    
    def apply_updates(self, updates: List[Dict[str, str]]):
        """
        应用XML更新
        
        :param updates: 更新列表，格式：
            [{"old_text": "旧文本", "new_text": "新文本", "xpath": "xpath"}]
        """
        for update in updates:
            if 'xpath' in update:
                try:
                    elem = self.root.xpath(update['xpath'], namespaces=self.ns)[0]
                    if elem.text and update['old_text'] in elem.text:
                        elem.text = elem.text.replace(update['old_text'], update['new_text'])
                except:
                    # 尝试全局搜索替换
                    for t in self.root.findall('.//ns0:t', self.ns):
                        if t.text and update['old_text'] in t.text:
                            t.text = t.text.replace(update['old_text'], update['new_text'])
            else:
                # 全局替换
                for t in self.root.findall('.//ns0:t', self.ns):
                    if t.text and update.get('old_text') in (t.text or ''):
                        t.text = t.text.replace(update['old_text'], update['new_text'])
    
    def save(self):
        """保存修改后的XML"""
        self.tree.write(self.page_path, encoding='utf-8', xml_declaration=True)


class LLMPageProcessor:
    """LLM页面处理器"""
    
    def __init__(self, llm_connector: LLMConnector):
        self.llm = llm_connector
    
    def process_page_with_llm(self, 
                             page_num: int,
                             page_info: Dict[str, Any],
                             data_context: Dict[str, Any],
                             template_name: str = 'tender_form') -> Dict[str, Any]:
        """
        使用LLM处理页面
        
        :param page_num: 页码
        :param page_info: 页面信息
        :param data_context: 数据上下文
        :param template_name: 使用的提示词模板
        :return: 处理结果
        """
        print(f"  🤖 第{page_num}页: 调用大模型处理...")
        
        template = PromptLibrary.get_template(template_name)
        
        # 准备提示词参数
        prompt_vars = {
            'page_num': page_num,
            'page_title': data_context.get('page_title', ''),
            'fields_to_fill': self._format_fields(page_info.get('blank_fields', []) + 
                                                  page_info.get('placeholder_fields', [])),
            'provided_data': self._format_data(data_context.get('data', {})),
        }
        
        system_prompt, user_prompt = template.format(**prompt_vars)
        
        try:
            # 调用LLM
            response = self.llm.call(user_prompt, system_prompt)
            
            # 解析LLM响应
            result = self._parse_llm_response(response)
            
            return {
                'page_num': page_num,
                'status': 'success',
                'updates': result.get('updates', []),
                'filled_fields': result.get('filled_fields', []),
                'unfilled_fields': result.get('unfilled_fields', []),
                'raw_response': response
            }
        
        except Exception as e:
            print(f"  ❌ 第{page_num}页处理失败: {e}")
            return {
                'page_num': page_num,
                'status': 'failed',
                'error': str(e)
            }
    
    def _format_fields(self, fields: List[Dict]) -> str:
        """格式化字段列表"""
        if not fields:
            return "无需填写字段"
        
        formatted = []
        for field in fields[:10]:  # 只显示前10个
            if 'placeholder_name' in field:
                formatted.append(f"  - ${{{field['placeholder_name']}}}: {field.get('context', '')[:100]}")
            else:
                formatted.append(f"  - [空白]: {field.get('context', '')[:100]}")
        
        return '\n'.join(formatted)
    
    def _format_data(self, data: Dict) -> str:
        """格式化数据"""
        if not data:
            return "无额外数据"
        
        lines = []
        for key, value in data.items():
            if isinstance(value, dict):
                lines.append(f"{key}:")
                for k, v in value.items():
                    lines.append(f"  {k}: {v}")
            else:
                lines.append(f"{key}: {value}")
        
        return '\n'.join(lines)
    
    def _parse_llm_response(self, response: str) -> Dict[str, Any]:
        """
        解析LLM响应
        
        :param response: LLM的原始响应
        :return: 解析后的更新列表
        """
        try:
            # 尝试从response中提取JSON
            import re
            
            # 查找JSON块
            json_match = re.search(r'\{[\s\S]*\}', response)
            if json_match:
                json_str = json_match.group()
                result = json.loads(json_str)
                
                # 标准化响应格式
                updates = []
                for item in result.get('xml_updates', []):
                    updates.append({
                        'old_text': item.get('old_content', ''),
                        'new_text': item.get('new_content', ''),
                        'xpath': item.get('xpath', '')
                    })
                
                return {
                    'updates': updates,
                    'filled_fields': result.get('fields_filled', []),
                    'unfilled_fields': result.get('unfilled_fields', [])
                }
        
        except json.JSONDecodeError:
            pass
        
        # 如果无法解析JSON，尝试从文本中提取
        print(f"  ⚠️  无法解析JSON响应，尝试文本提取...")
        return {
            'updates': [],
            'filled_fields': [],
            'unfilled_fields': []
        }


def process_all_pages_with_llm(input_dir: str = 'split_pages',
                               llm_config_file: str = 'llm_config.json',
                               data_file: str = 'fill_data.json',
                               template_name: str = 'tender_form') -> Dict[str, Any]:
    """
    处理所有页面
    
    :param input_dir: 页面文件目录
    :param llm_config_file: LLM配置文件
    :param data_file: 填充数据文件
    :param template_name: 使用的模板
    :return: 处理结果统计
    """
    
    # 加载LLM配置
    from llm_connector import load_config_from_file
    
    try:
        llm_config = load_config_from_file(llm_config_file)
    except FileNotFoundError:
        print(f"❌ 未找到LLM配置文件: {llm_config_file}")
        print("请先创建配置文件，使用: setup_llm_config.py")
        return None
    
    # 加载数据
    fill_data = {}
    if os.path.exists(data_file):
        with open(data_file, 'r', encoding='utf-8') as f:
            fill_data = json.load(f)
    
    # 初始化处理器
    llm = LLMConnector(llm_config)
    processor = LLMPageProcessor(llm)
    
    # 获取所有页面文件
    page_files = sorted(glob.glob(os.path.join(input_dir, 'page_*.xml')),
                       key=lambda x: int(x.split('page_')[1].split('.')[0]))
    
    print(f"\n开始处理 {len(page_files)} 个页面...")
    print("=" * 60)
    
    results = {
        'total_pages': len(page_files),
        'processed': 0,
        'successful': 0,
        'failed': 0,
        'page_results': []
    }
    
    for page_file in page_files:
        page_num = int(page_file.split('page_')[1].split('.')[0])
        
        print(f"\n【第{page_num}页】")
        
        try:
            # 分析页面
            analyzer = XMLPageAnalyzer(page_file)
            page_info = analyzer.get_page_info()
            
            # 准备数据上下文
            data_context = {
                'page_title': f'第{page_num}页',
                'data': fill_data.get(f'page_{page_num}', {})
            }
            
            # 调用LLM处理
            result = processor.process_page_with_llm(
                page_num,
                page_info,
                data_context,
                template_name
            )
            
            results['page_results'].append(result)
            
            if result['status'] == 'success':
                # 应用更新
                if result['updates']:
                    analyzer.apply_updates(result['updates'])
                    analyzer.save()
                    print(f"  ✓ 已应用{len(result['updates'])}处修改")
                    results['successful'] += 1
                else:
                    print(f"  ℹ️  页面无需修改")
            else:
                results['failed'] += 1
                print(f"  ❌ 处理失败: {result.get('error', 'Unknown error')}")
            
            results['processed'] += 1
        
        except Exception as e:
            print(f"  ❌ 页面处理异常: {e}")
            results['failed'] += 1
            results['processed'] += 1
    
    print("\n" + "=" * 60)
    print(f"\n处理完成！")
    print(f"  总页数: {results['total_pages']}")
    print(f"  成功: {results['successful']}")
    print(f"  失败: {results['failed']}")
    
    return results


if __name__ == '__main__':
    process_all_pages_with_llm()
