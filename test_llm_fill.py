# -*- coding: utf-8 -*-
"""
Test script for LLM intelligent filling of page2 and page8
"""

import json
import os
import sys
import shutil
from datetime import datetime

# Import custom modules
from llm_connector import LLMConnector, load_config_from_file
from prompt_library import PromptLibrary
from process_with_llm import XMLPageAnalyzer, LLMPageProcessor
from merge_pages import merge_pages
from xml_to_docx import xml_to_docx

# Test data
TEST_DATA = {
    "investorName": "Shanghai Stars Technology Co., Ltd.",
    "registeredAddress": "No. 88 Keyuan Road, Zhang Jiang High-Tech Park, Pudong, Shanghai",
    "postalCode": "201203",
    "contact": {
        "name": "Li Ming",
        "phone": "021-56789012",
        "fax": "021-56789013",
        "website": "www.stars-tech.cn"
    },
    "legalRepresentative": {
        "name": "Wang Qiang",
        "title": "Senior Engineer",
        "phone": "13800138000"
    },
    "technicalDirector": {
        "name": "Zhao Li",
        "title": "Senior Engineer",
        "phone": "13900139000"
    },
    "qualifications": {
        "type": "ISO9001",
        "level": "Grade 1",
        "certificateNo": "QMS-20240123"
    },
    "registrationNo": "91310000MA1K5R7U2B",
    "totalEmployees": 120,
    "registeredCapital": "RMB 50 million",
    "seniorStaffCount": 15,
    "foundingDate": "2010-06-15",
    "midLevelStaffCount": 30,
    "bankName": "Industrial and Commercial Bank of China Shanghai Branch",
    "technicalStaffCount": 50,
    "bankAccountNo": "6222001234567890",
    "registeredEngineers": "3 Registered Safety Engineers, 2 Registered Structural Engineers",
    "businessScope": "Software development, technical consulting, technical services, system integration",
    "affiliatedCompanies": "Control relationship with Shanghai Stars Investment Co., Ltd.",
    "remarks": "None"
}

GLOBAL_DATA = {
    "projectName": "Test Project",
    "packageNo": "xxx11",
    "bidder": "xxxx Co., Ltd.",
    "date": "2026-01-13"
}

def print_header(text):
    """Print header"""
    print("\n" + "="*60)
    print(f"  {text}")
    print("="*60)

def print_step(step_num, text):
    """Print step"""
    print(f"\n[Step {step_num}] {text}")

def main():
    print_header("LLM Intelligent Fill Test - Page 2 & Page 8")
    
    # Step 1: Check LLM configuration
    print_step(1, "Check LLM configuration")
    config_file = 'llm_config.json'
    
    if not os.path.exists(config_file):
        print(f"WARNING: {config_file} not found")
        print("Please run: python setup_llm_config.py")
        return
    
    try:
        llm_config = load_config_from_file(config_file)
        print(f"OK Loaded config: {llm_config.provider.value} - {llm_config.model}")
    except Exception as e:
        print(f"ERROR Failed to load config: {e}")
        return
    
    # Step 2: Read and analyze page 2
    print_step(2, "Read and analyze page_2.xml")
    
    if not os.path.exists('split_pages/page_2.xml'):
        print("ERROR: split_pages/page_2.xml not found")
        return
    
    try:
        analyzer_page2 = XMLPageAnalyzer('split_pages/page_2.xml')
        text_content_page2 = analyzer_page2.extract_text_content()
        print(f"OK Text content (first 200 chars): {text_content_page2[:200]}...")
    except Exception as e:
        print(f"ERROR: {e}")
        return
    
    # Step 3: Read and analyze page 8
    print_step(3, "Read and analyze page_8.xml")
    
    if not os.path.exists('split_pages/page_8.xml'):
        print("ERROR: split_pages/page_8.xml not found")
        return
    
    try:
        analyzer_page8 = XMLPageAnalyzer('split_pages/page_8.xml')
        text_content_page8 = analyzer_page8.extract_text_content()
        print(f"OK Text content (first 200 chars): {text_content_page8[:200]}...")
    except Exception as e:
        print(f"ERROR: {e}")
        return
    
    # Step 4: Show data to be filled
    print_step(4, "Show data to be filled")
    
    print("\nGlobal data:")
    for key, value in GLOBAL_DATA.items():
        print(f"  {key}: {value}")
    
    print("\nEnterprise details (first 5 items):")
    for i, (key, value) in enumerate(TEST_DATA.items()):
        if i >= 5:
            print(f"  ... total {len(TEST_DATA)} items")
            break
        print(f"  {key}: {value}")
    
    # Step 5: Fill page 2 with LLM
    print_step(5, "Fill page_2.xml with LLM")
    
    # Read raw XML content for both pages
    with open('split_pages/page_2.xml', 'r', encoding='utf-8') as f:
        page_2_xml_content = f.read()
    
    with open('split_pages/page_8.xml', 'r', encoding='utf-8') as f:
        page_8_xml_content = f.read()
    
    # Prepare data for LLM
    data_str = json.dumps(TEST_DATA, ensure_ascii=False, indent=2)
    global_data_str = json.dumps(GLOBAL_DATA, ensure_ascii=False, indent=2)
    
    if llm_config.api_key == 'sk-your-api-key-here' or not llm_config.api_key:
        print("WARNING: No valid API key configured, skipping actual fill")
        print("Please configure valid API key in llm_config.json")
        print("Will use original XML files for merge")
        page_2_xml_modified = page_2_xml_content
        page_8_xml_modified = page_8_xml_content
    else:
        try:
            print("LOADING: Calling LLM API (may take 10-30 seconds)...")
            
            connector = LLMConnector(llm_config)
            
            # Process page 2
            print("Processing page_2...")
            prompt_page2 = f"""You are an XML content expert. Your task is to intelligently fill the XML content based on provided data.

Below is the XML content of page 2 (Business Section):
<XML_CONTENT>
{page_2_xml_content}
</XML_CONTENT>

Below is the data you should use to fill the XML:
<PROVIDED_DATA>
{data_str}
</PROVIDED_DATA>

Global data:
<GLOBAL_DATA>
{global_data_str}
</GLOBAL_DATA>

Instructions:
1. Analyze the XML structure and identify fields that need to be filled
2. Fill the fields with appropriate data from PROVIDED_DATA
3. Replace placeholder text (like {{placeholder_name}}) with actual values
4. Maintain the exact XML structure and formatting
5. Return ONLY the modified XML content, nothing else

Return the complete modified XML:"""

            response_page2 = connector.call(prompt_page2, system_prompt="You are an expert at filling XML templates with precise data. Return only valid XML content.")
            print(response_page2)
            if response_page2 and response_page2.strip().startswith('<'):
                # Extract XML from response
                page_2_xml_modified = response_page2.strip()
                print("OK page_2 filled successfully")
            else:
                print("WARNING: LLM response is not valid XML, using original")
                page_2_xml_modified = page_2_xml_content
            
            # Process page 8
            print_step(6, "Fill page_8.xml with LLM")
            print("Processing page_8...")
            
            prompt_page8 = f"""You are an XML content expert. Your task is to intelligently fill the XML content based on provided data.

Below is the XML content of page 8 (Enterprise Information Table):
<XML_CONTENT>
{page_8_xml_content}
</XML_CONTENT>

Below is the data you should use to fill the XML:
<PROVIDED_DATA>
{data_str}
</PROVIDED_DATA>

Global data:
<GLOBAL_DATA>
{global_data_str}
</GLOBAL_DATA>

Instructions:
1. Analyze the XML structure and identify fields that need to be filled
2. Fill the table cells and form fields with appropriate data from PROVIDED_DATA
3. Replace placeholder text (like {{placeholder_name}}) with actual values
4. Maintain the exact XML structure and formatting
5. Return ONLY the modified XML content, nothing else

Return the complete modified XML:"""

            response_page8 = connector.call(prompt_page8, system_prompt="You are an expert at filling XML templates with precise data. Return only valid XML content.")
            
            if response_page8 and response_page8.strip().startswith('<'):
                # Extract XML from response
                page_8_xml_modified = response_page8.strip()
                print("OK page_8 filled successfully")
            else:
                print("WARNING: LLM response is not valid XML, using original")
                page_8_xml_modified = page_8_xml_content
        
        except Exception as e:
            print(f"ERROR: LLM call failed: {e}")
            import traceback
            traceback.print_exc()
            print("Will use original XML files for merge")
            page_2_xml_modified = page_2_xml_content
            page_8_xml_modified = page_8_xml_content
    
    # Step 7: Merge page 2 and page 8
    print_step(7, "Merge page_2 and page_8")
    
    try:
        # Create temporary directory
        test_split_dir = 'test_split_pages'
        os.makedirs(test_split_dir, exist_ok=True)
        
        # Write modified XML content to temporary directory
        with open(f'{test_split_dir}/page_1.xml', 'w', encoding='utf-8') as f:
            f.write(page_2_xml_modified)
        
        with open(f'{test_split_dir}/page_2.xml', 'w', encoding='utf-8') as f:
            f.write(page_8_xml_modified)
        
        print(f"OK Prepared modified files for merging:")
        print(f"  {test_split_dir}/page_1.xml (modified page_2)")
        print(f"  {test_split_dir}/page_2.xml (modified page_8)")
        
        # Merge
        output_xml = 'test_merged_document.xml'
        merge_pages(test_split_dir, output_xml)
        print(f"OK Merged to: {output_xml}")
        
        # Step 8: Convert to Word
        print_step(8, "Convert to Word document")
        output_docx = 'test_output_page2_page8.docx'
        
        xml_to_docx(output_xml, output_docx, 'template.docx')
        print(f"OK Converted to: {output_docx}")
        
        # Clean up temporary files
        shutil.rmtree(test_split_dir)
        
        print_header("TEST COMPLETED SUCCESSFULLY!")
        print(f"\nOutput file: {output_docx}")
        print(f"This file contains merged content from page_2 and page_8")
        
    except Exception as e:
        print(f"ERROR: Merge or conversion failed: {e}")
        import traceback
        traceback.print_exc()
        return

if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nWARNING: Interrupted by user")
        sys.exit(1)
    except Exception as e:
        print(f"\nERROR: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
