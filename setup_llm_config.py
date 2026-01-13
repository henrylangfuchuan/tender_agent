"""
LLM 配置设置向导
帮助用户配置LLM连接信息
"""
import json
import os
from llm_connector import LLMProvider, LLMConfig, save_config_to_file


def setup_llm_config():
    """交互式配置LLM"""
    
    print("\n" + "=" * 60)
    print("  LLM 配置设置向导")
    print("=" * 60)
    
    print("\n请选择LLM服务商:")
    providers = [
        ("1", "OpenAI (ChatGPT)", LLMProvider.OPENAI),
        ("2", "Anthropic Claude", LLMProvider.CLAUDE),
        ("3", "阿里通义千问", LLMProvider.QWEN),
        ("4", "智谱清言", LLMProvider.ZHIPU),
        ("5", "自定义API", LLMProvider.CUSTOM),
    ]
    
    for code, name, _ in providers:
        print(f"  {code}. {name}")
    
    choice = input("\n请选择 (1-5): ").strip()
    
    provider = None
    for code, name, prov in providers:
        if choice == code:
            provider = prov
            break
    
    if provider is None:
        print("❌ 选择无效")
        return
    
    print(f"\n已选择: {choice}")
    
    # 配置API信息
    print("\n" + "-" * 60)
    print("请输入API配置信息")
    print("-" * 60)
    
    api_key = input("\n请输入 API Key: ").strip()
    if not api_key:
        print("❌ API Key不能为空")
        return
    
    api_url = input("\n请输入 API URL\n(例: https://api.openai.com/v1): ").strip()
    if not api_url:
        print("❌ API URL不能为空")
        return
    
    model = input("\n请输入 Model 名称\n(例: gpt-4, claude-3-opus): ").strip()
    if not model:
        print("❌ Model不能为空")
        return
    
    # 可选配置
    print("\n" + "-" * 60)
    print("高级配置 (留空使用默认值)")
    print("-" * 60)
    
    temperature_str = input("\n温度 (0.0-1.0, 默认0.3): ").strip()
    temperature = 0.3
    if temperature_str:
        try:
            temperature = float(temperature_str)
        except ValueError:
            print("⚠️  温度值无效，使用默认值 0.3")
    
    max_tokens_str = input("最大令牌数 (默认2000): ").strip()
    max_tokens = 2000
    if max_tokens_str:
        try:
            max_tokens = int(max_tokens_str)
        except ValueError:
            print("⚠️  令牌数无效，使用默认值 2000")
    
    # 创建配置
    config = LLMConfig(
        provider=provider,
        api_key=api_key,
        api_url=api_url,
        model=model,
        temperature=temperature,
        max_tokens=max_tokens
    )
    
    # 显示配置摘要
    print("\n" + "=" * 60)
    print("  配置摘要")
    print("=" * 60)
    print(f"服务商: {provider.value}")
    print(f"Model: {model}")
    print(f"API URL: {api_url}")
    print(f"温度: {temperature}")
    print(f"最大令牌数: {max_tokens}")
    
    # 确认保存
    confirm = input("\n确认保存配置? (y/n): ").strip().lower()
    if confirm == 'y':
        save_config_to_file(config, 'llm_config.json')
        print("\n✓ 配置已保存到 llm_config.json")
        return config
    else:
        print("\n❌ 已取消")
        return None


def create_sample_data_file():
    """创建示例数据文件"""
    
    sample_data = {
        "page_1": {
            "project_name": "中铝集团2026年采购项目",
            "package_number": "PKG-2026-001",
            "budget": "500万元"
        },
        "page_2": {
            "bidder_name": "某某公司",
            "bidder_address": "北京市朝阳区",
            "legal_representative": "张三",
            "contact_phone": "010-12345678"
        },
        "page_3": {
            "service_content": "提供项目管理咨询服务",
            "service_term": "12个月",
            "start_date": "2026-03-01",
            "end_date": "2027-02-28"
        }
    }
    
    with open('fill_data.json', 'w', encoding='utf-8') as f:
        json.dump(sample_data, f, ensure_ascii=False, indent=2)
    
    print("✓ 示例数据文件已创建: fill_data.json")


def create_sample_config():
    """创建示例配置文件"""
    
    sample_config = {
        "provider": "openai",
        "api_key": "sk-your-api-key-here",
        "api_url": "https://api.openai.com/v1",
        "model": "gpt-4",
        "temperature": 0.3,
        "max_tokens": 2000,
        "timeout": 60
    }
    
    with open('llm_config_sample.json', 'w', encoding='utf-8') as f:
        json.dump(sample_config, f, ensure_ascii=False, indent=2)
    
    print("✓ 示例配置文件已创建: llm_config_sample.json")


if __name__ == '__main__':
    import sys
    
    if len(sys.argv) > 1 and sys.argv[1] == 'sample':
        print("创建示例文件...")
        create_sample_data_file()
        create_sample_config()
    else:
        print("\n欢迎使用 Word 智能填写系统 LLM 配置向导\n")
        
        if os.path.exists('llm_config.json'):
            use_existing = input("检测到已有配置文件。是否重新配置? (y/n): ").strip().lower()
            if use_existing != 'y':
                print("使用现有配置")
                with open('llm_config.json', 'r') as f:
                    config = json.load(f)
                    print(f"\n当前配置 - 服务商: {config['provider']}, Model: {config['model']}")
                sys.exit(0)
        
        config = setup_llm_config()
        if config:
            print("\n✓ 配置成功！")
            print("下一步: 编辑 fill_data.json 准备填充数据，然后运行智能填写脚本")
