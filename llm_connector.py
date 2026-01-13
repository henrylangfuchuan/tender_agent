"""
LLM 连接配置模块
支持 OpenAI、Claude、通义千问等多个LLM服务
"""
import os
import json
import requests
from typing import Dict, List, Optional
from dataclasses import dataclass
from enum import Enum

class LLMProvider(Enum):
    """支持的LLM服务商"""
    OPENAI = "openai"
    CLAUDE = "claude"
    QWEN = "qwen"  # 阿里通义千问
    ZHIPU = "zhipu"  # 智谱清言
    CUSTOM = "custom"  # 自定义API


@dataclass
class LLMConfig:
    """LLM配置"""
    provider: LLMProvider
    api_key: str
    api_url: str
    model: str
    temperature: float = 0.3
    max_tokens: int = 2000
    timeout: int = 60
    
    def to_dict(self) -> dict:
        return {
            'provider': self.provider.value,
            'api_key': self.api_key,
            'api_url': self.api_url,
            'model': self.model,
            'temperature': self.temperature,
            'max_tokens': self.max_tokens,
        }


class LLMConnector:
    """LLM 连接器 - 统一调用接口"""
    
    def __init__(self, config: LLMConfig):
        self.config = config
        self.provider = config.provider
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Word-LLM-Processor/1.0'
        })
    
    def call(self, prompt: str, system_prompt: Optional[str] = None) -> str:
        """
        调用LLM API
        
        :param prompt: 用户提示词
        :param system_prompt: 系统提示词
        :return: LLM的响应文本
        """
        if self.provider == LLMProvider.OPENAI:
            return self._call_openai(prompt, system_prompt)
        elif self.provider == LLMProvider.CLAUDE:
            return self._call_claude(prompt, system_prompt)
        elif self.provider == LLMProvider.QWEN:
            return self._call_qwen(prompt, system_prompt)
        elif self.provider == LLMProvider.ZHIPU:
            return self._call_zhipu(prompt, system_prompt)
        elif self.provider == LLMProvider.CUSTOM:
            return self._call_custom(prompt, system_prompt)
        else:
            raise ValueError(f"不支持的LLM提供商: {self.provider}")
    
    def _call_openai(self, prompt: str, system_prompt: Optional[str] = None) -> str:
        """调用OpenAI API"""
        messages = []
        
        if system_prompt:
            messages.append({
                "role": "system",
                "content": system_prompt
            })
        
        messages.append({
            "role": "user",
            "content": prompt
        })
        
        try:
            response = self.session.post(
                f"{self.config.api_url}/chat/completions",
                json={
                    "model": self.config.model,
                    "messages": messages,
                    "temperature": self.config.temperature,
                    "max_tokens": self.config.max_tokens,
                },
                headers={
                    "Authorization": f"Bearer {self.config.api_key}",
                    "Content-Type": "application/json"
                },
                timeout=self.config.timeout
            )
            response.raise_for_status()
            result = response.json()
            return result['choices'][0]['message']['content']
        
        except requests.exceptions.RequestException as e:
            raise RuntimeError(f"OpenAI API 调用失败: {e}")
    
    def _call_claude(self, prompt: str, system_prompt: Optional[str] = None) -> str:
        """调用Claude API"""
        try:
            response = self.session.post(
                f"{self.config.api_url}/messages",
                json={
                    "model": self.config.model,
                    "max_tokens": self.config.max_tokens,
                    "system": system_prompt or "",
                    "messages": [
                        {
                            "role": "user",
                            "content": prompt
                        }
                    ]
                },
                headers={
                    "x-api-key": self.config.api_key,
                    "anthropic-version": "2023-06-01",
                    "Content-Type": "application/json"
                },
                timeout=self.config.timeout
            )
            response.raise_for_status()
            result = response.json()
            return result['content'][0]['text']
        
        except requests.exceptions.RequestException as e:
            raise RuntimeError(f"Claude API 调用失败: {e}")
    
    def _call_qwen(self, prompt: str, system_prompt: Optional[str] = None) -> str:
        """调用阿里通义千问API"""
        try:
            response = self.session.post(
                self.config.api_url,
                json={
                    "model": self.config.model,
                    "messages": [
                        {
                            "role": "system",
                            "content": system_prompt or "你是一个有用的助手。"
                        },
                        {
                            "role": "user",
                            "content": prompt
                        }
                    ],
                    "temperature": self.config.temperature,
                    "max_tokens": self.config.max_tokens,
                },
                headers={
                    "Authorization": f"Bearer {self.config.api_key}",
                    "Content-Type": "application/json"
                },
                timeout=self.config.timeout
            )
            response.raise_for_status()
            result = response.json()
            return result['output']['choices'][0]['message']['content']
        
        except requests.exceptions.RequestException as e:
            raise RuntimeError(f"通义千问API 调用失败: {e}")
    
    def _call_zhipu(self, prompt: str, system_prompt: Optional[str] = None) -> str:
        """调用智谱清言API"""
        try:
            response = self.session.post(
                f"{self.config.api_url}/openai/v1/chat/completions",
                json={
                    "model": self.config.model,
                    "messages": [
                        {
                            "role": "system",
                            "content": system_prompt or ""
                        },
                        {
                            "role": "user",
                            "content": prompt
                        }
                    ],
                    "temperature": self.config.temperature,
                    "max_tokens": self.config.max_tokens,
                },
                headers={
                    "Authorization": f"Bearer {self.config.api_key}",
                    "Content-Type": "application/json"
                },
                timeout=self.config.timeout
            )
            response.raise_for_status()
            result = response.json()
            return result['choices'][0]['message']['content']
        
        except requests.exceptions.RequestException as e:
            raise RuntimeError(f"智谱清言API 调用失败: {e}")
    
    def _call_custom(self, prompt: str, system_prompt: Optional[str] = None) -> str:
        """调用腾讯云 DeepSeek V3 自定义API"""
        try:
            messages = []
            # system_prompt 官方示例没有 system，但是你可以自己加成 user 消息或者 role=system
            if system_prompt:
                messages.append({"Role": "system", "Content": system_prompt})
            messages.append({"Role": "user", "Content": prompt})

            payload = {
                "MaxTokens": self.config.max_tokens,
                "Temperature": self.config.temperature,
                "Model": self.config.model,
                "Messages": messages,
                "Stream": False   # 流式输出设 False，方便直接返回
            }

            response = self.session.post(
                self.config.api_url,
                json=payload,
                headers={
                    "Authorization": f"Bearer {self.config.api_key}",
                    "Content-Type": "application/json",
                    "X-TC-Action": "ChatCompletions"   # 官方要求
                },
                timeout=self.config.timeout
            )
            response.raise_for_status()
            result = response.json()

            # 官方返回通常在 result.Choices[0].Message.Content
            if "Choices" in result and len(result["Choices"]) > 0:
                choice = result["Choices"][0]
                if "Message" in choice and "Content" in choice["Message"]:
                    return choice["Message"]["Content"]

            # fallback
            return str(result)

        except requests.exceptions.RequestException as e:
            raise RuntimeError(f"自定义API 调用失败: {e}")


def load_config_from_file(config_file: str = 'llm_config.json') -> LLMConfig:
    """
    从配置文件加载LLM配置
    
    :param config_file: 配置文件路径
    :return: LLMConfig对象
    """
    if not os.path.exists(config_file):
        raise FileNotFoundError(f"配置文件不存在: {config_file}")
    
    with open(config_file, 'r', encoding='utf-8') as f:
        config_data = json.load(f)
    
    return LLMConfig(
        provider=LLMProvider(config_data['provider']),
        api_key=config_data['api_key'],
        api_url=config_data['api_url'],
        model=config_data['model'],
        temperature=config_data.get('temperature', 0.3),
        max_tokens=config_data.get('max_tokens', 2000),
        timeout=config_data.get('timeout', 60)
    )


def save_config_to_file(config: LLMConfig, config_file: str = 'llm_config.json'):
    """
    保存LLM配置到文件
    
    :param config: LLMConfig对象
    :param config_file: 配置文件路径
    """
    config_data = config.to_dict()
    
    with open(config_file, 'w', encoding='utf-8') as f:
        json.dump(config_data, f, ensure_ascii=False, indent=2)
    
    print(f"✓ 配置已保存到: {config_file}")


if __name__ == '__main__':
    # 使用示例
    print("LLM 连接模块加载成功")
    print("\n支持的提供商:")
    for provider in LLMProvider:
        print(f"  - {provider.value}")
    
    # ===== 自定义大模型测试示例 =====
    print("\n=== 自定义大模型连接测试 ===")
    
    try:
        # 从配置文件加载
        config = load_config_from_file('llm_config.json')
        
        # 仅测试自定义提供商
        if config.provider != LLMProvider.CUSTOM:
            print("⚠ 当前配置不是自定义API，将临时切换为 CUSTOM 测试")
            config.provider = LLMProvider.CUSTOM
        
        connector = LLMConnector(config)
        
        # 测试 prompt
        test_prompt = "请用一句话介绍人工智能的应用。"
        system_prompt = "你是一个知识渊博、回答简洁的助手。"
        
        print("请求中，请稍候...")
        result = connector.call(prompt=test_prompt, system_prompt=system_prompt)
        
        print("\n✅ 自定义API返回结果:")
        print(result)
    
    except Exception as e:
        print(f"\n❌ 测试自定义API失败: {e}")
