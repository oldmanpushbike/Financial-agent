# -*- coding: utf-8 -*-
"""LLM adapter — wraps Zhipu (OpenAI-compatible endpoint) for LangChain."""
from langchain_openai import ChatOpenAI

from agent.config import ZHIPU_API_KEY, ZHIPU_MODEL

# 智谱 BigModel 提供 OpenAI 兼容接口
ZHIPU_BASE_URL = "https://open.bigmodel.cn/api/paas/v4/"


def get_llm(temperature: float = 0.3, streaming: bool = True) -> ChatOpenAI:
    return ChatOpenAI(
        model=ZHIPU_MODEL,
        api_key=ZHIPU_API_KEY,
        base_url=ZHIPU_BASE_URL,
        temperature=temperature,
        streaming=streaming,
        max_retries=3,
    )
