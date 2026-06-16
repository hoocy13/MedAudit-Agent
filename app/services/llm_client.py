"""
LLM 客户端封装
提供同步/异步的 OpenAI 兼容 API 调用，含 JSON 安全解析与重试
"""

import re
import json
import logging

from openai import OpenAI, AsyncOpenAI

from app.config import settings

logger = logging.getLogger("MedAudit")


# ==========================================
# 客户端初始化
# ==========================================

def get_sync_client() -> OpenAI:
    """获取同步 OpenAI 客户端"""
    return OpenAI(
        base_url=settings.openai_base_url,
        api_key=settings.openai_api_key,
    )


def get_async_client() -> AsyncOpenAI:
    """获取异步 OpenAI 客户端"""
    return AsyncOpenAI(
        base_url=settings.openai_base_url,
        api_key=settings.openai_api_key,
    )


# ==========================================
# JSON 安全解析
# ==========================================

def clean_json_string(text: str) -> str:
    """清理 LLM 输出中常见的 JSON 格式问题"""
    text = text.strip()
    text = re.sub(r'```json\s*', '', text)
    text = re.sub(r'```\s*$', '', text)
    text = re.sub(r'//[^\n]*', '', text)
    text = re.sub(r',\s*([}\]])', r'\1', text)
    return text.strip()


def safe_json_parse(text: str, retries: int = 2, retry_messages=None, retry_prompt=None,
                    client: OpenAI = None) -> dict | list:
    """
    安全解析 JSON，失败时清理重试；仍失败则重调 LLM

    Args:
        text: LLM 原始输出
        retries: 重试次数
        retry_messages: 重调 LLM 时的消息列表
        retry_prompt: 重调 LLM 时的追加提示
        client: 同步 OpenAI 客户端（重调 LLM 时需要）
    """
    if client is None:
        client = get_sync_client()

    cleaned = clean_json_string(text)
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        logger.debug("首次 JSON 解析失败，清理后重试")

    for attempt in range(retries):
        try:
            return json.loads(cleaned)
        except json.JSONDecodeError:
            if attempt < retries - 1:
                match = re.search(r'(\[.*\]|\{.*\})', cleaned, re.DOTALL)
                if match:
                    cleaned = match.group(1)
                    logger.debug(f"提取 JSON 片段，第 {attempt + 1} 次重试解析")
                    continue
            if retry_messages and retry_prompt:
                logger.warning(f"JSON 解析失败，第 {attempt + 1} 次重调 LLM")
                retry_messages_copy = retry_messages + [
                    {"role": "user", "content": retry_prompt}
                ]
                response = client.chat.completions.create(
                    model=settings.openai_model_id,
                    messages=retry_messages_copy,
                    temperature=0.1
                )
                text = response.choices[0].message.content
                cleaned = clean_json_string(text)
            else:
                raise ValueError(f"JSON 解析失败，已重试 {retries} 次。原始内容: {text[:200]}...")

    raise ValueError(f"JSON 解析失败，已重试 {retries} 次。原始内容: {text[:200]}...")


async def async_safe_json_parse(text: str, retries: int = 2, retry_messages=None,
                                 retry_prompt=None, client: AsyncOpenAI = None) -> dict | list:
    """异步版本的 JSON 安全解析"""
    if client is None:
        client = get_async_client()

    cleaned = clean_json_string(text)
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        logger.debug("首次 JSON 解析失败，清理后重试")

    for attempt in range(retries):
        try:
            return json.loads(cleaned)
        except json.JSONDecodeError:
            if attempt < retries - 1:
                match = re.search(r'(\[.*\]|\{.*\})', cleaned, re.DOTALL)
                if match:
                    cleaned = match.group(1)
                    logger.debug(f"提取 JSON 片段，第 {attempt + 1} 次重试解析")
                    continue
            if retry_messages and retry_prompt:
                logger.warning(f"JSON 解析失败，第 {attempt + 1} 次重调 LLM")
                retry_messages_copy = retry_messages + [
                    {"role": "user", "content": retry_prompt}
                ]
                response = await client.chat.completions.create(
                    model=settings.openai_model_id,
                    messages=retry_messages_copy,
                    temperature=0.1
                )
                text = response.choices[0].message.content
                cleaned = clean_json_string(text)
            else:
                raise ValueError(f"JSON 解析失败，已重试 {retries} 次。原始内容: {text[:200]}...")

    raise ValueError(f"JSON 解析失败，已重试 {retries} 次。原始内容: {text[:200]}...")
