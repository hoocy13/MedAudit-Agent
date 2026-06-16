"""
Agent 3: 合规审核
对照规则库，让 LLM 输出结构化审核意见
"""

import json
import logging

from app.config import settings
from app.services.llm_client import get_sync_client, safe_json_parse

logger = logging.getLogger("MedAudit")


def audit_patient_data(patient_data: dict, rules: dict | None) -> list:
    """
    对照特殊规则，让大模型输出结构化审核意见

    Args:
        patient_data: 结构化患者数据
        rules: 匹配到的审核规则（可为 None）

    Returns:
        审核意见列表
    """
    if not rules:
        rules_str = "无特殊药品专属规则，仅按照通用规范审核：1. 保护患者隐私；2. 逻辑故事线通顺；3. 药物需用通用名。"
        logger.info("    无专属规则，使用通用规范审核")
    else:
        rules_str = json.dumps(rules, ensure_ascii=False, indent=2)
        logger.debug(f"    加载审核规则: {len(rules_str)} 字符")

    prompt = f"""
    你是一名资深的药企医学合规经理。请严格对照【审核规则】，对【患者结构化数据】进行逐条审查，指出不合规或存在安全隐患的地方。

    【患者结构化数据】:
    {json.dumps(patient_data, ensure_ascii=False, indent=2)}

    【审核规则】:
    {rules_str}

    【输出格式要求】:
    请仅返回一个标准的 JSON 数组，不要包裹 ```json 标记。数组中每个对象代表一个审核项。

    **重要：必须包含以下【基础质控检查】的4个固定审核维度，放在数组最前面：**

    1. **文本规范性检查** - 检查病例文本中是否存在错别字、语法错误、标点符号误用等问题。如发现错别字（如"的地得"混用、医学术语拼写错误等），给出具体位置和修正建议。
    2. **通用名规范检查** - 检查药品名称是否使用了通用名（INN），而非商品名或品牌名。如发现使用商品名（如"舒利迭"应为"沙美特罗替卡松"、"信必可"应为"布地奈德福莫特罗"等），需标注并建议替换为通用名。
    3. **医学逻辑审核** - 检查药物与适应症的医学逻辑是否合理，包括：药物作用机制与疾病病理是否匹配、剂量是否在合理范围、用药疗程是否符合临床实践、联合用药是否存在相互作用风险等。
    4. **单位规范检查** - 检查医学数值的单位是否正确规范，包括：剂量单位（mg/mg/mL）、给药频率单位（qd/bid/tid/q4w等）、检验指标单位（%、×10^9/L等）、体重/身高等单位。如发现单位缺失、错误或不规范，需标注并给出正确写法。

    然后再根据【审核规则】进行其他维度的审核。

    格式如下：
    [
        {{
            "aspect": "文本规范性检查",
            "status": "Pass 或 Warning 或 Fail",
            "detail": "具体的审查判定说明",
            "suggestion": "修改建议或'无需修改'"
        }},
        {{
            "aspect": "通用名规范检查",
            "status": "Pass 或 Warning 或 Fail",
            "detail": "具体的审查判定说明",
            "suggestion": "修改建议或'无需修改'"
        }},
        {{
            "aspect": "医学逻辑审核",
            "status": "Pass 或 Warning 或 Fail",
            "detail": "具体的审查判定说明",
            "suggestion": "修改建议或'无需修改'"
        }},
        {{
            "aspect": "单位规范检查",
            "status": "Pass 或 Warning 或 Fail",
            "detail": "具体的审查判定说明",
            "suggestion": "修改建议或'无需修改'"
        }},
        {{
            "aspect": "其他审核维度(如: 适应症合规/隐私保护/Storyline故事线/安全性风险/科学幻灯)",
            "status": "Pass 或 Warning 或 Fail",
            "detail": "具体的审查判定说明，引用病例中的具体事实",
            "suggestion": "修改建议或'无需修改'"
        }}
    ]
    """
    logger.debug("调用 LLM 进行合规审核")
    client = get_sync_client()
    messages = [{"role": "user", "content": prompt}]
    response = client.chat.completions.create(
        model=settings.openai_model_id,
        messages=messages,
        temperature=0.2
    )
    content = response.choices[0].message.content
    logger.debug(f"LLM 审核原始输出长度: {len(content)} 字符")
    return safe_json_parse(
        content,
        retries=2,
        retry_messages=messages,
        retry_prompt="你上次输出的不是合法JSON数组，请严格只输出JSON数组，不要有任何其他文字。",
        client=client,
    )
