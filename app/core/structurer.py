"""
Agent 2: LLM 结构化
将 PPT 原始文本 + 图片描述转化为结构化的患者信息 JSON
"""

import json
import logging

from app.config import settings
from app.services.llm_client import get_sync_client, safe_json_parse
from app.services.rules_loader import get_valid_drug_indications

logger = logging.getLogger("MedAudit")


def transform_raw_text_to_structured(raw_text: str, image_desc_text: str = "") -> dict:
    """
    利用LLM将PPT文本+图片描述转化为结构化的患者信息JSON

    Args:
        raw_text: PPT 提取的原始文本
        image_desc_text: 图片识别描述（可选）

    Returns:
        结构化患者数据字典
    """
    combined_text = raw_text
    if image_desc_text:
        combined_text += "\n\n【图片识别内容】:\n" + image_desc_text

    drug_indications = get_valid_drug_indications()
    constraint_str = ""
    if drug_indications:
        drug_list = list(drug_indications.keys())
        # 构建候选适应症说明（过滤掉 default/unknown 这种非医学适应症）
        indication_options = {}
        for drug, inds in drug_indications.items():
            real_inds = [i for i in inds if i not in ("default", "unknown")]
            indication_options[drug] = real_inds if real_inds else inds
        constraint_str = f"""
    【重要约束】：
    - "treatment_drug" 只能填药品通用名，必须从以下列表中选择：{json.dumps(drug_list, ensure_ascii=False)}。如不在列表中，填实际药品名。
    - "indication" 填该病例实际对应的医学适应症，从该药品的候选适应症中选择：
      {json.dumps(indication_options, ensure_ascii=False, indent=6)}
      注意：default 不是医学适应症，不要使用。如确实无法匹配，留空字符串 ""。
    - 注意：treatment_drug 只填药品名（如"西甲硅油"），不要包含适应症信息。
    """

    prompt = f"""
    你是一个精通临床医学数据处理的AI。请阅读以下从病例幻灯片(PPT)中提取出来的原始文本及图片识别内容，将其清洗并提取为标准的结构化数据。
    {constraint_str}
    【原始文本】:
    {combined_text}

    【输出格式要求】:
    必须只输出一个合法的 JSON 对象，不要包含任何 Markdown 标记（如 ```json）。字段要求如下：
    {{
        "patient_id": "患者编号或匿名化代号",
        "gender": "男/女/不详",
        "age": "年龄数字",
        "admission_date": "入院日期/就诊日期(YYYY-MM-DD)",
        "chief_complaint": "主诉",
        "present_illness": "现病史",
        "past_history": "既往史/个人史",
        "diagnosis": "初步诊断/最终诊断",
        "treatment_drug": "药品通用名（仅填药品名，如：西甲硅油、美泊利珠单抗、氟替美维）",
        "treatment_detail": "具体用法用量及治疗过程描述",
        "follow_up": "随访结果、检查指标及转归描述",
        "indication": "该病例实际对应的医学适应症（如：功能性消化不良、肠易激综合征、SEA、COPD、CRSwNP、EGPA、哮喘等，无法匹配则留空）",
        "privacy_issues": "如果图片或文本中发现患者个人信息泄露（姓名、住院号、医生信息等），在此列出；没有则填'无'"
    }}
    """
    logger.debug("调用 LLM 进行结构化数据提取")
    client = get_sync_client()
    messages = [{"role": "user", "content": prompt}]
    response = client.chat.completions.create(
        model=settings.openai_model_id,
        messages=messages,
        temperature=0.1
    )
    content = response.choices[0].message.content
    logger.debug(f"LLM 结构化原始输出长度: {len(content)} 字符")
    result = safe_json_parse(
        content,
        retries=2,
        retry_messages=messages,
        retry_prompt="你上次输出的不是合法JSON，请严格只输出JSON对象，不要有任何其他文字。",
        client=client,
    )

    # 后处理：清理 treatment_drug 字段，去掉混入的适应症信息
    if "treatment_drug" in result:
        drug = result["treatment_drug"]
        if " - " in drug:
            result["treatment_drug"] = drug.split(" - ")[0].strip()
        if "(" in result["treatment_drug"]:
            result["treatment_drug"] = result["treatment_drug"].split("(")[0].strip()

    # 后处理：indication 不应是药品名
    if "indication" in result:
        ind = result["indication"]
        if ind == result.get("treatment_drug"):
            result["indication"] = "unknown"

    return result
