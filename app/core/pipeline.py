"""
审核流水线
串联 4 个 Agent 的完整处理流程
"""

import os
import time
import logging

from app.config import settings
from app.core.extractor import (
    extract_text_from_pptx,
    extract_images_from_pptx,
    describe_images_with_llm,
)
from app.core.structurer import transform_raw_text_to_structured
from app.core.auditor import audit_patient_data
from app.services.rules_loader import match_rules

logger = logging.getLogger("MedAudit")


def run_audit_pipeline(pptx_path: str, output_dir: str) -> dict:
    """
    处理单个 PPT 文件的完整审核流程

    Args:
        pptx_path: PPT 文件路径
        output_dir: 输出目录（用于保存图片等中间产物）

    Returns:
        审核结果字典
    """
    filename = os.path.splitext(os.path.basename(pptx_path))[0]
    short_name = filename[:50] + ".." if len(filename) > 50 else filename

    logger.info(f"▶ 开始处理: {short_name}")

    # 第一步：提取文本
    logger.info("  ├─ [1/4] 解析PPT文本...")
    raw_text = extract_text_from_pptx(pptx_path)

    # 第二步：提取图片并识别
    logger.info("  ├─ [2/4] 提取图片...")
    image_info_list, img_log = extract_images_from_pptx(pptx_path, output_dir)
    logger.info(img_log)
    image_desc = ""
    if image_info_list:
        if settings.vision_enabled:
            logger.info(f"  │  调用视觉模型识别...")
            image_desc = describe_images_with_llm(image_info_list)
            logger.info("  │  图片识别完成")
    else:
        logger.info("  │  未发现嵌入图片")

    # 第三步：结构化
    logger.info("  ├─ [3/4] LLM结构化患者数据...")
    patient_json = transform_raw_text_to_structured(raw_text, image_desc)

    drug = patient_json['treatment_drug']
    indication = patient_json.get('indication', '未识别')
    logger.info(f"  │  用药: {drug} | 适应症: {indication}")

    if patient_json.get('privacy_issues') and patient_json['privacy_issues'] != '无':
        logger.info(f"  │  ⚠ 隐私风险: {patient_json['privacy_issues'][:60]}...")

    # 第四步：审核
    logger.info("  └─ [4/4] 合规审核...")
    rules = match_rules(drug, indication)
    audit_results = audit_patient_data(patient_json, rules)

    pass_n = sum(1 for r in audit_results if r['status'] == 'Pass')
    warn_n = sum(1 for r in audit_results if r['status'] == 'Warning')
    fail_n = sum(1 for r in audit_results if r['status'] == 'Fail')
    logger.info(f"     结果: ✓{pass_n} Pass | ⚠{warn_n} Warning | ✗{fail_n} Fail")

    return {
        "file": filename,
        "drug": drug,
        "indication": indication,
        "images": len(image_info_list),
        "audit_count": len(audit_results),
        "pass_count": pass_n,
        "warning_count": warn_n,
        "fail_count": fail_n,
        "patient_data": patient_json,
        "audit_results": audit_results,
    }
