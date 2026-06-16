"""
规则库加载与匹配服务
启动时加载 rules_config.json，提供药品/适应症的模糊匹配
"""

import os
import json
import logging

from app.config import settings

logger = logging.getLogger("MedAudit")

# 缓存：启动时加载一次
_all_rules: dict | None = None
_drug_indications: dict | None = None


def load_all_rules() -> dict:
    """加载全部规则（带缓存）"""
    global _all_rules
    if _all_rules is not None:
        return _all_rules

    rules_file = settings.rules_config_path
    if not os.path.exists(rules_file):
        logger.warning(f"规则配置文件不存在: {rules_file}")
        _all_rules = {}
        return _all_rules

    with open(rules_file, "r", encoding="utf-8") as f:
        _all_rules = json.load(f)

    logger.info(f"已加载规则库: {len(_all_rules)} 个药品")
    return _all_rules


def get_valid_drug_indications() -> dict:
    """获取所有有效的 药品→适应症 映射（带缓存）"""
    global _drug_indications
    if _drug_indications is not None:
        return _drug_indications

    all_rules = load_all_rules()
    _drug_indications = {}
    for drug, indications in all_rules.items():
        _drug_indications[drug] = list(indications.keys())
    return _drug_indications


def match_rules(drug_name: str, indication: str = None) -> dict | None:
    """
    根据药品名和适应症模糊匹配审核规则

    Args:
        drug_name: 药品通用名
        indication: 适应症

    Returns:
        匹配到的规则集，未匹配返回 None
    """
    all_rules = load_all_rules()

    # 模糊匹配药品
    matched_drug = None
    for key in all_rules.keys():
        if key in drug_name or drug_name in key:
            matched_drug = all_rules[key]
            logger.debug(f"规则匹配: 药品 '{drug_name}' → '{key}'")
            break

    if not matched_drug:
        logger.debug(f"未找到药品 '{drug_name}' 的审核规则")
        return None

    # 模糊匹配适应症
    if indication:
        for ind_key in matched_drug.keys():
            if ind_key in indication or indication in ind_key:
                logger.debug(f"规则匹配: 适应症 '{indication}' → '{ind_key}'")
                return matched_drug[ind_key]

    logger.debug(f"未匹配到适应症 '{indication}' 的专属规则，使用通用审核")
    return None


def reload_rules():
    """强制重新加载规则（用于热更新）"""
    global _all_rules, _drug_indications
    _all_rules = None
    _drug_indications = None
    load_all_rules()
    get_valid_drug_indications()
    logger.info("规则库已重新加载")
