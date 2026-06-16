"""
Agent 4: HTML 报告生成
使用 Jinja2 渲染审核报告
"""

import os
import logging

from jinja2 import Environment, FileSystemLoader, Template

from app.config import settings

logger = logging.getLogger("MedAudit")

# Jinja2 环境：从 templates 目录加载
_template_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), "templates")
_jinja_env = Environment(loader=FileSystemLoader(_template_dir))


def generate_report(patient_data: dict, audit_results: list, output_path: str):
    """渲染单例审核报告"""
    template = _jinja_env.get_template("single_report.html")
    html_out = template.render(patient=patient_data, audit_results=audit_results)

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html_out)
    logger.info(f"  审核报告已生成: {output_path}")


def generate_combined_report(cases: list, output_path: str):
    """渲染合并审核报告"""
    template = _jinja_env.get_template("combined_report.html")
    html_out = template.render(cases=cases)

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html_out)
    logger.info(f"合并审核报告已生成: {output_path}")


def render_report_html(patient_data: dict, audit_results: list) -> str:
    """渲染单例报告并返回 HTML 字符串（不写文件）"""
    template = _jinja_env.get_template("single_report.html")
    return template.render(patient=patient_data, audit_results=audit_results)


def render_combined_report_html(cases: list) -> str:
    """渲染合并报告并返回 HTML 字符串（不写文件）"""
    template = _jinja_env.get_template("combined_report.html")
    return template.render(cases=cases)
