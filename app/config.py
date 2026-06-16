"""
MedAudit Agent 配置管理
使用 pydantic-settings 从环境变量 / .env 文件加载配置
"""

import os
from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    """应用全局配置"""

    # --- LLM API ---
    openai_base_url: str = "https://token-plan-cn.xiaomimimo.com/v1"
    openai_api_key: str = ""
    openai_model_id: str = "mimo-v2.5-pro"
    openai_multimodal_model: str = ""  # 为空则跳过图片识别

    # --- 路径 ---
    upload_dir: str = "data/uploads"
    output_dir: str = "data/output_reports"
    rules_config_path: str = "config/rules_config.json"
    log_dir: str = "logs"

    # --- 并发 ---
    max_workers: int = 3

    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        extra="ignore",
    )

    @property
    def vision_enabled(self) -> bool:
        return bool(self.openai_multimodal_model)


# 全局单例
settings = Settings()
