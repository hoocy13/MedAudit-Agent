"""
MedAudit Agent - FastAPI 应用入口
医学病例合规自动审核系统
"""

import os
import logging
from contextlib import asynccontextmanager

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from app.config import settings
from app.api.audit import router as audit_router
from app.services.rules_loader import load_all_rules


# ==========================================
# 日志配置
# ==========================================
def setup_logging():
    """配置日志：控制台 INFO + 文件 DEBUG"""
    os.makedirs(settings.log_dir, exist_ok=True)

    logger = logging.getLogger("MedAudit")
    logger.setLevel(logging.DEBUG)

    # 控制台 handler
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(logging.Formatter("%(message)s"))
    logger.addHandler(console_handler)

    # 文件 handler
    file_handler = logging.FileHandler(
        os.path.join(settings.log_dir, "medaudit.log"), encoding="utf-8"
    )
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(
        logging.Formatter("%(asctime)s | %(levelname)-7s | %(message)s", datefmt="%Y-%m-%d %H:%M:%S")
    )
    logger.addHandler(file_handler)

    return logger


# ==========================================
# 应用生命周期
# ==========================================
@asynccontextmanager
async def lifespan(app: FastAPI):
    """启动时初始化，关闭时清理"""
    logger = setup_logging()
    logger.info("=" * 60)
    logger.info("  MedAudit Agent - 医学病例合规自动审核系统")
    logger.info("  FastAPI 服务启动中...")
    logger.info("=" * 60)
    logger.info(f"  模型: {settings.openai_model_id}")
    logger.info(f"  视觉: {'启用 (' + settings.openai_multimodal_model + ')' if settings.vision_enabled else '未启用'}")
    logger.info(f"  上传目录: {settings.upload_dir}")
    logger.info(f"  输出目录: {settings.output_dir}")

    # 预加载规则库
    load_all_rules()

    # 确保目录存在
    os.makedirs(settings.upload_dir, exist_ok=True)
    os.makedirs(settings.output_dir, exist_ok=True)

    logger.info("-" * 60)
    yield
    logger.info("服务关闭")


# ==========================================
# FastAPI 应用
# ==========================================
app = FastAPI(
    title="MedAudit Agent",
    description="医学病例合规自动审核系统 - 上传PPT病例文件，自动进行多维度合规审核",
    version="1.0.0",
    lifespan=lifespan,
)

# CORS 中间件
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 注册路由
app.include_router(audit_router)


# ==========================================
# 根端点
# ==========================================
@app.get("/", tags=["health"])
async def root():
    """健康检查"""
    return {
        "service": "MedAudit Agent",
        "status": "running",
        "version": "1.0.0",
        "docs": "/docs",
    }


@app.get("/health", tags=["health"])
async def health():
    """详细健康信息"""
    return {
        "status": "healthy",
        "model": settings.openai_model_id,
        "vision_enabled": settings.vision_enabled,
    }
