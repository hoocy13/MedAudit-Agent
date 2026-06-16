"""
MedAudit Agent 启动入口

启动方式:
    python main.py                  # 默认 0.0.0.0:8000
    python main.py --port 9000      # 自定义端口
    uvicorn app.main:app --reload   # 开发模式（自动重载）
"""

import uvicorn

if __name__ == "__main__":
    uvicorn.run(
        "app.main:app",
        host="0.0.0.0",
        port=8000,
        reload=True,
        log_level="info",
    )
