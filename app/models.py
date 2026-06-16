"""
MedAudit Agent 数据模型
定义 API 请求/响应的 Pydantic 模型
"""

from enum import Enum
from typing import Optional

from pydantic import BaseModel, Field, field_validator


# ==========================================
# 枚举
# ==========================================

class TaskStatus(str, Enum):
    """任务状态"""
    pending = "pending"
    processing = "processing"
    completed = "completed"
    failed = "failed"


class AuditStatus(str, Enum):
    """审核结论"""
    Pass = "Pass"
    Warning = "Warning"
    Fail = "Fail"


# ==========================================
# 患者数据
# ==========================================

class PatientData(BaseModel):
    """结构化患者数据"""
    patient_id: str = ""
    gender: str = ""
    age: str = ""
    admission_date: str = ""
    chief_complaint: str = ""
    present_illness: str = ""
    past_history: str = ""
    diagnosis: str = ""
    treatment_drug: str = ""
    treatment_detail: str = ""
    follow_up: str = ""
    indication: str = ""
    privacy_issues: str = ""

    @field_validator("age", mode="before")
    @classmethod
    def coerce_age_to_str(cls, v):
        """LLM 可能返回 int 类型的 age，统一转为 str"""
        return str(v) if v is not None else ""


# ==========================================
# 审核结果
# ==========================================

class AuditItem(BaseModel):
    """单条审核意见"""
    aspect: str = Field(description="审核维度")
    status: AuditStatus = Field(description="审核结论")
    detail: str = Field(description="具体审查判定说明")
    suggestion: str = Field(description="修改建议")


# ==========================================
# 任务结果
# ==========================================

class AuditResult(BaseModel):
    """单个文件的完整审核结果"""
    file: str = Field(description="原始文件名")
    drug: str = ""
    indication: str = ""
    images: int = 0
    audit_count: int = 0
    pass_count: int = 0
    warning_count: int = 0
    fail_count: int = 0
    patient_data: Optional[PatientData] = None
    audit_results: list[AuditItem] = []
    error: Optional[str] = None


# ==========================================
# 任务
# ==========================================

class TaskInfo(BaseModel):
    """任务信息"""
    task_id: str
    status: TaskStatus = TaskStatus.pending
    filenames: list[str] = []
    created_at: str = ""
    started_at: Optional[str] = None
    completed_at: Optional[str] = None
    error: Optional[str] = None


class TaskResult(BaseModel):
    """任务完整结果（含审核详情）"""
    task_id: str
    status: TaskStatus
    filenames: list[str] = []
    created_at: str = ""
    started_at: Optional[str] = None
    completed_at: Optional[str] = None
    elapsed_seconds: Optional[float] = None
    results: list[AuditResult] = []
    total_pass: int = 0
    total_warning: int = 0
    total_fail: int = 0
    error: Optional[str] = None
    report_url: Optional[str] = None


# ==========================================
# API 响应
# ==========================================

class UploadResponse(BaseModel):
    """上传响应"""
    task_id: str
    status: TaskStatus
    message: str
    filenames: list[str] = []


class TaskListResponse(BaseModel):
    """任务列表响应"""
    tasks: list[TaskInfo] = []
