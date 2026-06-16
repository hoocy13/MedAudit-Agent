"""
审核相关 API 路由
提供文件上传、批量审核、任务查询、报告下载等端点
"""

import os
import uuid
import shutil
import logging
from datetime import datetime

from fastapi import APIRouter, UploadFile, File, HTTPException, BackgroundTasks
from fastapi.responses import HTMLResponse

from app.config import settings
from app.models import (
    TaskStatus, TaskInfo, TaskResult, AuditResult,
    UploadResponse, TaskListResponse, AuditItem, PatientData,
)
from app.core.pipeline import run_audit_pipeline
from app.core.reporter import generate_combined_report, render_combined_report_html

logger = logging.getLogger("MedAudit")

router = APIRouter(prefix="/api/audit", tags=["audit"])

# ==========================================
# 内存任务存储（第一步不引入数据库）
# ==========================================
_tasks: dict[str, TaskResult] = {}


def _process_task(task_id: str, file_paths: list[str]):
    """后台任务：执行审核流水线"""
    task = _tasks[task_id]
    task.status = TaskStatus.processing
    task.started_at = datetime.now().isoformat()

    try:
        results = []
        for fpath in file_paths:
            result = run_audit_pipeline(fpath, settings.output_dir)
            results.append(AuditResult(**result))

        # 统计
        total_pass = sum(r.pass_count for r in results)
        total_warn = sum(r.warning_count for r in results)
        total_fail = sum(r.fail_count for r in results)

        task.results = results
        task.total_pass = total_pass
        task.total_warning = total_warn
        task.total_fail = total_fail
        task.status = TaskStatus.completed
        task.completed_at = datetime.now().isoformat()

        # 计算耗时
        if task.started_at:
            start = datetime.fromisoformat(task.started_at)
            task.elapsed_seconds = round((datetime.now() - start).total_seconds(), 1)

        # 生成合并报告
        if results:
            cases = [
                {
                    "filename": r.file,
                    "patient": r.patient_data.model_dump() if r.patient_data else {},
                    "audit_results": [a.model_dump() for a in r.audit_results],
                    "pass_count": r.pass_count,
                    "warning_count": r.warning_count,
                    "fail_count": r.fail_count,
                }
                for r in results
            ]
            report_dir = os.path.join(settings.output_dir, task_id)
            os.makedirs(report_dir, exist_ok=True)
            report_path = os.path.join(report_dir, "audit_report_combined.html")
            generate_combined_report(cases, report_path)
            task.report_url = f"/api/audit/{task_id}/report"
            logger.info(f"任务 {task_id} 报告已生成: {report_path}")

    except Exception as e:
        logger.error(f"任务 {task_id} 处理失败: {e}", exc_info=True)
        task.status = TaskStatus.failed
        task.error = str(e)
        task.completed_at = datetime.now().isoformat()


# ==========================================
# 端点
# ==========================================

@router.post("/upload", response_model=UploadResponse, summary="上传单个PPT文件并发起审核")
async def upload_single(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(..., description="PPT文件 (.pptx)"),
):
    """上传单个 PPT 文件，后台异步执行审核，立即返回 task_id"""
    if not file.filename.endswith(".pptx"):
        raise HTTPException(status_code=400, detail="仅支持 .pptx 格式文件")

    # 创建任务
    task_id = str(uuid.uuid4())[:8]
    upload_dir = os.path.join(settings.upload_dir, task_id)
    os.makedirs(upload_dir, exist_ok=True)

    # 保存文件
    file_path = os.path.join(upload_dir, file.filename)
    with open(file_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    # 初始化任务
    now = datetime.now().isoformat()
    _tasks[task_id] = TaskResult(
        task_id=task_id,
        status=TaskStatus.pending,
        filenames=[file.filename],
        created_at=now,
    )

    # 后台执行
    background_tasks.add_task(_process_task, task_id, [file_path])

    return UploadResponse(
        task_id=task_id,
        status=TaskStatus.pending,
        message=f"文件 '{file.filename}' 已上传，审核任务已创建",
        filenames=[file.filename],
    )


@router.post("/batch", response_model=UploadResponse, summary="上传多个PPT文件并发起批量审核")
async def upload_batch(
    background_tasks: BackgroundTasks,
    files: list[UploadFile] = File(..., description="多个PPT文件"),
):
    """上传多个 PPT 文件，后台异步执行批量审核"""
    pptx_files = [f for f in files if f.filename.endswith(".pptx")]
    if not pptx_files:
        raise HTTPException(status_code=400, detail="至少需要一个 .pptx 格式文件")

    task_id = str(uuid.uuid4())[:8]
    upload_dir = os.path.join(settings.upload_dir, task_id)
    os.makedirs(upload_dir, exist_ok=True)

    file_paths = []
    filenames = []
    for f in pptx_files:
        fpath = os.path.join(upload_dir, f.filename)
        with open(fpath, "wb") as out:
            shutil.copyfileobj(f.file, out)
        file_paths.append(fpath)
        filenames.append(f.filename)

    now = datetime.now().isoformat()
    _tasks[task_id] = TaskResult(
        task_id=task_id,
        status=TaskStatus.pending,
        filenames=filenames,
        created_at=now,
    )

    background_tasks.add_task(_process_task, task_id, file_paths)

    return UploadResponse(
        task_id=task_id,
        status=TaskStatus.pending,
        message=f"已上传 {len(filenames)} 个文件，批量审核任务已创建",
        filenames=filenames,
    )


@router.get("/tasks", response_model=TaskListResponse, summary="列出所有任务")
async def list_tasks():
    """返回所有任务的摘要列表"""
    tasks = [
        TaskInfo(
            task_id=t.task_id,
            status=t.status,
            filenames=t.filenames,
            created_at=t.created_at,
            started_at=t.started_at,
            completed_at=t.completed_at,
            error=t.error,
        )
        for t in _tasks.values()
    ]
    return TaskListResponse(tasks=tasks)


@router.get("/{task_id}", response_model=TaskResult, summary="查询任务状态和结果")
async def get_task(task_id: str):
    """根据 task_id 查询任务状态，完成后返回完整审核结果"""
    if task_id not in _tasks:
        raise HTTPException(status_code=404, detail=f"任务 {task_id} 不存在")
    return _tasks[task_id]


@router.get("/{task_id}/report", response_class=HTMLResponse, summary="下载HTML审核报告")
async def get_report(task_id: str):
    """下载指定任务的合并审核报告（HTML格式）"""
    if task_id not in _tasks:
        raise HTTPException(status_code=404, detail=f"任务 {task_id} 不存在")

    task = _tasks[task_id]
    if task.status != TaskStatus.completed:
        raise HTTPException(status_code=400, detail=f"任务尚未完成，当前状态: {task.status.value}")

    # 尝试从文件读取
    report_path = os.path.join(settings.output_dir, task_id, "audit_report_combined.html")
    if os.path.exists(report_path):
        with open(report_path, "r", encoding="utf-8") as f:
            return HTMLResponse(content=f.read())

    # 回退：实时渲染
    if task.results:
        cases = [
            {
                "filename": r.file,
                "patient": r.patient_data.model_dump() if r.patient_data else {},
                "audit_results": [a.model_dump() for a in r.audit_results],
                "pass_count": r.pass_count,
                "warning_count": r.warning_count,
                "fail_count": r.fail_count,
            }
            for r in task.results
        ]
        html = render_combined_report_html(cases)
        return HTMLResponse(content=html)

    raise HTTPException(status_code=404, detail="报告未生成")
