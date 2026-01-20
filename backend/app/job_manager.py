from __future__ import annotations

import os
import tempfile
import threading
import time
import uuid
from dataclasses import dataclass, field
from typing import Callable, Dict, List, Optional

from .models import DocumentType

JobStatus = str  # pending, running, completed, failed


@dataclass
class JobRecord:
    job_id: str
    file_path: str
    file_name: str
    content_type: str
    document_type_hint: Optional[DocumentType]
    date_format: str
    start_date: Optional[str] = None  # 読取開始日（YYYY-MM-DD形式）
    end_date: Optional[str] = None  # 読取終了日（YYYY-MM-DD形式、未指定=最新まで）
    status: JobStatus = "pending"
    stage: str = "queued"
    detail: Optional[str] = None
    document_type: Optional[DocumentType] = None
    result_files: Optional[Dict[str, str]] = None
    partial_files: Optional[Dict[str, str]] = None
    processed_chunks: int = 0
    total_chunks: int = 0
    assets_payload: Optional[List[dict]] = None
    transactions_payload: Optional[List[dict]] = None
    processing_mode: str = "gemini"
    gemini_model: Optional[str] = None
    created_at: float = field(default_factory=time.time)
    updated_at: float = field(default_factory=time.time)


class JobHandle:
    def __init__(self, manager: "JobManager", job_id: str):
        self._manager = manager
        self.job_id = job_id

    def update(self, **fields) -> None:
        self._manager._update(self.job_id, **fields)


ProcessorFunc = Callable[[JobRecord, JobHandle], None]


class JobManager:
    def __init__(self, processor: ProcessorFunc) -> None:
        self._processor = processor
        self._jobs: Dict[str, JobRecord] = {}
        self._lock = threading.Lock()

    def submit(
        self,
        payload: bytes,
        content_type: str,
        file_name: str,
        document_type_hint: Optional[DocumentType],
        date_format: str,
        *,
        processing_mode: str = "gemini",
        gemini_model: Optional[str] = None,
        start_date: Optional[str] = None,
        end_date: Optional[str] = None,
    ) -> JobRecord:
        temp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
        temp.write(payload)
        temp.flush()
        temp.close()
        job_id = uuid.uuid4().hex
        record = JobRecord(
            job_id=job_id,
            file_path=temp.name,
            file_name=file_name,
            content_type=content_type,
            document_type_hint=document_type_hint,
            date_format=date_format,
            start_date=start_date,
            end_date=end_date,
            processing_mode=processing_mode,
            gemini_model=gemini_model,
        )
        with self._lock:
            self._jobs[job_id] = record
        thread = threading.Thread(target=self._run_job, args=(job_id,), daemon=True)
        thread.start()
        return record

    def get(self, job_id: str) -> Optional[JobRecord]:
        with self._lock:
            return self._jobs.get(job_id)

    def _run_job(self, job_id: str) -> None:
        record = self.get(job_id)
        if not record:
            return
        handle = JobHandle(self, job_id)
        handle.update(status="running", stage="queued")
        try:
            self._processor(record, handle)
        except Exception as exc:  # noqa: BLE001
            handle.update(status="failed", stage="failed", detail=str(exc))
        finally:
            if record.file_path and os.path.exists(record.file_path):
                os.remove(record.file_path)

    def _update(self, job_id: str, **fields) -> None:
        with self._lock:
            job = self._jobs.get(job_id)
            if not job:
                return
            for key, value in fields.items():
                if hasattr(job, key):
                    setattr(job, key, value)
            job.updated_at = time.time()
