"""Shared registry so routers can access the active JobManager instance."""
from __future__ import annotations

from typing import Optional

from .job_manager import JobManager

job_manager: Optional[JobManager] = None

__all__ = ["job_manager"]
