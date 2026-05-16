"""
celery_app.py — Celery application factory shared by both the API (for .send_task)
and the worker (for task discovery). Kept in a separate module to avoid circular imports.
"""

import os
from celery import Celery

REDIS_URL: str = os.environ.get("REDIS_URL", "redis://localhost:6379/0")

celery_app = Celery(
    "learner_pdc",
    broker=REDIS_URL,
    backend=REDIS_URL,
    include=["tasks"],  # module in the worker image that defines the tasks
)

celery_app.conf.update(
    task_serializer="json",
    result_serializer="json",
    accept_content=["json"],
    timezone="Asia/Kolkata",
    enable_utc=True,
    # Keep results for 24 hours
    result_expires=86400,
    # Worker concurrency controlled by env
    worker_prefetch_multiplier=1,
    task_acks_late=True,
)
