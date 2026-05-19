"""
MongoDB clients and collection accessors (async Motor + sync PyMongo).

Motor is imported lazily so the Celery worker image (PyMongo only) can use
the sync helpers without installing motor.
"""

from __future__ import annotations

import os
from typing import TYPE_CHECKING, Any

from pymongo import MongoClient

if TYPE_CHECKING:
    from motor.motor_asyncio import AsyncIOMotorClient

MONGO_URL: str = os.environ.get("MONGO_URL", "mongodb://localhost:27017")
DB_NAME: str = os.environ.get("MONGO_DB", "learner_pdc")

_async_client: Any = None
_sync_client: MongoClient | None = None


def get_async_client() -> AsyncIOMotorClient:
    global _async_client
    if _async_client is None:
        from motor.motor_asyncio import AsyncIOMotorClient

        _async_client = AsyncIOMotorClient(MONGO_URL)
    return _async_client


def get_database():
    return get_async_client()[DB_NAME]


def get_faculty_collection():
    return get_database()["faculty"]


def get_audit_logs_collection():
    return get_database()["audit_logs"]


def get_tasks_collection():
    return get_database()["tasks"]


async def close_async_client():
    global _async_client
    if _async_client is not None:
        _async_client.close()
        _async_client = None


def get_sync_db():
    global _sync_client
    if _sync_client is None:
        _sync_client = MongoClient(MONGO_URL)
    return _sync_client[DB_NAME]


def get_sync_faculty_collection():
    return get_sync_db()["faculty"]


def get_sync_audit_logs_collection():
    return get_sync_db()["audit_logs"]


def get_sync_tasks_collection():
    return get_sync_db()["tasks"]
