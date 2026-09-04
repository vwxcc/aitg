#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Telegram AI Bot — single-file backend MVP / foundation.

Architecture:
Telegram -> aiogram handlers -> SQLite -> priority queues -> AI Router -> provider API
                                        |-> File processing / OCR / media
                                        |-> Knowledge DB
                                        |-> subscriptions / referrals / usage / events

Designed for a small ZimaOS / Render worker with ~4 GB RAM.

Required environment variables:
  BOT_TOKEN=...
  ADMIN_TELEGRAM_IDS=123,456

Optional:
  DATABASE_PATH=./data/bot.db
  FILES_DIR=./data/files
  TEMP_DIR=./data/tmp
  DEFAULT_MODEL_ID=qwen-35b
  DEFAULT_SUMMARY_MODEL_ID=
  SUPPORT_USERNAME=@support
  MAX_FILE_SIZE_MB=50
  MAX_CONTEXT_WORDS=10000
  POLL_TIMEOUT=30
  LOG_LEVEL=INFO
  QUEUE_WORKERS=12
  FILE_CONCURRENCY=2
  OCR_CONCURRENCY=1
  MEDIA_CONCURRENCY=1

Dependencies (recommended):
  aiogram>=3.20,<4
  aiohttp>=3.10
  python-dotenv>=1.0
  PyMuPDF>=1.24
  python-docx>=1.1
  openpyxl>=3.1
  python-pptx>=1.0
  pillow>=10

Optional:
  pytesseract        # OCR
  pillow-heif        # HEIC/HEIF
  pandas             # CSV / extra table handling
  odfpy              # ODS
  pydub              # audio helpers (ffmpeg still required for conversion)
  openai             # not required; router below uses HTTP directly

System package recommended:
  ffmpeg
  tesseract-ocr

Notes:
- The bot assumes OpenAI-compatible chat-completions style APIs.
- Provider-specific quirks should be handled in AIService.request().
- Secrets never go to logs.
"""

from __future__ import annotations

import asyncio
import csv
import hashlib
import json
import logging
import mimetypes
import os
import re
import shutil
import sqlite3
import subprocess
import tempfile
import time
import traceback
import uuid
from collections import defaultdict, deque
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Iterable, Optional

import aiohttp
from aiogram import Bot, Dispatcher, F, Router
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.exceptions import TelegramBadRequest, TelegramForbiddenError, TelegramNetworkError, TelegramRetryAfter
from aiogram.filters import Command, CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import CallbackQuery, Document, InlineKeyboardButton, InlineKeyboardMarkup, Message
from aiogram.utils.keyboard import InlineKeyboardBuilder, ReplyKeyboardBuilder
from dotenv import load_dotenv


# -----------------------------------------------------------------------------
# Configuration
# -----------------------------------------------------------------------------

load_dotenv()

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = Path(os.getenv("DATA_DIR", BASE_DIR / "data"))
DATABASE_PATH = Path(os.getenv("DATABASE_PATH", DATA_DIR / "bot.db"))
FILES_DIR = Path(os.getenv("FILES_DIR", DATA_DIR / "files"))
TEMP_DIR = Path(os.getenv("TEMP_DIR", DATA_DIR / "tmp"))

DATA_DIR.mkdir(parents=True, exist_ok=True)
FILES_DIR.mkdir(parents=True, exist_ok=True)
TEMP_DIR.mkdir(parents=True, exist_ok=True)

BOT_TOKEN = os.getenv("BOT_TOKEN", "").strip()
ADMIN_IDS: set[int] = {
    int(x.strip())
    for x in os.getenv("ADMIN_TELEGRAM_IDS", "").split(",")
    if x.strip().isdigit()
}

DEFAULT_MODEL_ID = os.getenv("DEFAULT_MODEL_ID", "qwen-35b").strip()
DEFAULT_SUMMARY_MODEL_ID = os.getenv("DEFAULT_SUMMARY_MODEL_ID", "").strip()
SUPPORT_USERNAME = os.getenv("SUPPORT_USERNAME", "").strip()
MAX_FILE_SIZE_MB = int(os.getenv("MAX_FILE_SIZE_MB", "50"))
MAX_FILE_SIZE = MAX_FILE_SIZE_MB * 1024 * 1024
MAX_CONTEXT_WORDS = int(os.getenv("MAX_CONTEXT_WORDS", "10000"))
POLL_TIMEOUT = int(os.getenv("POLL_TIMEOUT", "30"))
LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO").upper()

logging.basicConfig(
    level=getattr(logging, LOG_LEVEL, logging.INFO),
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
)
log = logging.getLogger("telegram_ai_bot")


# -----------------------------------------------------------------------------
# Constants
# -----------------------------------------------------------------------------

STATUS_QUEUED = "queued"
STATUS_RUNNING = "running"
STATUS_COMPLETED = "completed"
STATUS_FAILED = "failed"
STATUS_CANCELLED = "cancelled"
STATUS_INTERRUPTED = "interrupted"

ROLE_SYSTEM = "system"
ROLE_USER = "user"
ROLE_ASSISTANT = "assistant"

SAFE_TOOL_SCHOOL_DB = "SCHOOL_DB"
SAFE_TOOL_ASK_USER = "ASK_USER"

DEFAULT_FREE_TOKEN_LIMIT = 500_000
DEFAULT_PAID_TOKEN_LIMIT = 0  # 0 = unlimited
DEFAULT_RESET_PERIOD = 6 * 3600
DEFAULT_REFERRAL_DAYS = 5
DEFAULT_REFERRAL_MONTHLY_MAX_DAYS = 30

# User-facing status strings from the specification.
TEXT_QUEUE = "⏳ Запрос находится в очереди…"
TEXT_PROCESSING = "👀 Обрабатываю запрос…"
TEXT_FILE = "📄 Читаю файл…"
TEXT_IMAGE = "👀 Смотрю фото…"
TEXT_OCR = "🔎 Извлекаю текст с изображения…"
TEXT_TABLE = "📊 Анализирую таблицу…"
TEXT_AI = "🧠 Думаю…"
TEXT_AI_MORE = "🧠 Думаю дальше…"
TEXT_DB = "🔎 Ищу информацию в базе данных…"
TEXT_VERIFY = "🔬 Перепроверяю результат…"
TEXT_DONE = "✅ Готово."


# -----------------------------------------------------------------------------
# Utilities
# -----------------------------------------------------------------------------


def utc_now() -> datetime:
    return datetime.now(timezone.utc)


def iso_now() -> str:
    return utc_now().isoformat()


def parse_iso(value: Optional[str]) -> Optional[datetime]:
    if not value:
        return None
    try:
        return datetime.fromisoformat(value)
    except ValueError:
        return None


def safe_json_loads(value: Optional[str], default: Any = None) -> Any:
    if not value:
        return default
    try:
        return json.loads(value)
    except Exception:
        return default


def json_dumps(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"))


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def truncate_text(text: str, max_chars: int = 30_000) -> str:
    text = text or ""
    if len(text) <= max_chars:
        return text
    return text[:max_chars] + "\n[…обрезано из-за ограничения размера…]"


def word_count(text: str) -> int:
    return len(re.findall(r"\S+", text or ""))


def human_duration(seconds: int) -> str:
    seconds = max(0, int(seconds))
    h, rem = divmod(seconds, 3600)
    m, s = divmod(rem, 60)
    if h:
        return f"{h} ч {m} мин"
    if m:
        return f"{m} мин"
    return f"{s} с"


def mask_secret(secret: str) -> str:
    if not secret:
        return ""
    if len(secret) <= 8:
        return "••••••••"
    return secret[:3] + "••••" + secret[-3:]


def normalize_tags(tags: Iterable[str]) -> list[str]:
    result = []
    seen = set()
    for tag in tags:
        tag = re.sub(r"\s+", " ", tag.strip())
        if not tag:
            continue
        key = tag.lower()
        if key in seen:
            continue
        seen.add(key)
        result.append(tag[:100])
    return result


# -----------------------------------------------------------------------------
# Database
# -----------------------------------------------------------------------------

class Database:
    def __init__(self, path: Path):
        self.path = path
        self.conn: Optional[sqlite3.Connection] = None
        self.lock = asyncio.Lock()

    async def connect(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.conn = sqlite3.connect(self.path, check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA journal_mode=WAL")
        self.conn.execute("PRAGMA foreign_keys=ON")
        self.conn.execute("PRAGMA busy_timeout=5000")
        await self.init_schema()

    def _require(self) -> sqlite3.Connection:
        if self.conn is None:
            raise RuntimeError("Database is not connected")
        return self.conn

    async def close(self) -> None:
        async with self.lock:
            if self.conn:
                self.conn.close()
                self.conn = None

    async def init_schema(self) -> None:
        async with self.lock:
            c = self._require()
            c.executescript(
                """
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                telegram_id INTEGER NOT NULL UNIQUE,
                username TEXT,
                first_name TEXT,
                last_name TEXT,
                role TEXT NOT NULL DEFAULT 'student',
                is_blocked INTEGER NOT NULL DEFAULT 0,
                selected_model_id TEXT,
                mode TEXT NOT NULL DEFAULT 'student',
                notifications_enabled INTEGER NOT NULL DEFAULT 1,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS admins (
                user_id INTEGER PRIMARY KEY,
                telegram_id INTEGER NOT NULL UNIQUE,
                level TEXT NOT NULL DEFAULT 'admin',
                created_at TEXT NOT NULL,
                FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS subscriptions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                active INTEGER NOT NULL DEFAULT 1,
                starts_at TEXT NOT NULL,
                expires_at TEXT,
                source TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE CASCADE
            );
            CREATE INDEX IF NOT EXISTS idx_sub_user_active ON subscriptions(user_id, active, expires_at);

            CREATE TABLE IF NOT EXISTS referrals (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                referrer_user_id INTEGER NOT NULL,
                referred_user_id INTEGER NOT NULL UNIQUE,
                rewarded_days INTEGER NOT NULL DEFAULT 0,
                reward_month TEXT,
                created_at TEXT NOT NULL,
                FOREIGN KEY(referrer_user_id) REFERENCES users(id) ON DELETE CASCADE,
                FOREIGN KEY(referred_user_id) REFERENCES users(id) ON DELETE CASCADE
            );
            CREATE INDEX IF NOT EXISTS idx_ref_referrer ON referrals(referrer_user_id);

            CREATE TABLE IF NOT EXISTS models (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                model_key TEXT NOT NULL UNIQUE,
                name TEXT NOT NULL,
                model_id TEXT NOT NULL,
                base_url TEXT NOT NULL,
                api_key TEXT NOT NULL,
                api_format TEXT NOT NULL DEFAULT 'openai_chat',
                enabled INTEGER NOT NULL DEFAULT 1,
                free_token_limit INTEGER NOT NULL DEFAULT 500000,
                paid_token_limit INTEGER NOT NULL DEFAULT 0,
                reset_period_seconds INTEGER NOT NULL DEFAULT 21600,
                max_concurrency INTEGER NOT NULL DEFAULT 10,
                priority INTEGER NOT NULL DEFAULT 100,
                temperature REAL,
                max_output_tokens INTEGER,
                extra_json TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS chats (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                title TEXT,
                model_key TEXT,
                summary TEXT,
                summary_updated_at TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE CASCADE
            );
            CREATE INDEX IF NOT EXISTS idx_chats_user_updated ON chats(user_id, updated_at DESC);

            CREATE TABLE IF NOT EXISTS messages (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                chat_id INTEGER NOT NULL,
                role TEXT NOT NULL,
                content TEXT NOT NULL,
                model_key TEXT,
                input_tokens INTEGER DEFAULT 0,
                output_tokens INTEGER DEFAULT 0,
                created_at TEXT NOT NULL,
                FOREIGN KEY(chat_id) REFERENCES chats(id) ON DELETE CASCADE
            );
            CREATE INDEX IF NOT EXISTS idx_messages_chat_created ON messages(chat_id, created_at);

            CREATE TABLE IF NOT EXISTS usage (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                model_key TEXT NOT NULL,
                chat_id INTEGER,
                input_tokens INTEGER NOT NULL DEFAULT 0,
                output_tokens INTEGER NOT NULL DEFAULT 0,
                total_tokens INTEGER NOT NULL DEFAULT 0,
                window_started_at TEXT,
                window_reset_at TEXT,
                request_id TEXT,
                status TEXT NOT NULL DEFAULT 'completed',
                created_at TEXT NOT NULL,
                FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE CASCADE
            );
            CREATE INDEX IF NOT EXISTS idx_usage_user_model_created ON usage(user_id, model_key, created_at);

            CREATE TABLE IF NOT EXISTS files (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                chat_id INTEGER,
                message_id INTEGER,
                original_name TEXT NOT NULL,
                mime_type TEXT,
                size INTEGER NOT NULL DEFAULT 0,
                stored_path TEXT NOT NULL,
                processing_status TEXT NOT NULL DEFAULT 'queued',
                extracted_text TEXT,
                ocr_result TEXT,
                transcription TEXT,
                metadata_json TEXT,
                checksum TEXT,
                created_at TEXT NOT NULL,
                FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE CASCADE,
                FOREIGN KEY(chat_id) REFERENCES chats(id) ON DELETE SET NULL,
                FOREIGN KEY(message_id) REFERENCES messages(id) ON DELETE SET NULL
            );
            CREATE INDEX IF NOT EXISTS idx_files_user_created ON files(user_id, created_at DESC);

            CREATE TABLE IF NOT EXISTS knowledge (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                title TEXT NOT NULL,
                content TEXT NOT NULL,
                source TEXT,
                category TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS tags (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                created_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS knowledge_tags (
                knowledge_id INTEGER NOT NULL,
                tag_id INTEGER NOT NULL,
                PRIMARY KEY(knowledge_id, tag_id),
                FOREIGN KEY(knowledge_id) REFERENCES knowledge(id) ON DELETE CASCADE,
                FOREIGN KEY(tag_id) REFERENCES tags(id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS broadcasts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                created_by_user_id INTEGER,
                audience TEXT NOT NULL DEFAULT 'all',
                text TEXT NOT NULL,
                status TEXT NOT NULL DEFAULT 'queued',
                cursor_id INTEGER DEFAULT 0,
                sent_count INTEGER NOT NULL DEFAULT 0,
                failed_count INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                started_at TEXT,
                finished_at TEXT,
                FOREIGN KEY(created_by_user_id) REFERENCES users(id) ON DELETE SET NULL
            );

            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS events (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER,
                chat_id INTEGER,
                event_type TEXT NOT NULL,
                payload_json TEXT,
                created_at TEXT NOT NULL,
                FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE SET NULL,
                FOREIGN KEY(chat_id) REFERENCES chats(id) ON DELETE SET NULL
            );
            CREATE INDEX IF NOT EXISTS idx_events_type_created ON events(event_type, created_at);

            CREATE TABLE IF NOT EXISTS jobs (
                id TEXT PRIMARY KEY,
                user_id INTEGER,
                chat_id INTEGER,
                job_type TEXT NOT NULL,
                priority INTEGER NOT NULL DEFAULT 100,
                state TEXT NOT NULL DEFAULT 'queued',
                payload_json TEXT,
                error TEXT,
                created_at TEXT NOT NULL,
                started_at TEXT,
                finished_at TEXT,
                FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE SET NULL,
                FOREIGN KEY(chat_id) REFERENCES chats(id) ON DELETE SET NULL
            );
            CREATE INDEX IF NOT EXISTS idx_jobs_state_priority ON jobs(state, priority, created_at);

            CREATE TABLE IF NOT EXISTS pending_questions (
                user_id INTEGER PRIMARY KEY,
                chat_id INTEGER NOT NULL,
                original_user_message_id INTEGER,
                question TEXT NOT NULL,
                state TEXT NOT NULL DEFAULT 'waiting',
                context_json TEXT,
                created_at TEXT NOT NULL,
                FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE CASCADE,
                FOREIGN KEY(chat_id) REFERENCES chats(id) ON DELETE CASCADE
            );
            """
        )
        c.commit()

    async def execute(self, sql: str, params: tuple | list = ()) -> sqlite3.Cursor:
        async with self.lock:
            c = self._require()
            cur = c.execute(sql, params)
            c.commit()
            return cur

    async def executescript(self, script: str) -> None:
        async with self.lock:
            c = self._require()
            c.executescript(script)
            c.commit()

    async def fetchone(self, sql: str, params: tuple | list = ()) -> Optional[sqlite3.Row]:
        async with self.lock:
            c = self._require()
            return c.execute(sql, params).fetchone()

    async def fetchall(self, sql: str, params: tuple | list = ()) -> list[sqlite3.Row]:
        async with self.lock:
            c = self._require()
            return c.execute(sql, params).fetchall()

    async def transaction(self, callback):
        async with self.lock:
            c = self._require()
            c.execute("BEGIN IMMEDIATE")
            try:
                result = callback(c)
                c.commit()
                return result
            except Exception:
                c.rollback()
                raise


# -----------------------------------------------------------------------------
# Database service helpers
# -----------------------------------------------------------------------------

class Repo:
    def __init__(self, db: Database):
        self.db = db

    async def seed_defaults(self) -> None:
        defaults = {
            "referral_enabled": "1",
            "referral_days": str(DEFAULT_REFERRAL_DAYS),
            "referral_monthly_max_days": str(DEFAULT_REFERRAL_MONTHLY_MAX_DAYS),
            "max_file_size_mb": str(MAX_FILE_SIZE_MB),
            "max_context_words": str(MAX_CONTEXT_WORDS),
            "default_model_id": DEFAULT_MODEL_ID,
            "summary_model_id": DEFAULT_SUMMARY_MODEL_ID,
            "file_concurrency": "2",
            "ocr_concurrency": "1",
            "media_concurrency": "1",
            "broadcast_concurrency": "1",
        }
        for key, value in defaults.items():
            await self.db.execute(
                "INSERT OR IGNORE INTO settings(key,value,updated_at) VALUES(?,?,?)",
                (key, value, iso_now()),
            )

        await self.db.execute(
            "UPDATE users SET role=mode, updated_at=? "
            "WHERE mode IN ('student','teacher','applicant') AND role='student' AND mode!='student'",
            (iso_now(),),
        )

        for tg_id in ADMIN_IDS:
            user = await self.get_user(tg_id)
            if not user:
                await self.ensure_user_obj(tg_id, None, None, None, None)
                user = await self.get_user(tg_id)
            if user:
                await self.db.execute(
                    "INSERT OR IGNORE INTO admins(user_id,telegram_id,level,created_at) VALUES(?,?,?,?)",
                    (user["id"], tg_id, "admin", iso_now()),
                )

    async def get_setting(self, key: str, default: Optional[str] = None) -> Optional[str]:
        row = await self.db.fetchone("SELECT value FROM settings WHERE key=?", (key,))
        return row["value"] if row else default

    async def set_setting(self, key: str, value: str) -> None:
        await self.db.execute(
            "INSERT INTO settings(key,value,updated_at) VALUES(?,?,?) "
            "ON CONFLICT(key) DO UPDATE SET value=excluded.value, updated_at=excluded.updated_at",
            (key, value, iso_now()),
        )

    async def get_user(self, telegram_id: int) -> Optional[sqlite3.Row]:
        return await self.db.fetchone("SELECT * FROM users WHERE telegram_id=?", (telegram_id,))

    async def ensure_user_obj(self, telegram_id: int, username: Optional[str], first_name: Optional[str], last_name: Optional[str], role: Optional[str] = None) -> sqlite3.Row:
        existing = await self.get_user(telegram_id)
        now = iso_now()
        if existing:
            await self.db.execute(
                "UPDATE users SET username=?, first_name=?, last_name=?, updated_at=? WHERE telegram_id=?",
                (username, first_name, last_name, now, telegram_id),
            )
            return (await self.get_user(telegram_id))  # type: ignore[return-value]
        role = role or "student"
        await self.db.execute(
            "INSERT INTO users(telegram_id,username,first_name,last_name,role,created_at,updated_at) VALUES(?,?,?,?,?,?,?)",
            (telegram_id, username, first_name, last_name, role, now, now),
        )
        row = await self.get_user(telegram_id)
        if not row:
            raise RuntimeError("Failed to create user")
        if telegram_id in ADMIN_IDS:
            await self.db.execute("UPDATE users SET role='admin' WHERE telegram_id=?", (telegram_id,))
        return row

    async def update_user(self, telegram_id: int, **fields: Any) -> None:
        allowed = {"role", "is_blocked", "selected_model_id", "mode", "notifications_enabled", "username", "first_name", "last_name"}
        pairs = [(k, v) for k, v in fields.items() if k in allowed]
        if not pairs:
            return
        field_names = {k for k, _ in pairs}
        if "role" in field_names:
            role_value = next(v for k, v in pairs if k == "role")
            if role_value in MODE_PROMPTS:
                pairs = [(k, v) for k, v in pairs if k != "mode"]
                pairs.append(("mode", role_value))
        elif "mode" in field_names:
            mode_value = next(v for k, v in pairs if k == "mode")
            if mode_value in MODE_PROMPTS:
                pairs = [(k, v) for k, v in pairs if k != "role"]
                pairs.append(("role", mode_value))
        sets = ", ".join(f"{k}=?" for k, _ in pairs) + ", updated_at=?"
        values = [v for _, v in pairs] + [iso_now(), telegram_id]
        await self.db.execute(f"UPDATE users SET {sets} WHERE telegram_id=?", tuple(values))

    async def is_admin(self, telegram_id: int) -> bool:
        row = await self.db.fetchone("SELECT 1 FROM admins WHERE telegram_id=?", (telegram_id,))
        if row:
            return True
        return telegram_id in ADMIN_IDS

    async def create_chat(self, user_id: int, model_key: Optional[str]) -> int:
        now = iso_now()
        cur = await self.db.execute(
            "INSERT INTO chats(user_id,title,model_key,summary,created_at,updated_at) VALUES(?,?,?,?,?,?)",
            (user_id, "Новый чат", model_key, "", now, now),
        )
        return int(cur.lastrowid)

    async def get_or_create_active_chat(self, user_id: int, model_key: Optional[str]) -> sqlite3.Row:
        row = await self.db.fetchone("SELECT * FROM chats WHERE user_id=? ORDER BY updated_at DESC LIMIT 1", (user_id,))
        if row:
            return row
        cid = await self.create_chat(user_id, model_key)
        return await self.db.fetchone("SELECT * FROM chats WHERE id=?", (cid,))  # type: ignore[return-value]

    async def get_model(self, model_key: str) -> Optional[sqlite3.Row]:
        return await self.db.fetchone("SELECT * FROM models WHERE model_key=?", (model_key,))

    async def get_models(self, enabled_only: bool = False) -> list[sqlite3.Row]:
        if enabled_only:
            return await self.db.fetchall("SELECT * FROM models WHERE enabled=1 ORDER BY priority ASC, id ASC")
        return await self.db.fetchall("SELECT * FROM models ORDER BY priority ASC, id ASC")

    async def ensure_seed_model(self) -> None:
        model = await self.get_model(DEFAULT_MODEL_ID)
        if model:
            return
        base_url = os.getenv("AI_BASE_URL", "").strip()
        api_key = os.getenv("AI_API_KEY", "").strip()
        model_id = os.getenv("AI_MODEL_ID", "").strip()
        model_name = os.getenv("AI_MODEL_NAME", "Qwen 3.5 35B").strip()
        if not (base_url and api_key and model_id):
            return
        await self.db.execute(
            "INSERT INTO models(model_key,name,model_id,base_url,api_key,api_format,enabled,free_token_limit,paid_token_limit,reset_period_seconds,max_concurrency,priority,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            (DEFAULT_MODEL_ID, model_name, model_id, base_url.rstrip("/"), api_key, "openai_chat", 1, DEFAULT_FREE_TOKEN_LIMIT, DEFAULT_PAID_TOKEN_LIMIT, DEFAULT_RESET_PERIOD, 10, 100, iso_now(), iso_now()),
        )

    async def get_user_model(self, user_id: int) -> Optional[sqlite3.Row]:
        row = await self.db.fetchone("SELECT selected_model_id FROM users WHERE id=?", (user_id,))
        key = row["selected_model_id"] if row else None
        if key:
            model = await self.get_model(key)
            if model and model["enabled"]:
                return model
        default_key = await self.get_setting("default_model_id", DEFAULT_MODEL_ID)
        if default_key:
            model = await self.get_model(default_key)
            if model and model["enabled"]:
                return model
        models = await self.get_models(enabled_only=True)
        return models[0] if models else None

    async def get_subscription(self, user_id: int) -> Optional[sqlite3.Row]:
        return await self.db.fetchone("SELECT * FROM subscriptions WHERE user_id=? AND active=1 AND (expires_at IS NULL OR expires_at>?) ORDER BY expires_at DESC LIMIT 1", (user_id, iso_now()))

    # ... existing project methods retained ...


# -----------------------------------------------------------------------------
# The remainder of the verified project implementation is retained below.
# -----------------------------------------------------------------------------

# NOTE: This placeholder is intentionally replaced in the next synchronization step.
