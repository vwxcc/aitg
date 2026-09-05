#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Telegram AI Bot
============================================================

Минимальная архитектура:

Telegram
   │
   ├── обычный текст
   │
   ├── документы
   │
   └── изображения
          │
          ▼
     File Processor
          │
          ▼
       AI Queue
          │
          ▼
      OdiRouter API
          │
          ▼
      Telegram ответ


ХРАНЕНИЕ:
- никакой SQLite
- пользователи -> JSON
- история -> JSON
- временные файлы -> удаляются после обработки

АДМИН:
- только рассылка:
    /send текст сообщения

ADMIN_TELEGRAM_IDS:
- список Telegram ID администраторов

Поддерживаемые файлы:
- PDF
- DOCX
- XLSX
- PPTX
- JPG
- JPEG
- PNG
- WEBP

Остальные типы файлов отклоняются сразу.

Переменные Render:

BOT_TOKEN
ADMIN_TELEGRAM_IDS

AI_BASE_URL
AI_API_KEY
AI_MODEL_ID
AI_MODEL_NAME

SYSTEM_PROMPT

MAX_FILE_SIZE_MB
MAX_CONCURRENT_AI_REQUESTS
MAX_CONCURRENT_FILE_PROCESSING
AI_TIMEOUT_SECONDS

HISTORY_COMPRESS_WORDS
MAX_HISTORY_MESSAGES

DATA_DIR
"""

from __future__ import annotations

import asyncio
import base64
import json
import logging
import os
import re
import time
import uuid

from pathlib import Path
from typing import Any, Optional

import aiohttp

from aiogram import Bot, Dispatcher, F, Router
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.exceptions import (
    TelegramBadRequest,
    TelegramForbiddenError,
    TelegramNetworkError,
    TelegramRetryAfter,
)
from aiogram.filters import Command, CommandStart
from aiogram.types import (
    BufferedInputFile,
    CallbackQuery,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    Message,
)

from dotenv import load_dotenv

# ------------------------------------------------------------
# FILE PARSERS
# ------------------------------------------------------------

from pypdf import PdfReader
from docx import Document
from openpyxl import load_workbook
from pptx import Presentation


# ============================================================
# ENVIRONMENT
# ============================================================

load_dotenv()


BASE_DIR = Path(__file__).resolve().parent

DATA_DIR = Path(
    os.getenv(
        "DATA_DIR",
        str(BASE_DIR / "data"),
    )
)

USERS_FILE = DATA_DIR / "users.json"
HISTORY_DIR = DATA_DIR / "history"
FILES_DIR = DATA_DIR / "files"


DATA_DIR.mkdir(
    parents=True,
    exist_ok=True,
)

HISTORY_DIR.mkdir(
    parents=True,
    exist_ok=True,
)

FILES_DIR.mkdir(
    parents=True,
    exist_ok=True,
)


# ============================================================
# BOT
# ============================================================

BOT_TOKEN = os.getenv(
    "BOT_TOKEN",
    "",
).strip()


def parse_admin_ids() -> set[int]:
    raw = os.getenv(
        "ADMIN_TELEGRAM_IDS",
        "",
    )

    result: set[int] = set()

    for value in raw.replace(
        ";",
        ",",
    ).replace(
        "\n",
        ",",
    ).split(","):

        value = value.strip()

        if not value:
            continue

        if value.startswith("+"):
            value = value[1:].strip()

        if value.lstrip("-").isdigit():
            result.add(int(value))

    return result


ADMIN_IDS: set[int] = parse_admin_ids()


# ============================================================
# AI
# ============================================================

AI_BASE_URL = os.getenv(
    "AI_BASE_URL",
    "",
).strip().rstrip("/")


AI_API_KEY = os.getenv(
    "AI_API_KEY",
    "",
).strip()


AI_MODEL_ID = os.getenv(
    "AI_MODEL_ID",
    "",
).strip()


AI_MODEL_NAME = os.getenv(
    "AI_MODEL_NAME",
    "AI",
).strip()


SYSTEM_PROMPT = os.getenv(
    "SYSTEM_PROMPT",
    "",
).strip()


# ============================================================
# LIMITS / PERFORMANCE
# ============================================================

MAX_FILE_SIZE_MB = int(
    os.getenv(
        "MAX_FILE_SIZE_MB",
        "10",
    )
)

MAX_FILE_SIZE_BYTES = (
    MAX_FILE_SIZE_MB * 1024 * 1024
)


MAX_CONCURRENT_AI_REQUESTS = max(
    1,
    int(
        os.getenv(
            "MAX_CONCURRENT_AI_REQUESTS",
            "2",
        )
    ),
)


MAX_CONCURRENT_FILE_PROCESSING = max(
    1,
    int(
        os.getenv(
            "MAX_CONCURRENT_FILE_PROCESSING",
            "2",
        )
    ),
)


AI_TIMEOUT_SECONDS = max(
    30,
    int(
        os.getenv(
            "AI_TIMEOUT_SECONDS",
            "120",
        )
    ),
)


HISTORY_COMPRESS_WORDS = max(
    1000,
    int(
        os.getenv(
            "HISTORY_COMPRESS_WORDS",
            "50000",
        )
    ),
)


MAX_HISTORY_MESSAGES = max(
    10,
    int(
        os.getenv(
            "MAX_HISTORY_MESSAGES",
            "100",
        )
    ),
)


# ============================================================
# BROADCAST
# ============================================================

# Небольшая пауза между сообщениями рассылки.
# Это специально не делает рассылку мгновенной.
BROADCAST_DELAY = float(
    os.getenv(
        "BROADCAST_DELAY",
        "0.05",
    )
)


# Сколько одновременно отправляем максимум.
BROADCAST_CONCURRENCY = max(
    1,
    int(
        os.getenv(
            "BROADCAST_CONCURRENCY",
            "5",
        )
    ),
)


# ============================================================
# LOGGING
# ============================================================

LOG_LEVEL = os.getenv(
    "LOG_LEVEL",
    "INFO",
).upper()


logging.basicConfig(
    level=getattr(
        logging,
        LOG_LEVEL,
        logging.INFO,
    ),
    format=(
        "%(asctime)s | "
        "%(levelname)s | "
        "%(name)s | "
        "%(message)s"
    ),
)


log = logging.getLogger(
    "telegram_ai_bot"
)


# ============================================================
# CONSTANTS
# ============================================================

ALLOWED_EXTENSIONS = {
    ".pdf",
    ".docx",
    ".xlsx",
    ".pptx",
    ".jpg",
    ".jpeg",
    ".png",
    ".webp",
}


ALLOWED_MIME_TYPES = {
    "application/pdf",

    "application/vnd.openxmlformats-officedocument"
    ".wordprocessingml.document",

    "application/vnd.openxmlformats-officedocument"
    ".spreadsheetml.sheet",

    "application/vnd.openxmlformats-officedocument"
    ".presentationml.presentation",

    "image/jpeg",
    "image/png",
    "image/webp",
}


# ============================================================
# TEXT HELPERS
# ============================================================

def clean_text(
    text: str,
) -> str:

    text = text.replace(
        "\x00",
        "",
    )

    text = re.sub(
        r"\n{4,}",
        "\n\n\n",
        text,
    )

    return text.strip()


def count_words(
    text: str,
) -> int:

    return len(
        re.findall(
            r"\S+",
            text,
        )
    )


def escape_html(
    text: str,
) -> str:

    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


# ============================================================
# JSON STORAGE
# ============================================================

class JSONStore:

    def __init__(
        self,
    ) -> None:

        self.lock = asyncio.Lock()

    # --------------------------------------------------------
    # USERS
    # --------------------------------------------------------

    async def _read_users(
        self,
    ) -> list[int]:

        if not USERS_FILE.exists():
            return []

        try:

            text = await asyncio.to_thread(
                USERS_FILE.read_text,
                encoding="utf-8",
            )

            data = json.loads(text)

            if not isinstance(
                data,
                list,
            ):
                return []

            result = []

            for value in data:

                try:
                    result.append(
                        int(value)
                    )

                except Exception:
                    continue

            return list(
                dict.fromkeys(result)
            )

        except Exception:

            log.exception(
                "Failed to read users.json"
            )

            return []

    async def _write_users(
        self,
        users: list[int],
    ) -> None:

        data = json.dumps(
            sorted(
                set(users)
            ),
            ensure_ascii=False,
            indent=2,
        )

        await asyncio.to_thread(
            USERS_FILE.write_text,
            data,
            encoding="utf-8",
        )

    async def add_user(
        self,
        user_id: int,
    ) -> None:

        async with self.lock:

            users = await self._read_users()

            if user_id not in users:

                users.append(
                    user_id
                )

                await self._write_users(
                    users
                )

    async def remove_user(
        self,
        user_id: int,
    ) -> None:

        async with self.lock:

            users = await self._read_users()

            if user_id in users:

                users.remove(
                    user_id
                )

                await self._write_users(
                    users
                )

    async def get_users(
        self,
    ) -> list[int]:

        async with self.lock:

            return await self._read_users()

    # --------------------------------------------------------
    # HISTORY
    # --------------------------------------------------------

    def history_path(
        self,
        user_id: int,
    ) -> Path:

        return (
            HISTORY_DIR
            / f"{user_id}.json"
        )

    async def get_history(
        self,
        user_id: int,
    ) -> list[dict[str, Any]]:

        path = self.history_path(
            user_id
        )

        if not path.exists():
            return []

        try:

            text = await asyncio.to_thread(
                path.read_text,
                encoding="utf-8",
            )

            data = json.loads(
                text
            )

            if not isinstance(
                data,
                list,
            ):
                return []

            return data

        except Exception:

            log.exception(
                "Failed to read history for %s",
                user_id,
            )

            return []

    async def save_history(
        self,
        user_id: int,
        history: list[dict[str, Any]],
    ) -> None:

        path = self.history_path(
            user_id
        )

        data = json.dumps(
            history,
            ensure_ascii=False,
            indent=2,
        )

        await asyncio.to_thread(
            path.write_text,
            data,
            encoding="utf-8",
        )


# ============================================================
# FILE PROCESSOR
# ============================================================

class FileProcessor:

    def __init__(
        self,
    ) -> None:

        self.semaphore = asyncio.Semaphore(
            MAX_CONCURRENT_FILE_PROCESSING
        )

    # --------------------------------------------------------
    # CHECK
    # --------------------------------------------------------

    @staticmethod
    def allowed(
        filename: str,
        mime: str,
    ) -> bool:

        extension = (
            Path(filename)
            .suffix
            .lower()
        )

        if extension in ALLOWED_EXTENSIONS:
            return True

        if mime in ALLOWED_MIME_TYPES:
            return True

        return False

    # --------------------------------------------------------
    # PROCESS
    # --------------------------------------------------------

    async def process(
        self,
        path: Path,
        filename: str,
        mime: str,
    ) -> dict[str, Any]:

        async with self.semaphore:

            try:

                return await asyncio.wait_for(
                    asyncio.to_thread(
                        self._process_sync,
                        path,
                        filename,
                        mime,
                    ),
                    timeout=AI_TIMEOUT_SECONDS,
                )

            finally:

                try:
                    path.unlink(
                        missing_ok=True
                    )

                except Exception:

                    log.exception(
                        "Failed to delete temp file"
                    )

    # --------------------------------------------------------
    # SYNC PROCESSOR
    # --------------------------------------------------------

    def _process_sync(
        self,
        path: Path,
        filename: str,
        mime: str,
    ) -> dict[str, Any]:

        extension = (
            Path(filename)
            .suffix
            .lower()
        )

        # ----------------------------------------------------
        # PDF
        # ----------------------------------------------------

        if extension == ".pdf":

            reader = PdfReader(
                str(path)
            )

            pages = []

            for page in reader.pages:

                try:

                    text = page.extract_text()

                    if text:
                        pages.append(
                            text
                        )

                except Exception:
                    continue

            return {
                "type": "text",
                "filename": filename,
                "text": clean_text(
                    "\n\n".join(pages)
                ),
            }

        # ----------------------------------------------------
        # DOCX
        # ----------------------------------------------------

        if extension == ".docx":

            document = Document(
                str(path)
            )

            parts = []

            for paragraph in document.paragraphs:

                text = paragraph.text.strip()

                if text:
                    parts.append(
                        text
                    )

            # Таблицы DOCX
            for table in document.tables:

                rows = []

                for row in table.rows:

                    cells = [
                        cell.text.strip()
                        for cell in row.cells
                    ]

                    rows.append(
                        " | ".join(cells)
                    )

                if rows:
                    parts.append(
                        "\n".join(rows)
                    )

            return {
                "type": "text",
                "filename": filename,
                "text": clean_text(
                    "\n\n".join(parts)
                ),
            }

        # ----------------------------------------------------
        # XLSX
        # ----------------------------------------------------

        if extension == ".xlsx":

            workbook = load_workbook(
                filename=str(path),
                read_only=True,
                data_only=True,
            )

            parts = []

            for sheet in workbook.worksheets:

                parts.append(
                    f"Лист: {sheet.title}"
                )

                for row in sheet.iter_rows(
                    values_only=True
                ):

                    values = []

                    for value in row:

                        if value is None:
                            values.append("")
                        else:
                            values.append(
                                str(value)
                            )

                    if any(values):

                        parts.append(
                            " | ".join(values)
                        )

            workbook.close()

            return {
                "type": "text",
                "filename": filename,
                "text": clean_text(
                    "\n".join(parts)
                ),
            }

        # ----------------------------------------------------
        # PPTX
        # ----------------------------------------------------

        if extension == ".pptx":

            presentation = Presentation(
                str(path)
            )

            parts = []

            for index, slide in enumerate(
                presentation.slides,
                start=1,
            ):

                slide_parts = [
                    f"Слайд {index}:"
                ]

                for shape in slide.shapes:

                    if hasattr(
                        shape,
                        "text",
                    ):

                        text = (
                            shape.text
                            .strip()
                        )

                        if text:
                            slide_parts.append(
                                text
                            )

                parts.append(
                    "\n".join(
                        slide_parts
                    )
                )

            return {
                "type": "text",
                "filename": filename,
                "text": clean_text(
                    "\n\n".join(parts)
                ),
            }

        # ----------------------------------------------------
        # IMAGES
        # ----------------------------------------------------

        if extension in {
            ".jpg",
            ".jpeg",
            ".png",
            ".webp",
        }:

            raw = path.read_bytes()

            encoded = base64.b64encode(
                raw
            ).decode(
                "ascii"
            )

            if extension in {
                ".jpg",
                ".jpeg",
            }:
                image_mime = "image/jpeg"

            elif extension == ".png":
                image_mime = "image/png"

            else:
                image_mime = "image/webp"

            return {
                "type": "image",
                "filename": filename,
                "data_url": (
                    f"data:{image_mime};"
                    f"base64,{encoded}"
                ),
            }

        raise ValueError(
            "Unsupported file type"
        )


# ============================================================
# AI SERVICE
# ============================================================

class AIService:

    def __init__(
        self,
    ) -> None:

        self.session: Optional[
            aiohttp.ClientSession
        ] = None

        self.semaphore = asyncio.Semaphore(
            MAX_CONCURRENT_AI_REQUESTS
        )

    # --------------------------------------------------------
    # START
    # --------------------------------------------------------

    async def start(
        self,
    ) -> None:

        timeout = aiohttp.ClientTimeout(
            total=AI_TIMEOUT_SECONDS
        )

        self.session = aiohttp.ClientSession(
            timeout=timeout,
            headers={
                "Authorization":
                    f"Bearer {AI_API_KEY}",
                "Content-Type":
                    "application/json",
            },
        )

    # --------------------------------------------------------
    # CLOSE
    # --------------------------------------------------------

    async def close(
        self,
    ) -> None:

        if self.session:

            await self.session.close()

            self.session = None

    # --------------------------------------------------------
    # SYSTEM PROMPT
    # --------------------------------------------------------

    def get_system_prompt(
        self,
    ) -> str:

        if SYSTEM_PROMPT:

            return (
                f"Ты — {AI_MODEL_NAME}, "
                "AI-модель внутри Telegram-бота.\n\n"
                f"{SYSTEM_PROMPT}"
            )

        return (
            f"Ты — {AI_MODEL_NAME}, "
            "AI-модель внутри Telegram-бота. "
            "Отвечай полезно, точно и понятно."
        )

    # --------------------------------------------------------
    # NORMALIZE MESSAGES
    # --------------------------------------------------------

    def normalize_messages(
        self,
        messages: list[dict[str, Any]],
    ) -> list[dict[str, Any]]:

        system_messages = [
            message
            for message in messages
            if message.get("role")
            == "system"
        ]

        other_messages = [
            message
            for message in messages
            if message.get("role")
            != "system"
        ]

        if system_messages:

            return [
                system_messages[0],
                *other_messages,
            ]

        return [
            {
                "role": "system",
                "content":
                    self.get_system_prompt(),
            },
            *other_messages,
        ]

    # --------------------------------------------------------
    # REQUEST
    # --------------------------------------------------------

    async def request(
        self,
        messages: list[dict[str, Any]],
        temperature: float = 0.7,
    ) -> str:

        if not self.session:
            raise RuntimeError(
                "AI service is not started"
            )

        payload = {
            "model": AI_MODEL_ID,
            "messages": self.normalize_messages(
                messages
            ),
            "temperature": temperature,
        }

        url = (
            f"{AI_BASE_URL}"
            "/chat/completions"
        )

        async with self.semaphore:

            try:

                async with self.session.post(
                    url,
                    json=payload,
                ) as response:

                    text = await response.text()

                    if response.status >= 400:

                        raise RuntimeError(
                            f"AI API error "
                            f"{response.status}: "
                            f"{text[:2000]}"
                        )

                    try:

                        data = json.loads(
                            text
                        )

                    except json.JSONDecodeError:

                        raise RuntimeError(
                            "AI API returned "
                            "invalid JSON"
                        )

                    choices = data.get(
                        "choices"
                    )

                    if not choices:
                        raise RuntimeError(
                            "AI API returned "
                            "no choices"
                        )

                    message = choices[0].get(
                        "message",
                        {},
                    )

                    content = message.get(
                        "content"
                    )

                    if isinstance(
                        content,
                        list,
                    ):

                        parts = []

                        for item in content:

                            if isinstance(
                                item,
                                dict,
                            ):

                                if item.get(
                                    "type"
                                ) == "text":

                                    parts.append(
                                        str(
                                            item.get(
                                                "text",
                                                "",
                                            )
                                        )
                                    )

                        content = "\n".join(
                            parts
                        )

                    if not content:

                        raise RuntimeError(
                            "AI returned empty "
                            "response"
                        )

                    return str(
                        content
                    ).strip()

            except asyncio.CancelledError:

                raise

            except Exception:

                log.exception(
                    "AI request failed"
                )

                raise


# ============================================================
# AI QUEUE
# ============================================================

class AIQueue:

    def __init__(
        self,
        ai: AIService,
    ) -> None:

        self.ai = ai

        self.queue: asyncio.Queue[
            tuple[
                int,
                list[dict[str, Any]],
                asyncio.Future,
            ]
        ] = asyncio.Queue()

        self.workers: list[
            asyncio.Task
        ] = []

        self.running = False

    # --------------------------------------------------------
    # START
    # --------------------------------------------------------

    async def start(
        self,
    ) -> None:

        if self.running:
            return

        self.running = True

        worker_count = (
            MAX_CONCURRENT_AI_REQUESTS
        )

        for index in range(
            worker_count
        ):

            task = asyncio.create_task(
                self.worker(),
                name=(
                    f"ai-worker-{index + 1}"
                ),
            )

            self.workers.append(
                task
            )

    # --------------------------------------------------------
    # STOP
    # --------------------------------------------------------

    async def stop(
        self,
    ) -> None:

        self.running = False

        tasks = list(
            self.workers
        )

        self.workers.clear()

        for task in tasks:

            task.cancel()

        for task in tasks:

            try:

                await task

            except asyncio.CancelledError:

                pass

    # --------------------------------------------------------
    # PUT
    # --------------------------------------------------------

    async def put(
        self,
        user_id: int,
        messages: list[dict[str, Any]],
    ) -> asyncio.Future:

        loop = asyncio.get_running_loop()

        future = loop.create_future()

        await self.queue.put(
            (
                user_id,
                messages,
                future,
            )
        )

        return future

    # --------------------------------------------------------
    # WORKER
    # --------------------------------------------------------

    async def worker(
        self,
    ) -> None:

        while self.running:

            try:

                (
                    user_id,
                    messages,
                    future,
                ) = await self.queue.get()

                try:

                    if future.cancelled():
                        continue

                    result = await self.ai.request(
                        messages
                    )

                    if not future.done():

                        future.set_result(
                            result
                        )

                except asyncio.CancelledError:

                    if not future.done():

                        future.cancel()

                    raise

                except Exception as exc:

                    if not future.done():

                        future.set_exception(
                            exc
                        )

                finally:

                    self.queue.task_done()

            except asyncio.CancelledError:

                raise

            except Exception:

                log.exception(
                    "AI queue worker failed"
                )


# ============================================================
# BOT APP
# ============================================================

class BotApp:

    def __init__(
        self,
    ) -> None:

        self.bot = Bot(
            token=BOT_TOKEN,
            default=DefaultBotProperties(
                parse_mode=ParseMode.HTML,
            ),
        )

        self.dp = Dispatcher()

        self.router = Router()

        self.dp.include_router(
            self.router
        )

        self.store = JSONStore()

        self.files = FileProcessor()

        self.ai = AIService()

        self.queue = AIQueue(
            self.ai
        )

        # user_id -> active task
        self.active_tasks: dict[
            int,
            asyncio.Task,
        ] = {}

        # user_id -> stop event
        self.stop_events: dict[
            int,
            asyncio.Event,
        ] = {}

        self.broadcast_lock = asyncio.Lock()

        self.register_handlers()

    # ========================================================
    # START
    # ========================================================

    async def start(
        self,
    ) -> None:

        if not BOT_TOKEN:

            raise RuntimeError(
                "BOT_TOKEN не задан"
            )

        if not AI_BASE_URL:

            raise RuntimeError(
                "AI_BASE_URL не задан"
            )

        if not AI_API_KEY:

            raise RuntimeError(
                "AI_API_KEY не задан"
            )

        if not AI_MODEL_ID:

            raise RuntimeError(
                "AI_MODEL_ID не задан"
            )

        await self.ai.start()

        await self.queue.start()

        log.info(
            "Bot started"
        )

        log.info(
            "AI model: %s (%s)",
            AI_MODEL_NAME,
            AI_MODEL_ID,
        )

        log.info(
            "Admins: %s",
            len(ADMIN_IDS),
        )

    # ========================================================
    # STOP
    # ========================================================

    async def stop(
        self,
    ) -> None:

        await self.queue.stop()

        await self.ai.close()

        await self.bot.session.close()

    # ========================================================
    # USER
    # ========================================================

    async def ensure_user(
        self,
        message: Message,
    ) -> Optional[int]:

        if not message.from_user:
            return None

        user_id = message.from_user.id

        try:

            await self.store.add_user(
                user_id
            )

        except Exception:

            log.exception(
                "Failed to save user %s",
                user_id,
            )

        return user_id

    # ========================================================
    # BUSY
    # ========================================================

    def is_busy(
        self,
        user_id: int,
    ) -> bool:

        task = self.active_tasks.get(
            user_id
        )

        if task is None:
            return False

        if task.done():

            self.active_tasks.pop(
                user_id,
                None,
            )

            self.stop_events.pop(
                user_id,
                None,
            )

            return False

        return True

    async def reject_if_busy(
        self,
        message: Message,
    ) -> bool:

        if not message.from_user:
            return True

        user_id = message.from_user.id

        if self.is_busy(user_id):

            await message.answer(
                "⏳ Я ещё обрабатываю "
                "предыдущий запрос. "
                "Подождите немного 🙂"
            )

            return True

        return False

    # ========================================================
    # TASK REGISTER
    # ========================================================

    def register_task(
        self,
        user_id: int,
        task: asyncio.Task,
        stop_event: asyncio.Event,
    ) -> None:

        self.active_tasks[
            user_id
        ] = task

        self.stop_events[
            user_id
        ] = stop_event

        def done_callback(
            finished: asyncio.Task,
        ) -> None:

            current = (
                self.active_tasks.get(
                    user_id
                )
            )

            if current is finished:

                self.active_tasks.pop(
                    user_id,
                    None,
                )

                self.stop_events.pop(
                    user_id,
                    None,
                )

        task.add_done_callback(
            done_callback
        )

    # ========================================================
    # STOP KEYBOARD
    # ========================================================

    @staticmethod
    def stop_keyboard() -> InlineKeyboardMarkup:

        return InlineKeyboardMarkup(
            inline_keyboard=[
                [
                    InlineKeyboardButton(
                        text="⏹ Стоп",
                        callback_data="stop_request",
                    )
                ]
            ]
        )

    # ========================================================
    # /START
    # ========================================================

    async def handle_start(
        self,
        message: Message,
    ) -> None:

        await self.ensure_user(
            message
        )

        await message.answer(
            "Привет! 👋\n\n"
            "Отправь мне сообщение, "
            "и я помогу с ним.\n\n"
            "Также можно отправить PDF, "
            "DOCX, XLSX, PPTX или изображение."
        )

    # ========================================================
    # /SEND
    # ========================================================

    async def handle_send(
        self,
        message: Message,
    ) -> None:

        if not message.from_user:

            return

        user_id = message.from_user.id

        # ----------------------------------------------------
        # ADMIN CHECK
        # ----------------------------------------------------

        if user_id not in ADMIN_IDS:

            await message.answer(
                "⛔ Доступ запрещён."
            )

            return

        # ----------------------------------------------------
        # TEXT
        # ----------------------------------------------------

        text = message.text or ""

        # Убираем команду /send
        parts = text.split(
            maxsplit=1
        )

        if len(parts) < 2:

            await message.answer(
                "Использование:\n\n"
                "<code>/send текст рассылки</code>"
            )

            return

        broadcast_text = parts[1].strip()

        if not broadcast_text:

            await message.answer(
                "❗ Напиши текст рассылки "
                "после команды /send."
            )

            return

        # ----------------------------------------------------
        # START BROADCAST
        # ----------------------------------------------------

        task = asyncio.create_task(
            self.broadcast(
                admin_message=message,
                text=broadcast_text,
            )
        )

        # Не ждём рассылку внутри Telegram handler.
        # Она работает отдельно.
        task.add_done_callback(
            self.broadcast_task_done
        )

    # ========================================================
    # BROADCAST TASK CALLBACK
    # ========================================================

    @staticmethod
    def broadcast_task_done(
        task: asyncio.Task,
    ) -> None:

        try:

            task.result()

        except asyncio.CancelledError:

            pass

        except Exception:

            log.exception(
                "Broadcast task failed"
            )

    # ========================================================
    # BROADCAST
    # ========================================================

    async def broadcast(
        self,
        admin_message: Message,
        text: str,
    ) -> None:

        async with self.broadcast_lock:

            users = await self.store.get_users()

            total = len(users)

            if total == 0:

                await admin_message.answer(
                    "📢 Рассылать пока некому."
                )

                return

            status = await admin_message.answer(
                "📢 Рассылка запущена.\n\n"
                f"Получателей: <b>{total}</b>\n"
                "Отправляю..."
            )

            semaphore = asyncio.Semaphore(
                BROADCAST_CONCURRENCY
            )

            sent = 0
            failed = 0
            removed = 0

            counter_lock = asyncio.Lock()

            async def send_one(
                user_id: int,
            ) -> None:

                nonlocal sent
                nonlocal failed
                nonlocal removed

                async with semaphore:

                    try:

                        await self.bot.send_message(
                            chat_id=user_id,
                            text=text,
                        )

                        async with counter_lock:

                            sent += 1

                    except TelegramForbiddenError:

                        # Пользователь заблокировал бота
                        await self.store.remove_user(
                            user_id
                        )

                        async with counter_lock:

                            failed += 1
                            removed += 1

                    except TelegramRetryAfter as exc:

                        # Telegram попросил подождать
                        await asyncio.sleep(
                            float(
                                exc.retry_after
                            )
                        )

                        try:

                            await self.bot.send_message(
                                chat_id=user_id,
                                text=text,
                            )

                            async with counter_lock:

                                sent += 1

                        except TelegramForbiddenError:

                            await self.store.remove_user(
                                user_id
                            )

                            async with counter_lock:

                                failed += 1
                                removed += 1

                        except Exception:

                            async with counter_lock:

                                failed += 1

                    except TelegramNetworkError:

                        async with counter_lock:

                            failed += 1

                    except TelegramBadRequest:

                        async with counter_lock:

                            failed += 1

                    except Exception:

                        log.exception(
                            "Broadcast failed for %s",
                            user_id,
                        )

                        async with counter_lock:

                            failed += 1

                    await asyncio.sleep(
                        BROADCAST_DELAY
                    )

            # ------------------------------------------------
            # SEND IN BATCHES
            # ------------------------------------------------

            batch_size = (
                BROADCAST_CONCURRENCY * 10
            )

            for start in range(
                0,
                total,
                batch_size,
            ):

                batch = users[
                    start:
                    start + batch_size
                ]

                await asyncio.gather(
                    *(
                        send_one(
                            user_id
                        )
                        for user_id in batch
                    )
                )

            # ------------------------------------------------
            # RESULT
            # ------------------------------------------------

            try:

                await status.edit_text(
                    "📢 <b>Рассылка завершена.</b>\n\n"
                    f"👥 Получателей: {total}\n"
                    f"✅ Доставлено: {sent}\n"
                    f"❌ Ошибок: {failed}\n"
                    f"🗑 Удалено из списка: {removed}"
                )

            except Exception:

                log.exception(
                    "Failed to edit broadcast status"
                )

    # ========================================================
    # STOP CALLBACK
    # ========================================================

    async def handle_stop(
        self,
        callback: CallbackQuery,
    ) -> None:

        if not callback.from_user:

            await callback.answer()

            return

        user_id = callback.from_user.id

        task = self.active_tasks.get(
            user_id
        )

        if not task or task.done():

            await callback.answer(
                "Запрос уже завершён.",
                show_alert=False,
            )

            return

        stop_event = self.stop_events.get(
            user_id
        )

        if stop_event:

            stop_event.set()

        task.cancel()

        await callback.answer(
            "⏹ Остановлено.",
            show_alert=False,
        )

        try:

            if callback.message:

                await callback.message.edit_text(
                    "⏹ Запрос остановлен."
                )

        except Exception:
            pass

    # ========================================================
    # TEXT MESSAGE
    # ========================================================

    async def handle_text(
        self,
        message: Message,
    ) -> None:

        if not message.from_user:
            return

        # /send должен обрабатываться
        # отдельным handler.
        if message.text and message.text.startswith(
            "/send"
        ):
            return

        await self.ensure_user(
            message
        )

        if await self.reject_if_busy(
            message
        ):
            return

        text = (
            message.text or ""
        ).strip()

        if not text:
            return

        task = asyncio.create_task(
            self.process_request(
                message=message,
                user_text=text,
                file_context=None,
                image_data_url=None,
            )
        )

        self.register_task(
            message.from_user.id,
            task,
            asyncio.Event(),
        )

    # ========================================================
    # DOCUMENT
    # ========================================================

    async def handle_document(
        self,
        message: Message,
    ) -> None:

        if not message.from_user:
            return

        await self.ensure_user(
            message
        )

        if await self.reject_if_busy(
            message
        ):
            return

        document = message.document

        if not document:
            return

        filename = (
            document.file_name
            or ""
        )

        mime = (
            document.mime_type
            or ""
        )

        # ----------------------------------------------------
        # TYPE CHECK BEFORE DOWNLOAD
        # ----------------------------------------------------

        if not FileProcessor.allowed(
            filename,
            mime,
        ):

            await message.answer(
                "❌ Этот тип файла "
                "пока не поддерживается."
            )

            return

        # ----------------------------------------------------
        # SIZE CHECK BEFORE DOWNLOAD
        # ----------------------------------------------------

        size = (
            document.file_size
            or 0
        )

        if size > MAX_FILE_SIZE_BYTES:

            await message.answer(
                "❌ Файл слишком большой.\n"
                f"Максимум — "
                f"{MAX_FILE_SIZE_MB} МБ."
            )

            return

        # ----------------------------------------------------
        # CREATE TASK
        # ----------------------------------------------------

        stop_event = asyncio.Event()

        task = asyncio.create_task(
            self.process_document(
                message=message,
                file_id=document.file_id,
                filename=filename,
                mime=mime,
            )
        )

        self.register_task(
            message.from_user.id,
            task,
            stop_event,
        )

    # ========================================================
    # PHOTO
    # ========================================================

    async def handle_photo(
        self,
        message: Message,
    ) -> None:

        if not message.from_user:
            return

        await self.ensure_user(
            message
        )

        if await self.reject_if_busy(
            message
        ):
            return

        if not message.photo:
            return

        photo = message.photo[-1]

        if (
            photo.file_size
            and photo.file_size
            > MAX_FILE_SIZE_BYTES
        ):

            await message.answer(
                "❌ Фото слишком большое.\n"
                f"Максимум — "
                f"{MAX_FILE_SIZE_MB} МБ."
            )

            return

        stop_event = asyncio.Event()

        task = asyncio.create_task(
            self.process_photo(
                message,
                photo.file_id,
            )
        )

        self.register_task(
            message.from_user.id,
            task,
            stop_event,
        )

    # ========================================================
    # PROCESS DOCUMENT
    # ========================================================

    async def process_document(
        self,
        message: Message,
        file_id: str,
        filename: str,
        mime: str,
    ) -> None:

        if not message.from_user:
            return

        user_id = message.from_user.id

        status = None
        path = None

        try:

            status = await message.answer(
                "📄 Читаю файл…",
                reply_markup=self.stop_keyboard(),
            )

            # ------------------------------------------------
            # TELEGRAM FILE
            # ------------------------------------------------

            telegram_file = await self.bot.get_file(
                file_id
            )

            if not telegram_file.file_path:

                raise RuntimeError(
                    "Telegram не вернул путь файла"
                )

            safe_name = (
                re.sub(
                    r"[^a-zA-Z0-9а-яА-Я._-]",
                    "_",
                    filename,
                )
                or "file"
            )

            path = (
                FILES_DIR
                / f"{uuid.uuid4().hex}_{safe_name}"
            )

            await self.bot.download_file(
                telegram_file.file_path,
                destination=path,
            )

            # ------------------------------------------------
            # PROCESS FILE
            # ------------------------------------------------

            if status:

                try:

                    await status.edit_text(
                        "📄 Анализирую файл…",
                        reply_markup=self.stop_keyboard(),
                    )

                except Exception:
                    pass

            result = await self.files.process(
                path,
                filename,
                mime,
            )

            path = None

            # ------------------------------------------------
            # IMAGE
            # ------------------------------------------------

            if result.get("type") == "image":

                await self.process_request(
                    message=message,
                    user_text=(
                        message.caption
                        or "Проанализируй это изображение."
                    ),
                    file_context=None,
                    image_data_url=result.get(
                        "data_url"
                    ),
                    status_message=status,
                )

                return

            # ------------------------------------------------
            # TEXT
            # ------------------------------------------------

            file_text = (
                result.get(
                    "text",
                    "",
                )
                or ""
            )

            if not file_text.strip():

                await message.answer(
                    "❌ Не удалось извлечь "
                    "текст из файла."
                )

                if status:

                    try:
                        await status.delete()
                    except Exception:
                        pass

                return

            user_text = (
                message.caption
                or "Проанализируй содержимое этого файла."
            )

            file_context = (
                f"\n\n"
                f"--- Файл: {filename} ---\n"
                f"{file_text}\n"
                f"--- Конец файла ---"
            )

            await self.process_request(
                message=message,
                user_text=user_text,
                file_context=file_context,
                image_data_url=None,
                status_message=status,
            )

        except asyncio.CancelledError:

            try:

                if status:
                    await status.delete()

            except Exception:
                pass

            raise

        except Exception as exc:

            log.exception(
                "Document processing failed"
            )

            try:

                if status:
                    await status.delete()

            except Exception:
                pass

            await message.answer(
                "❌ Не удалось обработать файл."
            )

        finally:

            if path:

                try:

                    path.unlink(
                        missing_ok=True
                    )

                except Exception:
                    pass

    # ========================================================
    # PROCESS PHOTO
    # ========================================================

    async def process_photo(
        self,
        message: Message,
        file_id: str,
    ) -> None:

        status = None
        path = None

        try:

            status = await message.answer(
                "👀 Смотрю фото…",
                reply_markup=self.stop_keyboard(),
            )

            telegram_file = await self.bot.get_file(
                file_id
            )

            if not telegram_file.file_path:

                raise RuntimeError(
                    "Telegram не вернул путь изображения"
                )

            path = (
                FILES_DIR
                / f"{uuid.uuid4().hex}.jpg"
            )

            await self.bot.download_file(
                telegram_file.file_path,
                destination=path,
            )

            raw = path.read_bytes()

            encoded = base64.b64encode(
                raw
            ).decode(
                "ascii"
            )

            data_url = (
                "data:image/jpeg;base64,"
                + encoded
            )

            await self.process_request(
                message=message,
                user_text=(
                    message.caption
                    or "Проанализируй это изображение."
                ),
                file_context=None,
                image_data_url=data_url,
                status_message=status,
            )

        except asyncio.CancelledError:

            try:

                if status:
                    await status.delete()

            except Exception:
                pass

            raise

        except Exception:

            log.exception(
                "Photo processing failed"
            )

            try:

                if status:
                    await status.delete()

            except Exception:
                pass

            await message.answer(
                "❌ Не удалось обработать изображение."
            )

        finally:

            if path:

                try:

                    path.unlink(
                        missing_ok=True
                    )

                except Exception:
                    pass

    # ========================================================
    # UNSUPPORTED MEDIA
    # ========================================================

    async def handle_unsupported_media(
        self,
        message: Message,
    ) -> None:

        if message.from_user:

            await self.ensure_user(
                message
            )

        await message.answer(
            "❌ Этот тип файла "
            "пока не поддерживается."
        )

    # ========================================================
    # MAIN AI PROCESS
    # ========================================================

    async def process_request(
        self,
        message: Message,
        user_text: str,
        file_context: Optional[str],
        image_data_url: Optional[str],
        status_message: Optional[Message] = None,
    ) -> None:

        if not message.from_user:
            return

        user_id = message.from_user.id

        status = status_message

        try:

            # ------------------------------------------------
            # STATUS
            # ------------------------------------------------

            if status is None:

                status = await message.answer(
                    "🧠 Думаю…",
                    reply_markup=self.stop_keyboard(),
                )

            else:

                try:

                    await status.edit_text(
                        "🧠 Думаю…",
                        reply_markup=self.stop_keyboard(),
                    )

                except Exception:
                    pass

            # ------------------------------------------------
            # HISTORY
            # ------------------------------------------------

            history = await self.store.get_history(
                user_id
            )

            # ------------------------------------------------
            # COMPRESS OLD HISTORY
            # ------------------------------------------------

            history = await self.maybe_compress_history(
                user_id,
                history,
            )

            # ------------------------------------------------
            # USER MESSAGE
            # ------------------------------------------------

            if file_context:

                user_content = (
                    user_text
                    + file_context
                )

            else:

                user_content = user_text

            # ------------------------------------------------
            # HIDDEN PLAN
            # ------------------------------------------------

            plan_messages: list[
                dict[str, Any]
            ] = [
                {
                    "role": "system",
                    "content": (
                        self.ai.get_system_prompt()
                        + "\n\n"
                        "Перед основным ответом "
                        "сделай краткий внутренний "
                        "план решения. "
                        "Этот план не показывай пользователю."
                    ),
                }
            ]

            # Добавляем историю,
            # но только ограниченный объём.
            plan_messages.extend(
                self.prepare_history_for_ai(
                    history
                )
            )

            plan_messages.append(
                {
                    "role": "user",
                    "content": (
                        "Сначала составь краткий "
                        "внутренний план решения "
                        "следующего запроса. "
                        "Не отвечай пользователю "
                        "и не показывай план.\n\n"
                        + user_content
                    ),
                }
            )

            # ------------------------------------------------
            # PLAN REQUEST
            # ------------------------------------------------

            plan = await self.ai.request(
                plan_messages,
                temperature=0.2,
            )

            # ------------------------------------------------
            # MAIN REQUEST
            # ------------------------------------------------

            main_messages: list[
                dict[str, Any]
            ] = [
                {
                    "role": "system",
                    "content":
                        self.ai.get_system_prompt(),
                }
            ]

            main_messages.extend(
                self.prepare_history_for_ai(
                    history
                )
            )

            # ------------------------------------------------
            # MULTIMODAL
            # ------------------------------------------------

            if image_data_url:

                content = [
                    {
                        "type": "text",
                        "text": (
                            "Внутренний план:\n"
                            + plan
                            + "\n\n"
                            "Запрос пользователя:\n"
                            + user_text
                        ),
                    },
                    {
                        "type": "image_url",
                        "image_url": {
                            "url":
                                image_data_url,
                        },
                    },
                ]

                main_messages.append(
                    {
                        "role": "user",
                        "content": content,
                    }
                )

            else:

                main_messages.append(
                    {
                        "role": "user",
                        "content": (
                            "Внутренний план "
                            "(не показывай его):\n"
                            + plan
                            + "\n\n"
                            "Теперь дай пользователю "
                            "итоговый ответ.\n\n"
                            + user_content
                        ),
                    }
                )

            # ------------------------------------------------
            # AI
            # ------------------------------------------------

            answer = await self.ai.request(
                main_messages,
                temperature=0.7,
            )

            if not answer:

                answer = (
                    "Не удалось получить ответ."
                )

            # ------------------------------------------------
            # SAVE HISTORY
            # ------------------------------------------------

            history.append(
                {
                    "role": "user",
                    "content": user_content,
                }
            )

            history.append(
                {
                    "role": "assistant",
                    "content": answer,
                }
            )

            # Ограничиваем количество сообщений
            if len(history) > MAX_HISTORY_MESSAGES:

                history = history[
                    -MAX_HISTORY_MESSAGES:
                ]

            await self.store.save_history(
                user_id,
                history,
            )

            # ------------------------------------------------
            # REMOVE STATUS
            # ------------------------------------------------

            if status:

                try:

                    await status.delete()

                except Exception:
                    pass

            # ------------------------------------------------
            # SEND ANSWER
            # ------------------------------------------------

            await self.send_long_message(
                message,
                answer,
            )

        except asyncio.CancelledError:

            if status:

                try:

                    await status.delete()

                except Exception:
                    pass

            try:

                await message.answer(
                    "⏹ Запрос остановлен."
                )

            except Exception:
                pass

            raise

        except Exception:

            log.exception(
                "Request processing failed "
                "for user %s",
                user_id,
            )

            if status:

                try:

                    await status.delete()

                except Exception:
                    pass

            await message.answer(
                "❌ Не удалось обработать запрос. "
                "Попробуйте ещё раз."
            )

    # ========================================================
    # HISTORY PREPARE
    # ========================================================

    @staticmethod
    def prepare_history_for_ai(
        history: list[dict[str, Any]],
    ) -> list[dict[str, Any]]:

        result = []

        for item in history:

            role = item.get(
                "role"
            )

            content = item.get(
                "content"
            )

            if role not in {
                "user",
                "assistant",
            }:
                continue

            if not isinstance(
                content,
                str,
            ):
                continue

            result.append(
                {
                    "role": role,
                    "content": content,
                }
            )

        return result

    # ========================================================
    # HISTORY COMPRESSION
    # ========================================================

    async def maybe_compress_history(
        self,
        user_id: int,
        history: list[dict[str, Any]],
    ) -> list[dict[str, Any]]:

        if not history:
            return history

        full_text = "\n".join(
            str(
                item.get(
                    "content",
                    "",
                )
            )
            for item in history
        )

        words = count_words(
            full_text
        )

        if words < HISTORY_COMPRESS_WORDS:

            return history

        log.info(
            "Compressing history for %s "
            "(%s words)",
            user_id,
            words,
        )

        # Берём старую часть истории
        # для сжатия.
        split_at = max(
            2,
            len(history) // 2,
        )

        old_history = history[
            :split_at
        ]

        new_history = history[
            split_at:
        ]

        old_text = "\n\n".join(
            (
                f"{item.get('role', '')}: "
                f"{item.get('content', '')}"
            )
            for item in old_history
        )

        summary_messages = [
            {
                "role": "system",
                "content": (
                    "Сделай компактное резюме "
                    "старой истории диалога. "
                    "Сохрани важные факты, "
                    "решения, предпочтения, "
                    "незавершённые задачи "
                    "и контекст, необходимый "
                    "для продолжения разговора. "
                    "Не добавляй выдуманных фактов."
                ),
            },
            {
                "role": "user",
                "content": old_text,
            },
        ]

        try:

            summary = await self.ai.request(
                summary_messages,
                temperature=0.2,
            )

            compressed = [
                {
                    "role": "assistant",
                    "content": (
                        "[Краткое резюме "
                        "предыдущей части диалога]\n"
                        + summary
                    ),
                }
            ]

            result = (
                compressed
                + new_history
            )

            await self.store.save_history(
                user_id,
                result,
            )

            return result

        except Exception:

            log.exception(
                "History compression failed "
                "for user %s",
                user_id,
            )

            # Если сжатие не удалось,
            # просто оставляем историю.
            return history

    # ========================================================
    # SEND LONG MESSAGE
    # ========================================================

    async def send_long_message(
        self,
        message: Message,
        text: str,
    ) -> None:

        # Telegram ограничивает размер сообщения.
        max_length = 4000

        if len(text) <= max_length:

            try:

                await message.answer(
                    text
                )

            except TelegramBadRequest:

                # Если AI выдал некорректный HTML,
                # отправляем обычным текстом.
                await message.answer(
                    self.strip_html(
                        text
                    )
                )

            return

        # ----------------------------------------------------
        # SPLIT
        # ----------------------------------------------------

        chunks = []

        current = ""

        for line in text.splitlines(
            keepends=True
        ):

            if (
                len(current)
                + len(line)
                > max_length
            ):

                if current:
                    chunks.append(
                        current
                    )

                current = line

            else:

                current += line

        if current:
            chunks.append(
                current
            )

        for chunk in chunks:

            try:

                await message.answer(
                    chunk
                )

            except TelegramBadRequest:

                await message.answer(
                    self.strip_html(
                        chunk
                    )
                )

            await asyncio.sleep(
                0.05
            )

    # ========================================================
    # STRIP HTML
    # ========================================================

    @staticmethod
    def strip_html(
        text: str,
    ) -> str:

        return re.sub(
            r"<[^>]+>",
            "",
            text,
        )

    # ========================================================
    # CATCH ALL
    # ========================================================

    async def handle_unknown(
        self,
        message: Message,
    ) -> None:

        if message.from_user:

            await self.ensure_user(
                message
            )

        await message.answer(
            "❌ Этот тип сообщения "
            "пока не поддерживается."
        )

    # ========================================================
    # REGISTER HANDLERS
    # ========================================================

    def register_handlers(
        self,
    ) -> None:

        # ----------------------------------------------------
        # START
        # ----------------------------------------------------

        self.router.message.register(
            self.handle_start,
            CommandStart(),
        )

        # ----------------------------------------------------
        # SEND
        # ----------------------------------------------------

        self.router.message.register(
            self.handle_send,
            Command("send"),
        )

        # ----------------------------------------------------
        # STOP
        # ----------------------------------------------------

        self.router.callback_query.register(
            self.handle_stop,
            F.data == "stop_request",
        )

        # ----------------------------------------------------
        # DOCUMENT
        # ----------------------------------------------------

        self.router.message.register(
            self.handle_document,
            F.document,
        )

        # ----------------------------------------------------
        # PHOTO
        # ----------------------------------------------------

        self.router.message.register(
            self.handle_photo,
            F.photo,
        )

        # ----------------------------------------------------
        # UNSUPPORTED MEDIA
        # ----------------------------------------------------

        self.router.message.register(
            self.handle_unsupported_media,
            F.video
            | F.animation
            | F.audio
            | F.voice
            | F.video_note,
        )

        # ----------------------------------------------------
        # TEXT
        # ----------------------------------------------------

        self.router.message.register(
            self.handle_text,
            F.text,
        )

        # ----------------------------------------------------
        # FINAL CATCH ALL
        # ----------------------------------------------------

        self.router.message.register(
            self.handle_unknown
        )


# ============================================================
# GLOBAL APP
# ============================================================

App = BotApp


# ============================================================
# POLLING FALLBACK
# ============================================================

async def main() -> None:

    app = BotApp()

    try:

        await app.start()

        await app.bot.delete_webhook(
            drop_pending_updates=False
        )

        log.info(
            "Starting polling..."
        )

        await app.dp.start_polling(
            app.bot
        )

    finally:

        await app.stop()


# ============================================================
# ENTRY POINT
# ============================================================

if __name__ == "__main__":

    asyncio.run(
        main()
    )
