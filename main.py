#!/usr/bin/env python3
# -*- coding: utf-8 -*-

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

from aiohttp import ClientSession, ClientTimeout
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
    CallbackQuery,
    FSInputFile,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    Message,
)
from dotenv import load_dotenv

from docx import Document
from openpyxl import load_workbook
from pptx import Presentation
from pypdf import PdfReader


# =============================================================================
# ENV
# =============================================================================

load_dotenv()

BOT_TOKEN = os.getenv("BOT_TOKEN", "").strip()

AI_BASE_URL = os.getenv("AI_BASE_URL", "").strip().rstrip("/")
AI_API_KEY = os.getenv("AI_API_KEY", "").strip()
AI_MODEL_ID = os.getenv("AI_MODEL_ID", "").strip()
AI_MODEL_NAME = os.getenv("AI_MODEL_NAME", "Qwen 3.5 35B").strip()

SYSTEM_PROMPT = os.getenv(
    "SYSTEM_PROMPT",
    "Отвечай полезно, точно и понятно.",
).strip()

MAX_FILE_SIZE_MB = max(
    1,
    int(os.getenv("MAX_FILE_SIZE_MB", "10")),
)

MAX_CONCURRENT_AI_REQUESTS = max(
    1,
    int(os.getenv("MAX_CONCURRENT_AI_REQUESTS", "2")),
)

MAX_CONCURRENT_FILE_PROCESSING = max(
    1,
    int(os.getenv("MAX_CONCURRENT_FILE_PROCESSING", "2")),
)

AI_TIMEOUT_SECONDS = max(
    10,
    int(os.getenv("AI_TIMEOUT_SECONDS", "120")),
)

HISTORY_COMPRESS_WORDS = max(
    1000,
    int(os.getenv("HISTORY_COMPRESS_WORDS", "50000")),
)

MAX_HISTORY_MESSAGES = max(
    10,
    int(os.getenv("MAX_HISTORY_MESSAGES", "100")),
)

DATA_DIR = Path(
    os.getenv("DATA_DIR", "data")
).expanduser()

BROADCAST_DELAY = max(
    0.0,
    float(os.getenv("BROADCAST_DELAY", "0.05")),
)

BROADCAST_CONCURRENCY = max(
    1,
    int(os.getenv("BROADCAST_CONCURRENCY", "5")),
)

LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO").upper()


# =============================================================================
# ADMIN IDS
# =============================================================================

def parse_admin_ids() -> set[int]:
    raw = os.getenv("ADMIN_TELEGRAM_IDS", "")

    result: set[int] = set()

    for value in raw.replace(";", ",").replace("\n", ",").split(","):
        value = value.strip()

        if not value:
            continue

        if value.startswith("+"):
            value = value[1:].strip()

        if value.lstrip("-").isdigit():
            result.add(int(value))

    return result


ADMIN_IDS: set[int] = parse_admin_ids()


# =============================================================================
# LOGGING
# =============================================================================

logging.basicConfig(
    level=getattr(logging, LOG_LEVEL, logging.INFO),
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
)

log = logging.getLogger("aitg")


# =============================================================================
# CONSTANTS
# =============================================================================

MAX_FILE_SIZE = MAX_FILE_SIZE_MB * 1024 * 1024

SUPPORTED_EXTENSIONS = {
    ".pdf",
    ".docx",
    ".xlsx",
    ".pptx",
    ".jpg",
    ".jpeg",
    ".png",
    ".webp",
}

SUPPORTED_MIME_TYPES = {
    "application/pdf",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    "image/jpeg",
    "image/png",
    "image/webp",
}

STOP_CALLBACK = "aitg:stop"

BUSY_TEXT = (
    "⏳ Я уже обрабатываю предыдущий запрос.\n"
    "Подождите немного 🙂"
)

QUEUE_FULL_TEXT = (
    "⏳ Сейчас большая нагрузка.\n"
    "Попробуйте через несколько секунд 🙂"
)

STOPPED_TEXT = "⏹ Запрос остановлен."


# =============================================================================
# DIRECTORIES
# =============================================================================

USERS_FILE = DATA_DIR / "users.json"
HISTORY_DIR = DATA_DIR / "history"
FILES_DIR = DATA_DIR / "files"

DATA_DIR.mkdir(parents=True, exist_ok=True)
HISTORY_DIR.mkdir(parents=True, exist_ok=True)
FILES_DIR.mkdir(parents=True, exist_ok=True)


# =============================================================================
# HELPERS
# =============================================================================

def safe_json_load(path: Path, default: Any) -> Any:
    if not path.exists():
        return default

    try:
        with path.open("r", encoding="utf-8") as file:
            return json.load(file)
    except Exception:
        log.exception("Failed to read JSON: %s", path)
        return default


def safe_json_save(path: Path, data: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)

    temp_path = path.with_suffix(
        path.suffix + f".{uuid.uuid4().hex}.tmp"
    )

    try:
        with temp_path.open("w", encoding="utf-8") as file:
            json.dump(
                data,
                file,
                ensure_ascii=False,
                indent=2,
            )

        temp_path.replace(path)

    finally:
        try:
            if temp_path.exists():
                temp_path.unlink()
        except Exception:
            pass


def count_words(text: str) -> int:
    return len(re.findall(r"\S+", text or ""))


def history_words(history: list[dict[str, Any]]) -> int:
    total = 0

    for item in history:
        content = item.get("content", "")

        if isinstance(content, str):
            total += count_words(content)

        elif isinstance(content, list):
            for part in content:
                if isinstance(part, dict):
                    text = part.get("text")

                    if isinstance(text, str):
                        total += count_words(text)

    return total


def split_long_text(
    text: str,
    max_length: int = 3900,
) -> list[str]:
    if len(text) <= max_length:
        return [text]

    parts: list[str] = []

    remaining = text

    while len(remaining) > max_length:
        cut = remaining.rfind(
            "\n",
            0,
            max_length,
        )

        if cut < max_length // 2:
            cut = remaining.rfind(
                " ",
                0,
                max_length,
            )

        if cut < max_length // 2:
            cut = max_length

        parts.append(
            remaining[:cut].strip()
        )

        remaining = remaining[cut:].strip()

    if remaining:
        parts.append(remaining)

    return parts


def image_to_data_url(
    path: Path,
    mime_type: str,
) -> str:
    encoded = base64.b64encode(
        path.read_bytes()
    ).decode("ascii")

    return f"data:{mime_type};base64,{encoded}"


def get_file_extension(
    file_name: str | None,
) -> str:
    if not file_name:
        return ""

    return Path(file_name).suffix.lower()


# =============================================================================
# JSON STORAGE
# =============================================================================

class JSONStore:
    def __init__(self):
        self.users_file = USERS_FILE
        self.history_dir = HISTORY_DIR

        self.users_lock = asyncio.Lock()

    # -------------------------------------------------------------------------
    # USERS
    # -------------------------------------------------------------------------

    async def get_users(self) -> list[int]:
        async with self.users_lock:
            data = await asyncio.to_thread(
                safe_json_load,
                self.users_file,
                [],
            )

        result: list[int] = []

        if isinstance(data, list):
            for value in data:
                try:
                    result.append(int(value))
                except (TypeError, ValueError):
                    continue

        return list(dict.fromkeys(result))

    async def add_user(
        self,
        user_id: int,
    ) -> None:
        async with self.users_lock:
            users = await asyncio.to_thread(
                safe_json_load,
                self.users_file,
                [],
            )

            if not isinstance(users, list):
                users = []

            normalized: list[int] = []

            for value in users:
                try:
                    normalized.append(int(value))
                except (TypeError, ValueError):
                    pass

            if user_id not in normalized:
                normalized.append(user_id)

                await asyncio.to_thread(
                    safe_json_save,
                    self.users_file,
                    normalized,
                )

    async def remove_user(
        self,
        user_id: int,
    ) -> None:
        async with self.users_lock:
            users = await asyncio.to_thread(
                safe_json_load,
                self.users_file,
                [],
            )

            if not isinstance(users, list):
                return

            result: list[int] = []

            for value in users:
                try:
                    current = int(value)
                except (TypeError, ValueError):
                    continue

                if current != user_id:
                    result.append(current)

            await asyncio.to_thread(
                safe_json_save,
                self.users_file,
                result,
            )

    # -------------------------------------------------------------------------
    # HISTORY
    # -------------------------------------------------------------------------

    def history_path(
        self,
        user_id: int,
    ) -> Path:
        return self.history_dir / f"{user_id}.json"

    async def get_history(
        self,
        user_id: int,
    ) -> list[dict[str, Any]]:
        path = self.history_path(user_id)

        data = await asyncio.to_thread(
            safe_json_load,
            path,
            [],
        )

        if not isinstance(data, list):
            return []

        result: list[dict[str, Any]] = []

        for item in data:
            if isinstance(item, dict):
                role = item.get("role")
                content = item.get("content")

                if role in {
                    "system",
                    "user",
                    "assistant",
                }:
                    result.append(
                        {
                            "role": role,
                            "content": content,
                        }
                    )

        return result

    async def save_history(
        self,
        user_id: int,
        history: list[dict[str, Any]],
    ) -> None:
        path = self.history_path(user_id)

        await asyncio.to_thread(
            safe_json_save,
            path,
            history,
        )


# =============================================================================
# FILE PROCESSOR
# =============================================================================

class FileProcessor:
    def __init__(self):
        self.semaphore = asyncio.Semaphore(
            MAX_CONCURRENT_FILE_PROCESSING
        )

    async def download_file(
        self,
        bot: Bot,
        telegram_file_id: str,
        original_name: str,
    ) -> Path:
        extension = get_file_extension(
            original_name
        )

        if extension not in SUPPORTED_EXTENSIONS:
            raise ValueError(
                "Этот тип файла не поддерживается."
            )

        path = FILES_DIR / (
            f"{uuid.uuid4().hex}{extension}"
        )

        telegram_file = await bot.get_file(
            telegram_file_id
        )

        if not telegram_file.file_path:
            raise RuntimeError(
                "Telegram не вернул путь к файлу."
            )

        await bot.download_file(
            telegram_file.file_path,
            destination=path,
        )

        try:
            size = path.stat().st_size
        except FileNotFoundError:
            raise RuntimeError(
                "Файл не был сохранён."
            )

        if size > MAX_FILE_SIZE:
            try:
                path.unlink()
            except Exception:
                pass

            raise ValueError(
                f"Файл слишком большой. "
                f"Максимальный размер: {MAX_FILE_SIZE_MB} МБ."
            )

        return path

    async def process_path(
        self,
        path: Path,
        original_name: str,
        mime_type: str | None,
    ) -> dict[str, Any]:
        extension = path.suffix.lower()

        async with self.semaphore:
            try:
                return await asyncio.wait_for(
                    asyncio.to_thread(
                        self._process_sync,
                        path,
                        original_name,
                        mime_type,
                    ),
                    timeout=AI_TIMEOUT_SECONDS,
                )

            except asyncio.TimeoutError:
                raise RuntimeError(
                    "Обработка файла заняла слишком много времени."
                )

    def _process_sync(
        self,
        path: Path,
        original_name: str,
        mime_type: str | None,
    ) -> dict[str, Any]:

        extension = path.suffix.lower()

        # ---------------------------------------------------------------------
        # PDF
        # ---------------------------------------------------------------------

        if extension == ".pdf":
            reader = PdfReader(str(path))

            pages: list[str] = []

            for page in reader.pages:
                try:
                    text = page.extract_text() or ""
                except Exception:
                    text = ""

                if text.strip():
                    pages.append(text.strip())

            return {
                "type": "text",
                "name": original_name,
                "text": "\n\n".join(pages).strip(),
            }

        # ---------------------------------------------------------------------
        # DOCX
        # ---------------------------------------------------------------------

        if extension == ".docx":
            document = Document(str(path))

            parts: list[str] = []

            for paragraph in document.paragraphs:
                text = paragraph.text.strip()

                if text:
                    parts.append(text)

            for table in document.tables:
                for row in table.rows:
                    cells = [
                        cell.text.strip()
                        for cell in row.cells
                    ]

                    if any(cells):
                        parts.append(
                            " | ".join(cells)
                        )

            return {
                "type": "text",
                "name": original_name,
                "text": "\n".join(parts).strip(),
            }

        # ---------------------------------------------------------------------
        # XLSX
        # ---------------------------------------------------------------------

        if extension == ".xlsx":
            workbook = load_workbook(
                filename=str(path),
                read_only=True,
                data_only=True,
            )

            parts: list[str] = []

            try:
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

                        if any(
                            value.strip()
                            for value in values
                        ):
                            parts.append(
                                " | ".join(values)
                            )
            finally:
                workbook.close()

            return {
                "type": "text",
                "name": original_name,
                "text": "\n".join(parts).strip(),
            }

        # ---------------------------------------------------------------------
        # PPTX
        # ---------------------------------------------------------------------

        if extension == ".pptx":
            presentation = Presentation(
                str(path)
            )

            parts: list[str] = []

            for index, slide in enumerate(
                presentation.slides,
                start=1,
            ):
                parts.append(
                    f"Слайд {index}:"
                )

                for shape in slide.shapes:
                    if not hasattr(shape, "text"):
                        continue

                    text = (
                        shape.text
                        .strip()
                    )

                    if text:
                        parts.append(text)

            return {
                "type": "text",
                "name": original_name,
                "text": "\n".join(parts).strip(),
            }

        # ---------------------------------------------------------------------
        # IMAGES
        # ---------------------------------------------------------------------

        if extension in {
            ".jpg",
            ".jpeg",
            ".png",
            ".webp",
        }:
            if extension in {
                ".jpg",
                ".jpeg",
            }:
                detected_mime = "image/jpeg"
            elif extension == ".png":
                detected_mime = "image/png"
            else:
                detected_mime = "image/webp"

            return {
                "type": "image",
                "name": original_name,
                "mime_type": detected_mime,
                "data_url": image_to_data_url(
                    path,
                    detected_mime,
                ),
            }

        raise ValueError(
            "Этот тип файла не поддерживается."
        )


# =============================================================================
# AI SERVICE
# =============================================================================

class AIService:
    def __init__(self):
        self.session: ClientSession | None = None

        self.semaphore = asyncio.Semaphore(
            MAX_CONCURRENT_AI_REQUESTS
        )

    async def start(self) -> None:
        if self.session is not None:
            return

        timeout = ClientTimeout(
            total=AI_TIMEOUT_SECONDS,
            connect=min(
                20,
                AI_TIMEOUT_SECONDS,
            ),
        )

        self.session = ClientSession(
            timeout=timeout,
            headers={
                "Authorization": f"Bearer {AI_API_KEY}",
                "Content-Type": "application/json",
            },
        )

    async def close(self) -> None:
        if self.session is not None:
            await self.session.close()

            self.session = None

    def get_system_prompt(self) -> str:
        return (
            f"Ты — {AI_MODEL_NAME}, "
            f"AI-модель, работающая внутри Telegram-бота.\n\n"
            f"{SYSTEM_PROMPT}"
        )

    @staticmethod
    def normalize_messages(
        messages: list[dict[str, Any]],
    ) -> list[dict[str, Any]]:
        system_messages = [
            message
            for message in messages
            if message.get("role") == "system"
        ]

        other_messages = [
            message
            for message in messages
            if message.get("role") != "system"
        ]

        if system_messages:
            return [
                system_messages[0],
                *other_messages,
            ]

        return other_messages

    async def chat(
        self,
        messages: list[dict[str, Any]],
        *,
        temperature: float | None = None,
    ) -> str:

        if self.session is None:
            raise RuntimeError(
                "AI service не запущен."
            )

        normalized = self.normalize_messages(
            messages
        )

        payload: dict[str, Any] = {
            "model": AI_MODEL_ID,
            "messages": normalized,
        }

        if temperature is not None:
            payload["temperature"] = temperature

        url = f"{AI_BASE_URL}/chat/completions"

        async with self.semaphore:
            try:
                async with self.session.post(
                    url,
                    json=payload,
                ) as response:

                    raw_text = await response.text()

                    if response.status >= 400:
                        raise RuntimeError(
                            f"AI API error "
                            f"{response.status}: "
                            f"{raw_text[:4000]}"
                        )

                    try:
                        data = json.loads(
                            raw_text
                        )
                    except json.JSONDecodeError:
                        raise RuntimeError(
                            "AI API вернул некорректный JSON."
                        )

                    return self.extract_content(
                        data
                    )

            except asyncio.TimeoutError:
                raise RuntimeError(
                    "AI API не ответил вовремя."
                )

    @staticmethod
    def extract_content(
        data: dict[str, Any],
    ) -> str:

        choices = data.get("choices")

        if not isinstance(choices, list) or not choices:
            raise RuntimeError(
                "AI API не вернул choices."
            )

        choice = choices[0]

        if not isinstance(choice, dict):
            raise RuntimeError(
                "AI API вернул некорректный choice."
            )

        message = choice.get("message")

        if not isinstance(message, dict):
            raise RuntimeError(
                "AI API не вернул message."
            )

        content = message.get("content")

        if isinstance(content, str):
            return content.strip()

        if isinstance(content, list):
            parts: list[str] = []

            for part in content:
                if not isinstance(part, dict):
                    continue

                text = part.get("text")

                if isinstance(text, str):
                    parts.append(text)

            result = "\n".join(parts).strip()

            if result:
                return result

        raise RuntimeError(
            "AI API не вернул текст ответа."
        )


# =============================================================================
# QUEUE
# =============================================================================

class QueueFullError(Exception):
    pass


class QueueItem:
    def __init__(
        self,
        future: asyncio.Future,
        messages: list[dict[str, Any]],
    ):
        self.future = future
        self.messages = messages


class AIQueue:
    def __init__(
        self,
        ai: AIService,
    ):
        self.ai = ai

        self.queue: asyncio.Queue[
            QueueItem
        ] = asyncio.Queue()

        self.worker_tasks: list[
            asyncio.Task
        ] = []

        self.running = False

    async def start(self) -> None:
        if self.running:
            return

        self.running = True

        self.worker_tasks = [
            asyncio.create_task(
                self.worker_loop(),
                name=f"ai-worker-{index + 1}",
            )
            for index in range(
                MAX_CONCURRENT_AI_REQUESTS
            )
        ]

    async def stop(self) -> None:
        self.running = False

        tasks = list(
            self.worker_tasks
        )

        self.worker_tasks.clear()

        for task in tasks:
            task.cancel()

        for task in tasks:
            try:
                await task
            except asyncio.CancelledError:
                pass

    async def submit(
        self,
        messages: list[dict[str, Any]],
    ) -> str:

        if not self.running:
            raise RuntimeError(
                "Очередь AI не запущена."
            )

        loop = asyncio.get_running_loop()

        future: asyncio.Future = (
            loop.create_future()
        )

        item = QueueItem(
            future=future,
            messages=messages,
        )

        try:
            self.queue.put_nowait(item)
        except asyncio.QueueFull:
            raise QueueFullError

        try:
            return await future
        except asyncio.CancelledError:
            if not future.done():
                future.cancel()

            raise

    async def worker_loop(self) -> None:
        while self.running:
            item = await self.queue.get()

            try:
                if item.future.cancelled():
                    continue

                try:
                    result = await self.ai.chat(
                        item.messages
                    )

                    if not item.future.done():
                        item.future.set_result(
                            result
                        )

                except asyncio.CancelledError:
                    if not item.future.done():
                        item.future.cancel()

                    raise

                except Exception as exc:
                    if not item.future.done():
                        item.future.set_exception(
                            exc
                        )

            finally:
                self.queue.task_done()


# =============================================================================
# BOT APPLICATION
# =============================================================================

class BotApp:
    def __init__(self):
        self.bot = Bot(
            BOT_TOKEN,
            default=DefaultBotProperties(
                parse_mode=ParseMode.HTML
            ),
        )

        self.dp = Dispatcher()
        self.router = Router()

        self.store = JSONStore()
        self.files = FileProcessor()
        self.ai = AIService()
        self.queue = AIQueue(self.ai)

        # ---------------------------------------------------------------------
        # САМОЕ ВАЖНОЕ:
        #
        # user_id -> активная asyncio.Task
        #
        # Если пользователь уже есть здесь, его новое сообщение
        # НЕ запускается и НЕ попадает в AI.
        # ---------------------------------------------------------------------

        self.active_tasks: dict[
            int,
            asyncio.Task,
        ] = {}

        self.active_lock = asyncio.Lock()

        self.broadcast_lock = asyncio.Lock()

        self.register_handlers()

    # =========================================================================
    # HANDLERS
    # =========================================================================

    def register_handlers(self) -> None:

        self.router.message.register(
            self.handle_start,
            CommandStart(),
        )

        self.router.message.register(
            self.handle_send,
            Command("send"),
        )

        self.router.callback_query.register(
            self.handle_stop,
            F.data == STOP_CALLBACK,
        )

        self.router.message.register(
            self.handle_document,
            F.document,
        )

        self.router.message.register(
            self.handle_photo,
            F.photo,
        )

        self.router.message.register(
            self.handle_text,
            F.text,
        )

        self.router.message.register(
            self.handle_unsupported,
        )

        self.dp.include_router(
            self.router
        )

    # =========================================================================
    # LIFECYCLE
    # =========================================================================

    async def start(self) -> None:
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
            "Bot started | model=%s | model_id=%s",
            AI_MODEL_NAME,
            AI_MODEL_ID,
        )

    async def stop(self) -> None:
        # Отменяем активные пользовательские задачи.
        async with self.active_lock:
            tasks = list(
                self.active_tasks.values()
            )

        for task in tasks:
            task.cancel()

        for task in tasks:
            try:
                await task
            except asyncio.CancelledError:
                pass
            except Exception:
                pass

        await self.queue.stop()
        await self.ai.close()
        await self.bot.session.close()

    # =========================================================================
    # ACTIVE REQUEST LOCK
    # =========================================================================

    async def is_busy(
        self,
        user_id: int,
    ) -> bool:

        async with self.active_lock:
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
                return False

            return True

    async def register_task(
        self,
        user_id: int,
        task: asyncio.Task,
    ) -> bool:

        async with self.active_lock:
            current = self.active_tasks.get(
                user_id
            )

            # Защита от гонки:
            # если другой запрос уже зарегистрирован,
            # этот запрос не принимаем.
            if (
                current is not None
                and not current.done()
            ):
                return False

            self.active_tasks[user_id] = task

            return True

    async def release_task(
        self,
        user_id: int,
        task: asyncio.Task,
    ) -> None:

        async with self.active_lock:
            current = self.active_tasks.get(
                user_id
            )

            if current is task:
                self.active_tasks.pop(
                    user_id,
                    None,
                )

    # =========================================================================
    # STOP KEYBOARD
    # =========================================================================

    @staticmethod
    def stop_keyboard() -> InlineKeyboardMarkup:
        return InlineKeyboardMarkup(
            inline_keyboard=[
                [
                    InlineKeyboardButton(
                        text="⏹ Стоп",
                        callback_data=STOP_CALLBACK,
                    )
                ]
            ]
        )

    async def handle_stop(
        self,
        callback: CallbackQuery,
    ) -> None:

        if callback.from_user is None:
            await callback.answer()
            return

        user_id = callback.from_user.id

        async with self.active_lock:
            task = self.active_tasks.get(
                user_id
            )

        if task is None or task.done():
            await callback.answer(
                "Активного запроса нет."
            )
            return

        task.cancel()

        await callback.answer(
            "Запрос остановлен."
        )

        try:
            if callback.message:
                await callback.message.edit_reply_markup(
                    reply_markup=None
                )
        except Exception:
            pass

    # =========================================================================
    # START
    # =========================================================================

    async def handle_start(
        self,
        message: Message,
    ) -> None:

        if not message.from_user:
            return

        user_id = message.from_user.id

        await self.store.add_user(
            user_id
        )

        await message.answer(
            f"👋 Привет!\n\n"
            f"Я — {AI_MODEL_NAME}.\n"
            f"Отправьте мне сообщение или поддерживаемый файл."
        )

    # =========================================================================
    # ADMIN BROADCAST
    # =========================================================================

    async def handle_send(
        self,
        message: Message,
    ) -> None:

        if not message.from_user:
            return

        user_id = message.from_user.id

        if user_id not in ADMIN_IDS:
            await message.answer(
                "⛔ Доступ запрещён."
            )
            return

        text = message.text or ""

        parts = text.split(
            maxsplit=1
        )

        if len(parts) < 2:
            await message.answer(
                "Использование:\n"
                "<code>/send текст рассылки</code>"
            )
            return

        broadcast_text = parts[1].strip()

        if not broadcast_text:
            await message.answer(
                "Текст рассылки пустой."
            )
            return

        # Рассылка запускается отдельно,
        # поэтому AI-обработка пользователей
        # от неё не блокируется.
        asyncio.create_task(
            self.broadcast(
                broadcast_text,
                user_id,
            )
        )

        await message.answer(
            "📨 Рассылка запущена."
        )

    async def broadcast(
        self,
        text: str,
        admin_id: int,
    ) -> None:

        async with self.broadcast_lock:
            users = await self.store.get_users()

            total = len(users)

            if total == 0:
                try:
                    await self.bot.send_message(
                        admin_id,
                        "📨 Пользователей для рассылки нет.",
                    )
                except Exception:
                    pass

                return

            try:
                status_message = (
                    await self.bot.send_message(
                        admin_id,
                        f"📨 Рассылка запущена.\n"
                        f"Получателей: {total}",
                    )
                )
            except Exception:
                status_message = None

            semaphore = asyncio.Semaphore(
                BROADCAST_CONCURRENCY
            )

            sent = 0
            failed = 0
            removed = 0

            counter_lock = asyncio.Lock()

            async def send_one(
                target_id: int,
            ) -> None:
                nonlocal sent
                nonlocal failed
                nonlocal removed

                async with semaphore:
                    try:
                        await self.bot.send_message(
                            target_id,
                            text,
                        )

                        async with counter_lock:
                            sent += 1

                    except TelegramRetryAfter as exc:
                        try:
                            await asyncio.sleep(
                                exc.retry_after
                            )

                            await self.bot.send_message(
                                target_id,
                                text,
                            )

                            async with counter_lock:
                                sent += 1

                        except Exception:
                            async with counter_lock:
                                failed += 1

                    except (
                        TelegramForbiddenError,
                        TelegramBadRequest,
                    ):
                        await self.store.remove_user(
                            target_id
                        )

                        async with counter_lock:
                            removed += 1

                    except (
                        TelegramNetworkError,
                    ):
                        async with counter_lock:
                            failed += 1

                    except Exception:
                        log.exception(
                            "Broadcast failed for %s",
                            target_id,
                        )

                        async with counter_lock:
                            failed += 1

                    finally:
                        if BROADCAST_DELAY > 0:
                            await asyncio.sleep(
                                BROADCAST_DELAY
                            )

            batch_size = max(
                1,
                BROADCAST_CONCURRENCY * 10,
            )

            for start in range(
                0,
                len(users),
                batch_size,
            ):
                batch = users[
                    start:start + batch_size
                ]

                await asyncio.gather(
                    *(
                        send_one(user_id)
                        for user_id in batch
                    )
                )

            final_text = (
                "📨 Рассылка завершена.\n\n"
                f"Всего: {total}\n"
                f"Отправлено: {sent}\n"
                f"Ошибок: {failed}\n"
                f"Удалено заблокированных: {removed}"
            )

            if status_message is not None:
                try:
                    await status_message.edit_text(
                        final_text
                    )
                    return
                except Exception:
                    pass

            try:
                await self.bot.send_message(
                    admin_id,
                    final_text,
                )
            except Exception:
                pass

    # =========================================================================
    # COMMON REQUEST ADMISSION
    # =========================================================================

    async def begin_user_request(
        self,
        message: Message,
    ) -> Optional[asyncio.Task]:

        if not message.from_user:
            return None

        user_id = message.from_user.id

        # Пользователь уже обрабатывается.
        # Никакой очереди и никакого AI-запроса.
        if await self.is_busy(user_id):
            await message.answer(
                BUSY_TEXT
            )
            return None

        # Создаём task заранее.
        #
        # ВАЖНО:
        # register_task выполняется до того, как task получит
        # управление и до любого await внутри обработки.
        task = asyncio.create_task(
            self.process_request_wrapper(
                message
            ),
            name=f"user-request-{user_id}",
        )

        registered = await self.register_task(
            user_id,
            task,
        )

        if not registered:
            task.cancel()

            try:
                await task
            except asyncio.CancelledError:
                pass
            except Exception:
                pass

            await message.answer(
                BUSY_TEXT
            )

            return None

        return task

    # =========================================================================
    # TEXT
    # =========================================================================

    async def handle_text(
        self,
        message: Message,
    ) -> None:

        if not message.from_user:
            return

        text = (
            message.text or ""
        ).strip()

        if not text:
            return

        await self.store.add_user(
            message.from_user.id
        )

        await self.begin_user_request(
            message
        )

    # =========================================================================
    # DOCUMENT
    # =========================================================================

    async def handle_document(
        self,
        message: Message,
    ) -> None:

        if not message.from_user:
            return

        await self.store.add_user(
            message.from_user.id
        )

        document = message.document

        if document is None:
            return

        file_name = (
            document.file_name
            or "file"
        )

        extension = get_file_extension(
            file_name
        )

        if extension not in SUPPORTED_EXTENSIONS:
            await message.answer(
                "❌ Этот тип файла не поддерживается.\n\n"
                "Поддерживаются: PDF, DOCX, XLSX, PPTX, "
                "JPG, JPEG, PNG и WEBP."
            )
            return

        if (
            document.file_size is not None
            and document.file_size > MAX_FILE_SIZE
        ):
            await message.answer(
                f"❌ Файл слишком большой.\n"
                f"Максимальный размер: "
                f"{MAX_FILE_SIZE_MB} МБ."
            )
            return

        await self.begin_user_request(
            message
        )

    # =========================================================================
    # PHOTO
    # =========================================================================

    async def handle_photo(
        self,
        message: Message,
    ) -> None:

        if not message.from_user:
            return

        await self.store.add_user(
            message.from_user.id
        )

        if not message.photo:
            return

        largest = message.photo[-1]

        if (
            largest.file_size is not None
            and largest.file_size > MAX_FILE_SIZE
        ):
            await message.answer(
                f"❌ Изображение слишком большое.\n"
                f"Максимальный размер: "
                f"{MAX_FILE_SIZE_MB} МБ."
            )
            return

        await self.begin_user_request(
            message
        )

    # =========================================================================
    # UNSUPPORTED
    # =========================================================================

    async def handle_unsupported(
        self,
        message: Message,
    ) -> None:

        if not message.from_user:
            return

        await message.answer(
            "❌ Этот тип сообщения не поддерживается.\n\n"
            "Можно отправить текст, PDF, DOCX, XLSX, PPTX, "
            "JPG, JPEG, PNG или WEBP."
        )

    # =========================================================================
    # MAIN REQUEST
    # =========================================================================

    async def process_request_wrapper(
        self,
        message: Message,
    ) -> None:

        if not message.from_user:
            return

        user_id = message.from_user.id

        current_task = asyncio.current_task()

        try:
            await self.process_request(
                message
            )

        except asyncio.CancelledError:
            try:
                await message.answer(
                    STOPPED_TEXT
                )
            except Exception:
                pass

            raise

        except QueueFullError:
            try:
                await message.answer(
                    QUEUE_FULL_TEXT
                )
            except Exception:
                pass

        except Exception:
            log.exception(
                "Request failed for user %s",
                user_id,
            )

            try:
                await message.answer(
                    "❌ Произошла ошибка при обработке запроса."
                )
            except Exception:
                pass

        finally:
            # КРИТИЧЕСКИ ВАЖНО:
            #
            # Только здесь пользователь снова становится свободным.
            #
            # Пока этот finally не выполнен,
            # любое новое сообщение получает BUSY_TEXT.
            await self.release_task(
                user_id,
                current_task,
            )

    async def process_request(
        self,
        message: Message,
    ) -> None:

        if not message.from_user:
            return

        user_id = message.from_user.id

        user_text = (
            message.text or ""
        ).strip()

        files_context: list[dict[str, Any]] = []

        # ---------------------------------------------------------------------
        # FILE
        # ---------------------------------------------------------------------

        if message.document is not None:
            document = message.document

            file_name = (
                document.file_name
                or "file"
            )

            extension = get_file_extension(
                file_name
            )

            if extension not in SUPPORTED_EXTENSIONS:
                raise ValueError(
                    "Этот тип файла не поддерживается."
                )

            temp_path: Path | None = None

            try:
                temp_path = await self.files.download_file(
                    self.bot,
                    document.file_id,
                    file_name,
                )

                processed = (
                    await self.files.process_path(
                        temp_path,
                        file_name,
                        document.mime_type,
                    )
                )

                files_context.append(
                    processed
                )

            finally:
                if temp_path is not None:
                    try:
                        temp_path.unlink(
                            missing_ok=True
                        )
                    except Exception:
                        pass

        # ---------------------------------------------------------------------
        # PHOTO
        # ---------------------------------------------------------------------

        elif message.photo:
            photo = message.photo[-1]

            temp_path: Path | None = None

            try:
                temp_path = await self.files.download_file(
                    self.bot,
                    photo.file_id,
                    "image.jpg",
                )

                processed = (
                    await self.files.process_path(
                        temp_path,
                        "image.jpg",
                        "image/jpeg",
                    )
                )

                files_context.append(
                    processed
                )

            finally:
                if temp_path is not None:
                    try:
                        temp_path.unlink(
                            missing_ok=True
                        )
                    except Exception:
                        pass

        # ---------------------------------------------------------------------
        # HISTORY
        # ---------------------------------------------------------------------

        history = await self.store.get_history(
            user_id
        )

        # Защита от повреждённой/слишком большой истории.
        if len(history) > MAX_HISTORY_MESSAGES:
            history = history[
                -MAX_HISTORY_MESSAGES:
            ]

        # ---------------------------------------------------------------------
        # STATUS
        # ---------------------------------------------------------------------

        status_message = await message.answer(
            "🤔 Обрабатываю запрос...",
            reply_markup=self.stop_keyboard(),
        )

        # ---------------------------------------------------------------------
        # BUILD USER MESSAGE
        # ---------------------------------------------------------------------

        main_user_content: Any = user_text

        text_parts: list[str] = []

        if user_text:
            text_parts.append(
                user_text
            )

        for file_data in files_context:
            if file_data.get("type") == "text":
                file_name = file_data.get(
                    "name",
                    "файл",
                )

                file_text = file_data.get(
                    "text",
                    "",
                )

                if file_text:
                    text_parts.append(
                        f"\n\n"
                        f"--- Файл: {file_name} ---\n"
                        f"{file_text}\n"
                        f"--- Конец файла ---"
                    )
                else:
                    text_parts.append(
                        f"\n\n"
                        f"--- Файл: {file_name} ---\n"
                        f"[В файле не удалось извлечь текст]\n"
                        f"--- Конец файла ---"
                    )

        text_content = "\n".join(
            text_parts
        ).strip()

        image_parts: list[dict[str, Any]] = []

        for file_data in files_context:
            if file_data.get("type") != "image":
                continue

            data_url = file_data.get(
                "data_url"
            )

            if not isinstance(
                data_url,
                str,
            ):
                continue

            image_parts.append(
                {
                    "type": "image_url",
                    "image_url": {
                        "url": data_url
                    },
                }
            )

        # ---------------------------------------------------------------------
        # MULTIMODAL USER CONTENT
        # ---------------------------------------------------------------------

        if image_parts:
            content_parts: list[dict[str, Any]] = []

            if text_content:
                content_parts.append(
                    {
                        "type": "text",
                        "text": text_content,
                    }
                )

            content_parts.extend(
                image_parts
            )

            main_user_content = (
                content_parts
            )

        else:
            main_user_content = (
                text_content
                or "Проанализируй прикреплённый файл."
            )

        # ---------------------------------------------------------------------
        # STAGE 1 — HIDDEN PLAN
        # ---------------------------------------------------------------------

        plan_messages: list[
            dict[str, Any]
        ] = [
            {
                "role": "system",
                "content": (
                    self.ai.get_system_prompt()
                    + "\n\n"
                    "Ты сейчас выполняешь внутренний этап подготовки ответа. "
                    "Пользователь не увидит этот этап. "
                    "Составь краткий план того, как лучше ответить "
                    "на последний запрос. Не отвечай пользователю напрямую."
                ),
            }
        ]

        # Для плана берём историю без system.
        for item in history:
            if item.get("role") in {
                "user",
                "assistant",
            }:
                plan_messages.append(
                    item
                )

        plan_messages.append(
            {
                "role": "user",
                "content": main_user_content,
            }
        )

        plan = await self.queue.submit(
            plan_messages
        )

        # ---------------------------------------------------------------------
        # STAGE 2 — MAIN ANSWER
        # ---------------------------------------------------------------------

        main_messages: list[
            dict[str, Any]
        ] = [
            {
                "role": "system",
                "content": self.ai.get_system_prompt(),
            }
        ]

        for item in history:
            if item.get("role") in {
                "user",
                "assistant",
            }:
                main_messages.append(
                    item
                )

        main_messages.append(
            {
                "role": "user",
                "content": main_user_content,
            }
        )

        # План скрыт от пользователя,
        # но передаётся модели как внутренняя инструкция.
        main_messages.append(
            {
                "role": "user",
                "content": (
                    "[ВНУТРЕННЯЯ ИНСТРУКЦИЯ ДЛЯ МОДЕЛИ]\n"
                    "Используй следующий подготовленный план "
                    "для формирования ответа. "
                    "Не упоминай этот план пользователю.\n\n"
                    f"{plan}"
                ),
            }
        )

        answer = await self.queue.submit(
            main_messages
        )

        if not answer.strip():
            answer = (
                "Не удалось получить текстовый ответ."
            )

        # ---------------------------------------------------------------------
        # SAVE HISTORY
        # ---------------------------------------------------------------------

        history.append(
            {
                "role": "user",
                "content": main_user_content,
            }
        )

        history.append(
            {
                "role": "assistant",
                "content": answer,
            }
        )

        # ---------------------------------------------------------------------
        # HISTORY COMPRESSION
        # ---------------------------------------------------------------------

        if (
            history_words(history)
            >= HISTORY_COMPRESS_WORDS
        ):
            history = await self.compress_history(
                history
            )

        if len(history) > MAX_HISTORY_MESSAGES:
            history = history[
                -MAX_HISTORY_MESSAGES:
            ]

        await self.store.save_history(
            user_id,
            history,
        )

        # ---------------------------------------------------------------------
        # SEND ANSWER
        # ---------------------------------------------------------------------

        try:
            await status_message.delete()
        except Exception:
            pass

        chunks = split_long_text(
            answer
        )

        for index, chunk in enumerate(
            chunks
        ):
            if index == len(chunks) - 1:
                await message.answer(
                    chunk,
                )
            else:
                await message.answer(
                    chunk,
                )

    # =========================================================================
    # HISTORY COMPRESSION
    # =========================================================================

    async def compress_history(
        self,
        history: list[dict[str, Any]],
    ) -> list[dict[str, Any]]:

        if not history:
            return history

        compression_messages: list[
            dict[str, Any]
        ] = [
            {
                "role": "system",
                "content": (
                    self.ai.get_system_prompt()
                    + "\n\n"
                    "Ты выполняешь скрытое сжатие истории диалога. "
                    "Пользователь не увидит этот запрос. "
                    "Сделай максимально полезное краткое резюме "
                    "предыдущего диалога: сохрани факты, решения, "
                    "контекст, предпочтения пользователя, "
                    "незавершённые задачи и важные детали. "
                    "Не добавляй выдуманные сведения."
                ),
            }
        ]

        for item in history:
            role = item.get("role")

            if role not in {
                "user",
                "assistant",
            }:
                continue

            content = item.get(
                "content",
                "",
            )

            if isinstance(
                content,
                list,
            ):
                # Изображения в старую историю
                # повторно отправлять не нужно.
                text_parts = []

                for part in content:
                    if not isinstance(
                        part,
                        dict,
                    ):
                        continue

                    if part.get("type") == "text":
                        text = part.get(
                            "text",
                            "",
                        )

                        if isinstance(
                            text,
                            str,
                        ):
                            text_parts.append(
                                text
                            )

                content = "\n".join(
                    text_parts
                )

            if not isinstance(
                content,
                str,
            ):
                content = str(content)

            compression_messages.append(
                {
                    "role": role,
                    "content": content,
                }
            )

        summary = await self.queue.submit(
            compression_messages
        )

        if not summary.strip():
            return history

        # Сохраняем только краткое резюме.
        # Оно хранится как assistant, а не system,
        # чтобы не возникала ошибка провайдера:
        # "System message must be at the beginning."
        return [
            {
                "role": "assistant",
                "content": (
                    "[КРАТКОЕ РЕЗЮМЕ ПРЕДЫДУЩЕГО ДИАЛОГА]\n"
                    + summary.strip()
                ),
            }
        ]

    # =========================================================================
    # CLEANUP STATUS MESSAGE
    # =========================================================================

    async def safe_remove_stop_button(
        self,
        message: Message,
    ) -> None:

        try:
            await message.edit_reply_markup(
                reply_markup=None
            )
        except Exception:
            pass


# =============================================================================
# GLOBAL APP
# =============================================================================

App = BotApp()


# =============================================================================
# POLLING MODE
# =============================================================================

async def main() -> None:
    await App.start()

    try:
        await App.dp.start_polling(
            App.bot
        )
    finally:
        await App.stop()


# =============================================================================
# ENTRY POINT
# =============================================================================

if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        pass
