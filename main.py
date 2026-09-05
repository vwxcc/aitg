#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import asyncio
import base64
import json
import logging
import mimetypes
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
from aiogram.filters import Command, CommandStart
from aiogram.types import (
    CallbackQuery,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    Message,
)
from docx import Document
from openpyxl import load_workbook
from pypdf import PdfReader
from pptx import Presentation


# ============================================================================
# CONFIG
# ============================================================================

BOT_TOKEN = os.getenv("BOT_TOKEN", "").strip()

AI_BASE_URL = os.getenv("AI_BASE_URL", "").strip().rstrip("/")
AI_API_KEY = os.getenv("AI_API_KEY", "").strip()
AI_MODEL_ID = os.getenv("AI_MODEL_ID", "").strip()
AI_MODEL_NAME = os.getenv("AI_MODEL_NAME", AI_MODEL_ID or "AI").strip()

SYSTEM_PROMPT = os.getenv(
    "SYSTEM_PROMPT",
    "Ты полезный AI-ассистент. Отвечай точно, понятно и по существу.",
).strip()

FREE_TOKEN_LIMIT = int(
    os.getenv("FREE_TOKEN_LIMIT", "100000")
)

SUBSCRIPTION_TOKEN_LIMIT = int(
    os.getenv("SUBSCRIPTION_TOKEN_LIMIT", "0")
)

RESET_PERIOD_SECONDS = int(
    os.getenv("RESET_PERIOD_SECONDS", "21600")
)

MAX_HISTORY_MESSAGES = int(
    os.getenv("MAX_HISTORY_MESSAGES", "30")
)

HISTORY_COMPRESS_WORDS = int(
    os.getenv("HISTORY_COMPRESS_WORDS", "50000")
)

# Для 512 MB RAM / 0.1 CPU
MAX_CONCURRENT_AI_REQUESTS = max(
    1,
    int(os.getenv("MAX_CONCURRENT_AI_REQUESTS", "2")),
)

MAX_CONCURRENT_FILE_PROCESSING = max(
    1,
    int(os.getenv("MAX_CONCURRENT_FILE_PROCESSING", "2")),
)

MAX_FILE_SIZE_MB = int(
    os.getenv("MAX_FILE_SIZE_MB", "10")
)

MAX_FILE_SIZE_BYTES = (
    MAX_FILE_SIZE_MB * 1024 * 1024
)

AI_TIMEOUT_SECONDS = int(
    os.getenv("AI_TIMEOUT_SECONDS", "120")
)

DATA_DIR = Path(
    os.getenv("DATA_DIR", "data")
)

USERS_DIR = DATA_DIR / "users"
FILES_DIR = DATA_DIR / "files"

PURCHASE_USERNAME = "@PovilDurov"


# ============================================================================
# РАЗРЕШЁННЫЕ ФАЙЛЫ
# ============================================================================

# ВАЖНО:
# Только эти расширения разрешены.
# Всё остальное отклоняется ДО скачивания.

ALLOWED_EXTENSIONS = {
    ".pdf",
    ".docx",
    ".xlsx",
    ".pptx",
    ".txt",
    ".csv",
    ".jpg",
    ".jpeg",
    ".png",
    ".webp",
}

ALLOWED_MIME_TYPES = {
    "application/pdf",
    "text/plain",
    "text/csv",

    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "application/vnd.openxmlformats-officedocument.presentationml.presentation",

    "image/jpeg",
    "image/png",
    "image/webp",
}


# ============================================================================
# LOGGING
# ============================================================================

log = logging.getLogger("aitg")

logging.basicConfig(
    level=logging.INFO,
    format=(
        "%(asctime)s | "
        "%(levelname)s | "
        "%(name)s | "
        "%(message)s"
    ),
)


# ============================================================================
# ADMIN IDS
# ============================================================================

ADMIN_IDS: set[int] = set()


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

        if value.startswith("+"):
            value = value[1:]

        if value.lstrip("-").isdigit():
            result.add(int(value))

    return result


ADMIN_IDS.update(
    parse_admin_ids()
)


# ============================================================================
# HELPERS
# ============================================================================

def now_ts() -> float:
    return time.time()


def word_count(text: str) -> int:
    return len(
        re.findall(
            r"\S+",
            text or "",
        )
    )


def safe_filename(name: str) -> str:

    name = Path(
        name or "file"
    ).name

    name = re.sub(
        r"[^\w.\- ()]+",
        "_",
        name,
        flags=re.UNICODE,
    )

    return name[:180] or "file"


def truncate_text(
    text: str,
    limit: int = 120_000,
) -> str:

    if len(text) <= limit:
        return text

    return (
        text[:limit]
        + "\n\n"
        "[Текст файла сокращён из-за лимита.]"
    )


def format_seconds(seconds: int) -> str:

    seconds = max(
        0,
        int(seconds),
    )

    if seconds < 60:
        return f"{seconds} сек."

    minutes = seconds // 60

    if minutes < 60:
        return f"{minutes} мин."

    hours = minutes // 60
    rest = minutes % 60

    if rest:
        return f"{hours} ч. {rest} мин."

    return f"{hours} ч."


def estimate_tokens(text: str) -> int:

    if not text:
        return 1

    return max(
        1,
        len(text) // 4,
    )


def extract_usage(
    data: dict[str, Any],
) -> tuple[int, int]:

    usage = data.get(
        "usage"
    ) or {}

    input_tokens = int(
        usage.get("prompt_tokens")
        or usage.get("input_tokens")
        or 0
    )

    output_tokens = int(
        usage.get("completion_tokens")
        or usage.get("output_tokens")
        or 0
    )

    return (
        input_tokens,
        output_tokens,
    )


# ============================================================================
# JSON STORAGE
# ============================================================================

class UserStore:

    def __init__(self):

        USERS_DIR.mkdir(
            parents=True,
            exist_ok=True,
        )

        FILES_DIR.mkdir(
            parents=True,
            exist_ok=True,
        )

        self.global_lock = asyncio.Lock()

    def path(
        self,
        user_id: int,
    ) -> Path:

        return (
            USERS_DIR
            / f"{user_id}.json"
        )

    def default_state(
        self,
        user_id: int,
    ) -> dict[str, Any]:

        return {
            "user_id": user_id,
            "created_at": now_ts(),
            "updated_at": now_ts(),

            "subscription_until": 0,

            "used_tokens": 0,
            "period_started_at": now_ts(),

            "history": [],
            "compressed_summary": "",
        }

    def read_sync(
        self,
        user_id: int,
    ) -> dict[str, Any]:

        path = self.path(
            user_id
        )

        if not path.exists():
            return self.default_state(
                user_id
            )

        try:

            data = json.loads(
                path.read_text(
                    encoding="utf-8"
                )
            )

            if not isinstance(
                data,
                dict,
            ):
                return self.default_state(
                    user_id
                )

            state = self.default_state(
                user_id
            )

            state.update(
                data
            )

            if not isinstance(
                state.get("history"),
                list,
            ):
                state["history"] = []

            return state

        except Exception:

            log.exception(
                "Ошибка чтения %s",
                path,
            )

            return self.default_state(
                user_id
            )

    def write_sync(
        self,
        user_id: int,
        state: dict[str, Any],
    ):

        path = self.path(
            user_id
        )

        temp_path = path.with_suffix(
            ".tmp"
        )

        state["updated_at"] = now_ts()

        temp_path.write_text(
            json.dumps(
                state,
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )

        temp_path.replace(
            path
        )

    async def get(
        self,
        user_id: int,
    ) -> dict[str, Any]:

        async with self.global_lock:

            return await asyncio.to_thread(
                self.read_sync,
                user_id,
            )

    async def save(
        self,
        user_id: int,
        state: dict[str, Any],
    ):

        async with self.global_lock:

            await asyncio.to_thread(
                self.write_sync,
                user_id,
                state,
            )

    async def get_or_create(
        self,
        user_id: int,
    ) -> dict[str, Any]:

        state = await self.get(
            user_id
        )

        await self.save(
            user_id,
            state,
        )

        return state

    async def all_user_ids(self) -> list[int]:

        result = []

        for path in USERS_DIR.glob(
            "*.json"
        ):

            try:
                result.append(
                    int(path.stem)
                )
            except ValueError:
                pass

        return result


# ============================================================================
# LIMITS
# ============================================================================

class Limits:

    def __init__(
        self,
        store: UserStore,
    ):

        self.store = store

    @staticmethod
    def has_subscription(
        state: dict[str, Any],
    ) -> bool:

        return (
            float(
                state.get(
                    "subscription_until",
                    0,
                )
            )
            > now_ts()
        )

    async def refresh_period(
        self,
        user_id: int,
        state: dict[str, Any],
    ) -> dict[str, Any]:

        started = float(
            state.get(
                "period_started_at",
                0,
            )
        )

        current = now_ts()

        if started <= 0:

            state["period_started_at"] = current
            state["used_tokens"] = 0

            await self.store.save(
                user_id,
                state,
            )

            return state

        if (
            current - started
            >= RESET_PERIOD_SECONDS
        ):

            state["period_started_at"] = current
            state["used_tokens"] = 0

            await self.store.save(
                user_id,
                state,
            )

        return state

    async def allowed(
        self,
        user_id: int,
    ) -> tuple[bool, int]:

        state = await self.store.get_or_create(
            user_id
        )

        state = await self.refresh_period(
            user_id,
            state,
        )

        subscribed = self.has_subscription(
            state
        )

        limit = (
            SUBSCRIPTION_TOKEN_LIMIT
            if subscribed
            else FREE_TOKEN_LIMIT
        )

        used = int(
            state.get(
                "used_tokens",
                0,
            )
        )

        if limit <= 0:
            return True, 0

        if used >= limit:

            reset_at = int(
                float(
                    state.get(
                        "period_started_at",
                        now_ts(),
                    )
                )
                + RESET_PERIOD_SECONDS
            )

            return False, reset_at

        return True, 0

    async def add_tokens(
        self,
        user_id: int,
        amount: int,
    ):

        if amount <= 0:
            return

        state = await self.store.get_or_create(
            user_id
        )

        state["used_tokens"] = (
            int(
                state.get(
                    "used_tokens",
                    0,
                )
            )
            + amount
        )

        await self.store.save(
            user_id,
            state,
        )


# ============================================================================
# FILE PROCESSOR
# ============================================================================

class UnsupportedFile(Exception):
    pass


class FileTooLarge(Exception):
    pass


class FileProcessor:

    def __init__(self):

        # Обработка файлов отдельная от AI.
        self.semaphore = asyncio.Semaphore(
            MAX_CONCURRENT_FILE_PROCESSING
        )

    @staticmethod
    def allowed(
        name: str,
        mime: str = "",
    ) -> bool:

        extension = Path(
            name or ""
        ).suffix.lower()

        # Расширение является главным фильтром.
        if extension in ALLOWED_EXTENSIONS:
            return True

        # Если расширение отсутствует,
        # разрешаем только известные MIME.
        if (
            not extension
            and mime.lower() in ALLOWED_MIME_TYPES
        ):
            return True

        return False

    @staticmethod
    def is_image(
        name: str,
        mime: str,
    ) -> bool:

        extension = Path(
            name or ""
        ).suffix.lower()

        return (
            extension
            in {
                ".jpg",
                ".jpeg",
                ".png",
                ".webp",
            }
            or mime.startswith(
                "image/"
            )
        )

    async def download(
        self,
        bot: Bot,
        file_id: str,
        original_name: str,
        file_size: Optional[int],
    ) -> Path:

        if (
            file_size is not None
            and file_size > MAX_FILE_SIZE_BYTES
        ):
            raise FileTooLarge

        filename = safe_filename(
            original_name
        )

        path = (
            FILES_DIR
            / f"{uuid.uuid4().hex}_{filename}"
        )

        telegram_file = await bot.get_file(
            file_id
        )

        await bot.download_file(
            telegram_file.file_path,
            destination=path,
        )

        actual_size = path.stat().st_size

        if actual_size > MAX_FILE_SIZE_BYTES:

            path.unlink(
                missing_ok=True
            )

            raise FileTooLarge

        return path

    async def process(
        self,
        path: Path,
        original_name: str,
        mime: str,
    ) -> dict[str, Any]:

        async with self.semaphore:

            try:

                return await asyncio.wait_for(
                    self._process_inner(
                        path,
                        original_name,
                        mime,
                    ),
                    timeout=AI_TIMEOUT_SECONDS,
                )

            finally:

                path.unlink(
                    missing_ok=True
                )

    async def _process_inner(
        self,
        path: Path,
        original_name: str,
        mime: str,
    ) -> dict[str, Any]:

        extension = Path(
            original_name
        ).suffix.lower()

        # Повторная проверка.
        if not self.allowed(
            original_name,
            mime,
        ):
            raise UnsupportedFile

        # TXT / CSV
        if extension in {
            ".txt",
            ".csv",
        }:

            text = await asyncio.to_thread(
                path.read_text,
                encoding="utf-8",
                errors="replace",
            )

            return {
                "kind": "text",
                "name": original_name,
                "text": truncate_text(
                    text
                ),
            }

        # PDF
        if extension == ".pdf":

            text = await asyncio.to_thread(
                self.read_pdf,
                path,
            )

            return {
                "kind": "text",
                "name": original_name,
                "text": truncate_text(
                    text
                ),
            }

        # DOCX
        if extension == ".docx":

            text = await asyncio.to_thread(
                self.read_docx,
                path,
            )

            return {
                "kind": "text",
                "name": original_name,
                "text": truncate_text(
                    text
                ),
            }

        # XLSX
        if extension == ".xlsx":

            text = await asyncio.to_thread(
                self.read_xlsx,
                path,
            )

            return {
                "kind": "text",
                "name": original_name,
                "text": truncate_text(
                    text
                ),
            }

        # PPTX
        if extension == ".pptx":

            text = await asyncio.to_thread(
                self.read_pptx,
                path,
            )

            return {
                "kind": "text",
                "name": original_name,
                "text": truncate_text(
                    text
                ),
            }

        # IMAGE
        if self.is_image(
            original_name,
            mime,
        ):

            raw = await asyncio.to_thread(
                path.read_bytes
            )

            encoded = base64.b64encode(
                raw
            ).decode(
                "ascii"
            )

            image_mime = mime

            if not image_mime.startswith(
                "image/"
            ):
                image_mime = (
                    mimetypes.guess_type(
                        original_name
                    )[0]
                    or "image/jpeg"
                )

            return {
                "kind": "image",
                "name": original_name,
                "data_url": (
                    f"data:{image_mime};"
                    f"base64,{encoded}"
                ),
            }

        raise UnsupportedFile

    @staticmethod
    def read_pdf(
        path: Path,
    ) -> str:

        reader = PdfReader(
            str(path)
        )

        result = []

        for index, page in enumerate(
            reader.pages,
            start=1,
        ):

            try:

                text = (
                    page.extract_text()
                    or ""
                )

            except Exception as exc:

                text = (
                    f"[Ошибка чтения страницы "
                    f"{index}: {exc}]"
                )

            if text.strip():

                result.append(
                    f"--- Страница {index} ---\n"
                    f"{text}"
                )

        return "\n\n".join(
            result
        )

    @staticmethod
    def read_docx(
        path: Path,
    ) -> str:

        document = Document(
            str(path)
        )

        result = []

        for paragraph in document.paragraphs:

            text = paragraph.text.strip()

            if text:
                result.append(text)

        for table in document.tables:

            for row in table.rows:

                result.append(
                    " | ".join(
                        cell.text
                        for cell in row.cells
                    )
                )

        return "\n".join(
            result
        )

    @staticmethod
    def read_xlsx(
        path: Path,
    ) -> str:

        workbook = load_workbook(
            str(path),
            read_only=True,
            data_only=True,
        )

        result = []

        try:

            for sheet in workbook.worksheets:

                result.append(
                    f"--- Лист: {sheet.title} ---"
                )

                for row in sheet.iter_rows(
                    values_only=True
                ):

                    values = [
                        ""
                        if value is None
                        else str(value)
                        for value in row
                    ]

                    if any(
                        value.strip()
                        for value in values
                    ):

                        result.append(
                            " | ".join(values)
                        )

        finally:

            workbook.close()

        return "\n".join(
            result
        )

    @staticmethod
    def read_pptx(
        path: Path,
    ) -> str:

        presentation = Presentation(
            str(path)
        )

        result = []

        for slide_number, slide in enumerate(
            presentation.slides,
            start=1,
        ):

            slide_text = []

            for shape in slide.shapes:

                if hasattr(
                    shape,
                    "text",
                ):

                    text = (
                        shape.text
                        or ""
                    ).strip()

                    if text:
                        slide_text.append(
                            text
                        )

            if slide_text:

                result.append(
                    f"--- Слайд {slide_number} ---\n"
                    + "\n".join(
                        slide_text
                    )
                )

        return "\n\n".join(
            result
        )


# ============================================================================
# AI SERVICE
# ============================================================================

class AIError(Exception):
    pass


class AIService:

    def __init__(self):

        self.session: Optional[
            aiohttp.ClientSession
        ] = None

    async def start(self):

        timeout = aiohttp.ClientTimeout(
            total=AI_TIMEOUT_SECONDS
        )

        self.session = aiohttp.ClientSession(
            timeout=timeout,
            headers={
                "Authorization": (
                    f"Bearer {AI_API_KEY}"
                ),
                "Content-Type": (
                    "application/json"
                ),
            },
        )

    async def close(self):

        if self.session:

            await self.session.close()

            self.session = None

    def system_prompt(self) -> str:

        return (
            f"Ты — {AI_MODEL_NAME}, "
            "AI-модель, работающая "
            "внутри Telegram-бота.\n\n"
            f"{SYSTEM_PROMPT}"
        )

    async def request(
        self,
        messages: list[dict[str, Any]],
        *,
        temperature: Optional[float] = None,
        max_tokens: Optional[int] = None,
    ) -> tuple[str, dict[str, Any]]:

        if not self.session:
            raise AIError(
                "AI service не запущен."
            )

        # ------------------------------------------------------------
        # SYSTEM ОБЯЗАТЕЛЬНО ПЕРВЫМ
        # ------------------------------------------------------------

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

        final_messages = []

        if system_messages:

            final_messages.append(
                system_messages[0]
            )

        else:

            final_messages.append({
                "role": "system",
                "content": self.system_prompt(),
            })

        final_messages.extend(
            other_messages
        )

        payload: dict[str, Any] = {
            "model": AI_MODEL_ID,
            "messages": final_messages,
        }

        if temperature is not None:
            payload["temperature"] = temperature

        if max_tokens is not None:
            payload["max_tokens"] = max_tokens

        url = (
            f"{AI_BASE_URL}"
            "/chat/completions"
        )

        try:

            async with self.session.post(
                url,
                json=payload,
            ) as response:

                raw = await response.text()

                if response.status >= 400:

                    raise AIError(
                        f"AI API error "
                        f"{response.status}: "
                        f"{raw[:3000]}"
                    )

                try:

                    data = json.loads(
                        raw
                    )

                except json.JSONDecodeError:

                    raise AIError(
                        "AI API вернул "
                        "некорректный JSON."
                    )

                choices = (
                    data.get(
                        "choices"
                    )
                    or []
                )

                if not choices:

                    raise AIError(
                        "AI API не вернул choices."
                    )

                assistant_message = (
                    choices[0].get(
                        "message"
                    )
                    or {}
                )

                content = (
                    assistant_message.get(
                        "content"
                    )
                )

                if isinstance(
                    content,
                    list,
                ):

                    content = "".join(
                        part.get(
                            "text",
                            "",
                        )
                        for part in content
                        if isinstance(
                            part,
                            dict,
                        )
                    )

                if not isinstance(
                    content,
                    str,
                ):

                    content = str(
                        content or ""
                    )

                return (
                    content.strip(),
                    data,
                )

        except asyncio.TimeoutError as exc:

            raise AIError(
                "AI API превысил таймаут."
            ) from exc

        except aiohttp.ClientError as exc:

            raise AIError(
                f"Ошибка соединения с AI API: "
                f"{exc}"
            ) from exc

    async def hidden_plan(
        self,
        history: list[dict[str, Any]],
        user_text: str,
        file_context: str = "",
    ) -> tuple[str, dict[str, Any]]:

        messages = [
            {
                "role": "system",
                "content": (
                    self.system_prompt()
                    + "\n\n"
                    "Ты выполняешь скрытый этап "
                    "планирования. Пользователь "
                    "этот этап НЕ увидит. "
                    "Кратко определи, как лучше "
                    "решить задачу. "
                    "Не отвечай пользователю."
                ),
            }
        ]

        messages.extend(
            history
        )

        content = user_text

        if file_context:

            content += (
                "\n\nКонтекст файла:\n"
                + file_context
            )

        messages.append({
            "role": "user",
            "content": content,
        })

        return await self.request(
            messages,
            temperature=0.2,
            max_tokens=1200,
        )

    async def answer(
        self,
        history: list[dict[str, Any]],
        user_text: str,
        plan: str,
        file_context: str = "",
        image_data_url: Optional[str] = None,
    ) -> tuple[str, dict[str, Any]]:

        messages = [
            {
                "role": "system",
                "content": self.system_prompt(),
            }
        ]

        messages.extend(
            history
        )

        context = []

        if plan:

            context.append(
                "Скрытый план:\n"
                + plan
            )

        if file_context:

            context.append(
                "Текст файла:\n"
                + file_context
            )

        final_text = user_text

        if context:

            final_text += (
                "\n\n"
                + "\n\n".join(
                    context
                )
            )

        if image_data_url:

            messages.append({
                "role": "user",
                "content": [
                    {
                        "type": "text",
                        "text": final_text,
                    },
                    {
                        "type": "image_url",
                        "image_url": {
                            "url": image_data_url,
                        },
                    },
                ],
            })

        else:

            messages.append({
                "role": "user",
                "content": final_text,
            })

        return await self.request(
            messages,
            temperature=0.7,
        )

    async def compress_history(
        self,
        history: list[dict[str, Any]],
        old_summary: str,
    ) -> tuple[str, dict[str, Any]]:

        parts = []

        if old_summary:

            parts.append(
                "Предыдущая сводка:\n"
                + old_summary
            )

        for item in history:

            role = item.get(
                "role",
                "",
            )

            content = item.get(
                "content",
                "",
            )

            if not isinstance(
                content,
                str,
            ):

                content = (
                    "[мультимодальное сообщение]"
                )

            parts.append(
                f"{role}: {content}"
            )

        prompt = (
            "Сожми историю разговора "
            "для дальнейшего продолжения. "
            "Сохрани важные факты, решения, "
            "предпочтения пользователя "
            "и незавершённые задачи. "
            "Не придумывай информацию. "
            "Верни только компактную сводку.\n\n"
            + "\n".join(parts)
        )

        return await self.request(
            [
                {
                    "role": "system",
                    "content": (
                        self.system_prompt()
                        + "\n\n"
                        "Это внутреннее сжатие "
                        "истории."
                    ),
                },
                {
                    "role": "user",
                    "content": prompt,
                },
            ],
            temperature=0.1,
            max_tokens=3000,
        )


# ============================================================================
# AI QUEUE
# ============================================================================

class AIQueue:

    def __init__(
        self,
        handler,
    ):

        self.handler = handler

        self.queue: asyncio.Queue[
            tuple[int, dict[str, Any]]
        ] = asyncio.Queue()

        self.workers: list[
            asyncio.Task
        ] = []

        self.running = False

    async def start(self):

        if self.running:
            return

        self.running = True

        for index in range(
            MAX_CONCURRENT_AI_REQUESTS
        ):

            task = asyncio.create_task(
                self.worker(),
                name=f"ai-worker-{index + 1}",
            )

            self.workers.append(
                task
            )

    async def stop(self):

        self.running = False

        for task in self.workers:
            task.cancel()

        for task in self.workers:

            try:
                await task
            except asyncio.CancelledError:
                pass

        self.workers.clear()

    async def put(
        self,
        user_id: int,
        payload: dict[str, Any],
    ):

        await self.queue.put(
            (
                user_id,
                payload,
            )
        )

    async def worker(self):

        while self.running:

            try:

                user_id, payload = (
                    await self.queue.get()
                )

            except asyncio.CancelledError:

                raise

            try:

                await self.handler(
                    user_id,
                    payload,
                )

            except asyncio.CancelledError:

                raise

            except Exception:

                log.exception(
                    "Ошибка AI job "
                    "для %s",
                    user_id,
                )

            finally:

                self.queue.task_done()


# ============================================================================
# BOT APP
# ============================================================================

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

        self.store = UserStore()
        self.limits = Limits(
            self.store
        )

        self.files = FileProcessor()
        self.ai = AIService()

        self.queue = AIQueue(
            self.run_ai_job
        )

        # ------------------------------------------------------------
        # Один активный запрос на одного пользователя.
        # ------------------------------------------------------------

        self.active_tasks: dict[
            int,
            asyncio.Task
        ] = {}

        # ------------------------------------------------------------
        # COMMANDS
        # ------------------------------------------------------------

        self.router.message.register(
            self.handle_start,
            CommandStart(),
        )

        self.router.message.register(
            self.handle_admin_command,
            Command(
                commands=[
                    "admin",
                    "give_subscription",
                    "set_subscription",
                    "delete_subscription",
                    "subscription",
                    "broadcast",
                ]
            ),
        )

        self.router.message.register(
            self.handle_stop_command,
            Command(
                commands=[
                    "stop"
                ]
            ),
        )

        # ------------------------------------------------------------
        # ДОКУМЕНТЫ
        # ------------------------------------------------------------

        self.router.message.register(
            self.handle_document,
            F.document,
        )

        # ------------------------------------------------------------
        # ФОТО
        # ------------------------------------------------------------

        self.router.message.register(
            self.handle_photo,
            F.photo,
        )

        # ------------------------------------------------------------
        # ВСЁ НЕПОДДЕРЖИВАЕМОЕ
        # ------------------------------------------------------------

        self.router.message.register(
            self.handle_unsupported_media,
            F.video
            | F.video_note
            | F.audio
            | F.voice
            | F.animation,
        )

        # ------------------------------------------------------------
        # ТЕКСТ
        # ------------------------------------------------------------

        self.router.message.register(
            self.handle_text,
            F.text,
        )

        # ------------------------------------------------------------
        # CALLBACKS
        # ------------------------------------------------------------

        self.router.callback_query.register(
            self.handle_callback,
        )

        self.dp.include_router(
            self.router
        )

    # =========================================================================
    # START / STOP
    # =========================================================================

    async def start(self):

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

        USERS_DIR.mkdir(
            parents=True,
            exist_ok=True,
        )

        FILES_DIR.mkdir(
            parents=True,
            exist_ok=True,
        )

        await self.ai.start()
        await self.queue.start()

        log.info(
            "Bot started: %s (%s)",
            AI_MODEL_NAME,
            AI_MODEL_ID,
        )

    async def stop(self):

        for task in list(
            self.active_tasks.values()
        ):

            task.cancel()

        for task in list(
            self.active_tasks.values()
        ):

            try:
                await task
            except asyncio.CancelledError:
                pass
            except Exception:
                pass

        self.active_tasks.clear()

        await self.queue.stop()
        await self.ai.close()

        await self.bot.session.close()

    # =========================================================================
    # BUSY USER
    # =========================================================================

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

            return False

        return True

    async def reject_if_busy(
        self,
        message: Message,
    ) -> bool:

        if not message.from_user:
            return True

        user_id = (
            message.from_user.id
        )

        if self.is_busy(
            user_id
        ):

            await message.answer(
                "⏳ Я ещё обрабатываю "
                "предыдущий запрос. "
                "Подождите немного 🙂"
            )

            return True

        return False

    def register_task(
        self,
        user_id: int,
        task: asyncio.Task,
    ):

        self.active_tasks[
            user_id
        ] = task

        def cleanup(
            finished_task: asyncio.Task,
        ):

            current = (
                self.active_tasks.get(
                    user_id
                )
            )

            if current is finished_task:

                self.active_tasks.pop(
                    user_id,
                    None,
                )

        task.add_done_callback(
            cleanup
        )

    # =========================================================================
    # START
    # =========================================================================

    async def handle_start(
        self,
        message: Message,
    ):

        if not message.from_user:
            return

        await self.store.get_or_create(
            message.from_user.id
        )

        await message.answer(
            "Привет! 👋\n\n"
            "Отправьте сообщение, "
            "и я помогу с ним.\n\n"
            "Поддерживаются файлы:\n"
            "PDF, DOCX, XLSX, PPTX, TXT, CSV, "
            "JPG, JPEG, PNG, WEBP.\n\n"
            f"Максимальный размер — "
            f"{MAX_FILE_SIZE_MB} МБ."
        )

    # =========================================================================
    # TEXT
    # =========================================================================

    async def handle_text(
        self,
        message: Message,
    ):

        if (
            not message.from_user
            or not message.text
        ):
            return

        if await self.reject_if_busy(
            message
        ):
            return

        user_id = (
            message.from_user.id
        )

        task = asyncio.create_task(
            self.enqueue_text(
                message
            )
        )

        self.register_task(
            user_id,
            task,
        )

    async def enqueue_text(
        self,
        message: Message,
    ):

        user_id = (
            message.from_user.id
        )

        await self.queue.put(
            user_id,
            {
                "message": message,
                "text": message.text or "",
                "file_context": "",
                "image_data_url": None,
            },
        )

    # =========================================================================
    # DOCUMENT
    # =========================================================================

    async def handle_document(
        self,
        message: Message,
    ):

        if (
            not message.from_user
            or not message.document
        ):
            return

        if await self.reject_if_busy(
            message
        ):
            return

        document = message.document

        filename = (
            document.file_name
            or ""
        )

        mime = (
            document.mime_type
            or ""
        )

        # ================================================================
        # КРИТИЧНО:
        # ПРОВЕРЯЕМ ТИП ДО СКАЧИВАНИЯ.
        # ================================================================

        if not FileProcessor.allowed(
            filename,
            mime,
        ):

            await message.answer(
                "❌ Этот тип файла "
                "пока не поддерживается."
            )

            return

        size = (
            document.file_size
            or 0
        )

        if size > MAX_FILE_SIZE_BYTES:

            await message.answer(
                f"❌ Файл слишком большой. "
                f"Максимум — "
                f"{MAX_FILE_SIZE_MB} МБ."
            )

            return

        user_id = (
            message.from_user.id
        )

        task = asyncio.create_task(
            self.process_document(
                message,
                document.file_id,
                filename,
                mime,
                size,
            )
        )

        self.register_task(
            user_id,
            task,
        )

    # =========================================================================
    # PHOTO
    # =========================================================================

    async def handle_photo(
        self,
        message: Message,
    ):

        if (
            not message.from_user
            or not message.photo
        ):
            return

        if await self.reject_if_busy(
            message
        ):
            return

        photo = message.photo[-1]

        size = (
            photo.file_size
            or 0
        )

        if size > MAX_FILE_SIZE_BYTES:

            await message.answer(
                f"❌ Файл слишком большой. "
                f"Максимум — "
                f"{MAX_FILE_SIZE_MB} МБ."
            )

            return

        user_id = (
            message.from_user.id
        )

        task = asyncio.create_task(
            self.process_document(
                message,
                photo.file_id,
                "photo.jpg",
                "image/jpeg",
                size,
            )
        )

        self.register_task(
            user_id,
            task,
        )

    # =========================================================================
    # FILE PROCESSING
    # =========================================================================

    async def process_document(
        self,
        message: Message,
        file_id: str,
        filename: str,
        mime: str,
        size: int,
    ):

        user_id = (
            message.from_user.id
        )

        path: Optional[Path] = None

        try:

            # ------------------------------------------------------------
            # СКАЧИВАНИЕ И ОБРАБОТКА ФАЙЛА
            # НЕ ЗАНИМАЮТ AI WORKER.
            # ------------------------------------------------------------

            path = await self.files.download(
                self.bot,
                file_id,
                filename,
                size,
            )

            result = await self.files.process(
                path,
                filename,
                mime,
            )

            path = None

            caption = (
                message.caption
                or ""
            ).strip()

            # ------------------------------------------------------------
            # IMAGE
            # ------------------------------------------------------------

            if result["kind"] == "image":

                user_text = (
                    caption
                    or "Проанализируй "
                    "прикреплённое изображение."
                )

                await self.queue.put(
                    user_id,
                    {
                        "message": message,
                        "text": user_text,
                        "file_context": "",
                        "image_data_url": (
                            result["data_url"]
                        ),
                    },
                )

                return

            # ------------------------------------------------------------
            # TEXT FILE
            # ------------------------------------------------------------

            text = (
                result.get(
                    "text",
                    "",
                )
                or ""
            ).strip()

            if not text:

                await message.answer(
                    "⚠️ В файле не удалось "
                    "найти читаемый текст."
                )

                return

            user_text = (
                caption
                or "Проанализируй "
                "прикреплённый файл."
            )

            await self.queue.put(
                user_id,
                {
                    "message": message,
                    "text": user_text,
                    "file_context": (
                        f"Файл: {filename}\n\n"
                        f"{text}"
                    ),
                    "image_data_url": None,
                },
            )

        except asyncio.CancelledError:

            if path:
                path.unlink(
                    missing_ok=True
                )

            raise

        except FileTooLarge:

            if path:
                path.unlink(
                    missing_ok=True
                )

            await message.answer(
                f"❌ Файл слишком большой. "
                f"Максимум — "
                f"{MAX_FILE_SIZE_MB} МБ."
            )

        except UnsupportedFile:

            if path:
                path.unlink(
                    missing_ok=True
                )

            await message.answer(
                "❌ Этот тип файла "
                "пока не поддерживается."
            )

        except Exception:

            if path:
                path.unlink(
                    missing_ok=True
                )

            log.exception(
                "Ошибка обработки файла "
                "для пользователя %s",
                user_id,
            )

            await message.answer(
                "❌ Не удалось обработать "
                "файл. Попробуйте другой."
            )

    # =========================================================================
    # UNSUPPORTED
    # =========================================================================

    async def handle_unsupported_media(
        self,
        message: Message,
    ):

        # НИЧЕГО НЕ СКАЧИВАЕМ.
        # Сразу отвечаем пользователю.

        await message.answer(
            "❌ Этот тип файла "
            "пока не поддерживается."
        )

    # =========================================================================
    # AI JOB
    # =========================================================================

    async def run_ai_job(
        self,
        user_id: int,
        payload: dict[str, Any],
    ):

        message: Message = (
            payload["message"]
        )

        text: str = (
            payload["text"]
        )

        file_context: str = (
            payload.get(
                "file_context",
                "",
            )
            or ""
        )

        image_data_url = payload.get(
            "image_data_url"
        )

        allowed, reset_at = (
            await self.limits.allowed(
                user_id
            )
        )

        if not allowed:

            await self.send_limit_message(
                message,
                reset_at,
            )

            return

        keyboard = InlineKeyboardMarkup(
            inline_keyboard=[
                [
                    InlineKeyboardButton(
                        text="⏹ Стоп",
                        callback_data=(
                            "stop_request"
                        ),
                    )
                ]
            ]
        )

        status_message: Optional[
            Message
        ] = None

        try:

            status_message = (
                await message.answer(
                    "🧠 Думаю...",
                    reply_markup=keyboard,
                )
            )

            state = (
                await self.store.get_or_create(
                    user_id
                )
            )

            history = list(
                state.get(
                    "history",
                    [],
                )
            )

            if MAX_HISTORY_MESSAGES > 0:

                history = history[
                    -MAX_HISTORY_MESSAGES:
                ]

            # ========================================================
            # HIDDEN PLAN
            # ========================================================

            plan, plan_data = (
                await self.ai.hidden_plan(
                    history,
                    text,
                    file_context,
                )
            )

            plan_in, plan_out = (
                extract_usage(
                    plan_data
                )
            )

            # ========================================================
            # MAIN ANSWER
            # ========================================================

            answer, answer_data = (
                await self.ai.answer(
                    history,
                    text,
                    plan,
                    file_context,
                    image_data_url,
                )
            )

            answer_in, answer_out = (
                extract_usage(
                    answer_data
                )
            )

            total_tokens = (
                plan_in
                + plan_out
                + answer_in
                + answer_out
            )

            if total_tokens <= 0:

                total_tokens = (
                    estimate_tokens(text)
                    + estimate_tokens(plan)
                    + estimate_tokens(answer)
                    + estimate_tokens(file_context)
                )

            await self.limits.add_tokens(
                user_id,
                total_tokens,
            )

            # ========================================================
            # HISTORY
            # ========================================================

            history_user_text = text

            if file_context:

                history_user_text += (
                    "\n\n"
                    "[Прикреплённый файл]\n"
                    + file_context
                )

            if image_data_url:

                history_user_text += (
                    "\n\n"
                    "[Изображение прикреплено]"
                )

            history.append({
                "role": "user",
                "content": truncate_text(
                    history_user_text,
                    80_000,
                ),
            })

            history.append({
                "role": "assistant",
                "content": answer,
            })

            state["history"] = history

            await self.store.save(
                user_id,
                state,
            )

            # ========================================================
            # COMPRESS
            # ========================================================

            await self.maybe_compress_history(
                user_id
            )

            # ========================================================
            # DELETE "THINKING..."
            # ========================================================

            if status_message:

                try:

                    await status_message.delete()

                except Exception:
                    pass

            await self.send_long_message(
                message,
                answer
                or "Не удалось получить ответ.",
            )

        except asyncio.CancelledError:

            if status_message:

                try:

                    await status_message.edit_text(
                        "⏹ Обработка остановлена."
                    )

                except Exception:
                    pass

            raise

        except AIError as exc:

            log.error(
                "AI error user=%s: %s",
                user_id,
                exc,
            )

            if status_message:

                try:

                    await status_message.edit_text(
                        "❌ Не удалось получить "
                        "ответ от AI. "
                        "Попробуйте ещё раз."
                    )

                except Exception:

                    await message.answer(
                        "❌ Не удалось получить "
                        "ответ от AI."
                    )

            else:

                await message.answer(
                    "❌ Не удалось получить "
                    "ответ от AI."
                )

        except Exception:

            log.exception(
                "Unexpected AI error "
                "user=%s",
                user_id,
            )

            if status_message:

                try:

                    await status_message.edit_text(
                        "❌ Произошла ошибка "
                        "при обработке."
                    )

                except Exception:
                    pass

            else:

                await message.answer(
                    "❌ Произошла ошибка "
                    "при обработке."
                )

    # =========================================================================
    # HISTORY COMPRESSION
    # =========================================================================

    async def maybe_compress_history(
        self,
        user_id: int,
    ):

        state = (
            await self.store.get_or_create(
                user_id
            )
        )

        history = list(
            state.get(
                "history",
                [],
            )
        )

        if not history:
            return

        total_words = 0

        for item in history:

            content = item.get(
                "content",
                "",
            )

            if isinstance(
                content,
                str,
            ):

                total_words += word_count(
                    content
                )

        if total_words < HISTORY_COMPRESS_WORDS:
            return

        keep_count = max(
            6,
            MAX_HISTORY_MESSAGES // 3,
        )

        old_history = history[
            :-keep_count
        ]

        recent_history = history[
            -keep_count:
        ]

        if not old_history:
            return

        try:

            summary, data = (
                await self.ai.compress_history(
                    old_history,
                    state.get(
                        "compressed_summary",
                        "",
                    ),
                )
            )

            if not summary:
                return

            input_tokens, output_tokens = (
                extract_usage(data)
            )

            await self.limits.add_tokens(
                user_id,
                input_tokens
                + output_tokens,
            )

            state = (
                await self.store.get_or_create(
                    user_id
                )
            )

            state[
                "compressed_summary"
            ] = truncate_text(
                summary,
                40_000,
            )

            state["history"] = (
                recent_history
            )

            await self.store.save(
                user_id,
                state,
            )

        except Exception:

            # Сжатие НЕ должно ломать чат.
            log.exception(
                "Ошибка сжатия истории "
                "user=%s",
                user_id,
            )

    # =========================================================================
    # STOP
    # =========================================================================

    async def handle_stop_command(
        self,
        message: Message,
    ):

        if not message.from_user:
            return

        user_id = (
            message.from_user.id
        )

        task = (
            self.active_tasks.get(
                user_id
            )
        )

        if not task or task.done():

            await message.answer(
                "Сейчас нечего останавливать 🙂"
            )

            return

        task.cancel()

    async def handle_callback(
        self,
        callback: CallbackQuery,
    ):

        if callback.data != "stop_request":

            await callback.answer()

            return

        user_id = (
            callback.from_user.id
        )

        task = (
            self.active_tasks.get(
                user_id
            )
        )

        if not task or task.done():

            await callback.answer(
                "Обработка уже завершена."
            )

            return

        task.cancel()

        await callback.answer(
            "Останавливаю…"
        )

    # =========================================================================
    # LIMIT
    # =========================================================================

    async def send_limit_message(
        self,
        message: Message,
        reset_at: int,
    ):

        if reset_at:

            remaining = max(
                0,
                reset_at
                - int(now_ts()),
            )

            reset_text = (
                format_seconds(
                    remaining
                )
            )

        else:

            reset_text = "позже"

        await message.answer(
            "⏳ Лимит временно исчерпан.\n\n"
            f"Он обновится через "
            f"{reset_text}.\n"
            f"Если нужна подписка — "
            f"напишите администратору "
            f"{PURCHASE_USERNAME}."
        )

    # =========================================================================
    # ADMIN
    # =========================================================================

    def is_admin(
        self,
        user_id: int,
    ) -> bool:

        return user_id in ADMIN_IDS

    async def handle_admin_command(
        self,
        message: Message,
    ):

        if not message.from_user:
            return

        if not self.is_admin(
            message.from_user.id
        ):

            await message.answer(
                "⛔ Доступ запрещён."
            )

            return

        text = (
            message.text
            or ""
        ).strip()

        parts = text.split(
            maxsplit=2
        )

        command = (
            parts[0]
            .split("@")[0]
            .lower()
        )

        # ------------------------------------------------------------
        # /admin
        # ------------------------------------------------------------

        if command == "/admin":

            await self.show_admin(
                message
            )

            return

        # ------------------------------------------------------------
        # /give_subscription
        # /set_subscription
        # ------------------------------------------------------------

        if command in {
            "/give_subscription",
            "/set_subscription",
        }:

            if len(parts) < 3:

                await message.answer(
                    "Использование:\n"
                    "/give_subscription USER_ID HOURS\n"
                    "/set_subscription USER_ID HOURS"
                )

                return

            try:

                target_id = int(
                    parts[1]
                )

                hours = float(
                    parts[2].replace(
                        ",",
                        ".",
                    )
                )

            except ValueError:

                await message.answer(
                    "❌ Неверный ID "
                    "или количество часов."
                )

                return

            if hours <= 0:

                await message.answer(
                    "❌ Количество часов "
                    "должно быть больше нуля."
                )

                return

            state = (
                await self.store.get_or_create(
                    target_id
                )
            )

            if command == (
                "/give_subscription"
            ):

                current = max(
                    now_ts(),
                    float(
                        state.get(
                            "subscription_until",
                            0,
                        )
                    ),
                )

                state[
                    "subscription_until"
                ] = (
                    current
                    + hours * 3600
                )

            else:

                state[
                    "subscription_until"
                ] = (
                    now_ts()
                    + hours * 3600
                )

            await self.store.save(
                target_id,
                state,
            )

            await message.answer(
                f"✅ Подписка пользователя "
                f"<code>{target_id}</code> "
                f"изменена.\n\n"
                f"До: <b>"
                f"{time.strftime('%d.%m.%Y %H:%M', time.localtime(state['subscription_until']))}"
                f"</b>"
            )

            return

        # ------------------------------------------------------------
        # /delete_subscription
        # ------------------------------------------------------------

        if command == (
            "/delete_subscription"
        ):

            if len(parts) < 2:

                await message.answer(
                    "Использование:\n"
                    "/delete_subscription USER_ID"
                )

                return

            try:

                target_id = int(
                    parts[1]
                )

            except ValueError:

                await message.answer(
                    "❌ Неверный ID."
                )

                return

            state = (
                await self.store.get_or_create(
                    target_id
                )
            )

            state[
                "subscription_until"
            ] = 0

            await self.store.save(
                target_id,
                state,
            )

            await message.answer(
                f"✅ Подписка пользователя "
                f"<code>{target_id}</code> "
                f"удалена."
            )

            return

        # ------------------------------------------------------------
        # /subscription
        # ------------------------------------------------------------

        if command == "/subscription":

            if len(parts) < 2:

                await message.answer(
                    "Использование:\n"
                    "/subscription USER_ID"
                )

                return

            try:

                target_id = int(
                    parts[1]
                )

            except ValueError:

                await message.answer(
                    "❌ Неверный ID."
                )

                return

            await message.answer(
                await self.subscription_info(
                    target_id
                )
            )

            return

        # ------------------------------------------------------------
        # /broadcast
        # ------------------------------------------------------------

        if command == "/broadcast":

            if len(parts) < 2:

                await message.answer(
                    "Использование:\n"
                    "/broadcast ТЕКСТ"
                )

                return

            broadcast_text = (
                text.split(
                    maxsplit=1
                )[1].strip()
            )

            await self.broadcast(
                message,
                broadcast_text,
            )

    async def show_admin(
        self,
        message: Message,
    ):

        keyboard = InlineKeyboardMarkup(
            inline_keyboard=[
                [
                    InlineKeyboardButton(
                        text="➕ Выдать / изменить",
                        callback_data=(
                            "admin_sub"
                        ),
                    )
                ],
                [
                    InlineKeyboardButton(
                        text="🗑 Удалить подписку",
                        callback_data=(
                            "admin_delete"
                        ),
                    )
                ],
                [
                    InlineKeyboardButton(
                        text="ℹ️ Информация",
                        callback_data=(
                            "admin_info"
                        ),
                    )
                ],
                [
                    InlineKeyboardButton(
                        text="📢 Рассылка",
                        callback_data=(
                            "admin_broadcast"
                        ),
                    )
                ],
            ]
        )

        await message.answer(
            "👑 <b>Админ-панель</b>\n\n"
            "Команды:\n\n"
            "<code>/give_subscription USER_ID HOURS</code>\n"
            "<code>/set_subscription USER_ID HOURS</code>\n"
            "<code>/delete_subscription USER_ID</code>\n"
            "<code>/subscription USER_ID</code>\n"
            "<code>/broadcast ТЕКСТ</code>",
            reply_markup=keyboard,
        )

    async def subscription_info(
        self,
        user_id: int,
    ) -> str:

        state = (
            await self.store.get_or_create(
                user_id
            )
        )

        until = float(
            state.get(
                "subscription_until",
                0,
            )
        )

        if until > now_ts():

            subscription = time.strftime(
                "%d.%m.%Y %H:%M",
                time.localtime(
                    until
                ),
            )

        else:

            subscription = "нет"

        return (
            f"👤 ID: <code>{user_id}</code>\n"
            f"💎 Подписка до: "
            f"<b>{subscription}</b>\n"
            f"🪙 Использовано токенов: "
            f"<b>{int(state.get('used_tokens', 0))}</b>"
        )

    async def broadcast(
        self,
        admin_message: Message,
        text: str,
    ):

        user_ids = (
            await self.store.all_user_ids()
        )

        if not user_ids:

            await admin_message.answer(
                "📢 Пользователей пока нет."
            )

            return

        sent = 0
        failed = 0

        for user_id in user_ids:

            try:

                await self.bot.send_message(
                    user_id,
                    text,
                )

                sent += 1

            except Exception:

                failed += 1

            await asyncio.sleep(
                0.05
            )

        await admin_message.answer(
            f"📢 Рассылка завершена.\n"
            f"Отправлено: {sent}\n"
            f"Ошибок: {failed}"
        )

    # =========================================================================
    # SEND LONG MESSAGE
    # =========================================================================

    async def send_long_message(
        self,
        message: Message,
        text: str,
    ):

        text = (
            text
            or "Пустой ответ."
        )

        limit = 3900

        chunks = []

        while len(text) > limit:

            split_at = text.rfind(
                "\n",
                0,
                limit,
            )

            if split_at < limit // 2:

                split_at = text.rfind(
                    " ",
                    0,
                    limit,
                )

            if split_at < limit // 2:

                split_at = limit

            chunks.append(
                text[:split_at]
            )

            text = text[
                split_at:
            ].lstrip()

        if text:
            chunks.append(
                text
            )

        for chunk in chunks:

            try:

                await message.answer(
                    chunk
                )

            except Exception:

                # Если модель прислала
                # некорректный HTML.
                plain = re.sub(
                    r"<[^>]+>",
                    "",
                    chunk,
                )

                await message.answer(
                    plain
                )


# ============================================================================
# ALIAS ДЛЯ render_start.py
# ============================================================================

App = BotApp


# ============================================================================
# POLLING MODE
# ============================================================================

async def main():

    app = BotApp()

    await app.start()

    try:

        await app.dp.start_polling(
            app.bot,
            allowed_updates=(
                app.dp.resolve_used_update_types()
            ),
        )

    finally:

        await app.stop()


if __name__ == "__main__":

    asyncio.run(
        main()
    )
