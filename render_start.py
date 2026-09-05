#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import asyncio
import hashlib
import os
import time

from aiohttp import web
from aiogram.types import Update

import main as bot_main


# ============================================================
# CONFIG
# ============================================================

HOST = "0.0.0.0"
PORT = int(os.getenv("PORT", "10000"))

# Максимальный размер HTTP-запроса.
# Telegram webhook обычно намного меньше.
MAX_HTTP_BODY = 50 * 1024 * 1024

# Простая защита health endpoint.
# Не влияет на Telegram webhook.
HEALTH_MIN_INTERVAL = 2.0

# Максимальное количество одновременно обрабатываемых
# webhook-запросов.
MAX_WEBHOOK_CONCURRENCY = 20


# ============================================================
# ADMIN IDS
# ============================================================

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


bot_main.ADMIN_IDS.clear()
bot_main.ADMIN_IDS.update(parse_admin_ids())


# ============================================================
# APP STARTUP
# ============================================================

async def prepare_app(app: bot_main.App) -> None:

    if not bot_main.BOT_TOKEN:
        raise RuntimeError("BOT_TOKEN не задан")

    if not bot_main.AI_BASE_URL:
        raise RuntimeError("AI_BASE_URL не задан")

    if not bot_main.AI_API_KEY:
        raise RuntimeError("AI_API_KEY не задан")

    if not bot_main.AI_MODEL_ID:
        raise RuntimeError("AI_MODEL_ID не задан")

    await app.ai.start()
    await app.queue.start()

    bot_main.log.info(
        "AI model: %s (%s)",
        bot_main.AI_MODEL_NAME,
        bot_main.AI_MODEL_ID,
    )


# ============================================================
# CLEANUP
# ============================================================

async def cleanup_app(app: bot_main.App) -> None:

    try:
        await app.queue.stop()

    finally:

        try:
            await app.ai.close()

        finally:

            await app.bot.session.close()


# ============================================================
# HTTP SERVER
# ============================================================

async def create_http_app(
    app: bot_main.App,
) -> web.Application:

    http_app = web.Application(
        client_max_size=MAX_HTTP_BODY,
    )

    # --------------------------------------------------------
    # WEBHOOK SECRET
    # --------------------------------------------------------

    secret = os.getenv(
        "WEBHOOK_SECRET",
        "",
    ).strip()

    if not secret:
        secret = hashlib.sha256(
            bot_main.BOT_TOKEN.encode("utf-8")
        ).hexdigest()[:48]

    # --------------------------------------------------------
    # WEBHOOK URL
    # --------------------------------------------------------

    external_url = os.getenv(
        "RENDER_EXTERNAL_URL",
        "",
    ).strip().rstrip("/")

    webhook_url = os.getenv(
        "WEBHOOK_URL",
        "",
    ).strip().rstrip("/")

    if not webhook_url:

        if not external_url:
            raise RuntimeError(
                "RENDER_EXTERNAL_URL не найден"
            )

        webhook_url = (
            f"{external_url}/telegram/webhook/{secret}"
        )

    # --------------------------------------------------------
    # WEBHOOK CONCURRENCY
    # --------------------------------------------------------

    webhook_semaphore = asyncio.Semaphore(
        MAX_WEBHOOK_CONCURRENCY
    )

    # --------------------------------------------------------
    # HEALTH RATE LIMIT
    # --------------------------------------------------------

    last_health_request = 0.0

    health_lock = asyncio.Lock()

    # --------------------------------------------------------
    # HEALTH
    # --------------------------------------------------------

    async def health(
        request: web.Request,
    ) -> web.Response:

        nonlocal last_health_request

        now = time.monotonic()

        async with health_lock:

            if (
                now - last_health_request
                < HEALTH_MIN_INTERVAL
            ):
                return web.Response(
                    status=429,
                    text="Too Many Requests",
                )

            last_health_request = now

        # Никаких запросов к AI,
        # файлам, очереди или хранилищу.
        return web.json_response(
            {
                "ok": True,
                "service": "aitg",
            }
        )

    # --------------------------------------------------------
    # WEBHOOK
    # --------------------------------------------------------

    async def webhook(
        request: web.Request,
    ) -> web.Response:

        # Проверяем URL secret
        if (
            request.match_info.get("secret")
            != secret
        ):
            return web.Response(
                status=404,
                text="Not found",
            )

        # Проверяем Telegram secret header
        header_secret = request.headers.get(
            "X-Telegram-Bot-Api-Secret-Token",
            "",
        )

        if header_secret != secret:
            return web.Response(
                status=403,
                text="Forbidden",
            )

        # Ограничиваем одновременно обрабатываемые
        # HTTP webhook-запросы.
        async with webhook_semaphore:

            try:

                data = await request.json()

                update = Update.model_validate(
                    data
                )

                await app.dp.feed_update(
                    app.bot,
                    update,
                )

                return web.json_response(
                    {"ok": True}
                )

            except asyncio.CancelledError:

                raise

            except Exception:

                bot_main.log.exception(
                    "Webhook update failed"
                )

                return web.json_response(
                    {"ok": False},
                    status=500,
                )

    # --------------------------------------------------------
    # ROUTES
    # --------------------------------------------------------

    http_app.router.add_get(
        "/",
        health,
    )

    http_app.router.add_get(
        "/health",
        health,
    )

    http_app.router.add_post(
        "/telegram/webhook/{secret}",
        webhook,
    )

    # --------------------------------------------------------
    # TELEGRAM WEBHOOK
    # --------------------------------------------------------

    await app.bot.set_webhook(
        webhook_url,
        secret_token=secret,
        drop_pending_updates=False,
        allowed_updates=(
            app.dp.resolve_used_update_types()
        ),
    )

    bot_main.log.info(
        "Webhook configured: %s",
        webhook_url,
    )

    return http_app


# ============================================================
# MAIN
# ============================================================

async def main_async() -> None:

    app = bot_main.App()

    runner: web.AppRunner | None = None

    try:

        # ----------------------------------------------------
        # START BOT SERVICES
        # ----------------------------------------------------

        await prepare_app(app)

        # ----------------------------------------------------
        # CREATE HTTP APP
        # ----------------------------------------------------

        http_app = await create_http_app(
            app
        )

        # ----------------------------------------------------
        # HTTP RUNNER
        # ----------------------------------------------------

        runner = web.AppRunner(
            http_app,
            access_log=None,
        )

        await runner.setup()

        # ----------------------------------------------------
        # TCP SERVER
        # ----------------------------------------------------

        site = web.TCPSite(
            runner,
            HOST,
            PORT,
            reuse_address=True,
        )

        await site.start()

        bot_main.log.info(
            "HTTP server listening on %s:%s",
            HOST,
            PORT,
        )

        # ----------------------------------------------------
        # KEEP PROCESS ALIVE
        # ----------------------------------------------------

        stop_event = asyncio.Event()

        await stop_event.wait()

    finally:

        if runner is not None:

            try:
                await runner.cleanup()

            except Exception:

                bot_main.log.exception(
                    "HTTP runner cleanup failed"
                )

        await cleanup_app(app)


# ============================================================
# ENTRY POINT
# ============================================================

if __name__ == "__main__":
    asyncio.run(
        main_async()
    )
