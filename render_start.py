#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Render Web Service entrypoint for the Telegram bot.

Keeps the existing App/handlers in main.py and replaces long polling with
Telegram webhook delivery so the service can listen on Render's $PORT.
"""

from __future__ import annotations

import asyncio
import hashlib
import os
from pathlib import Path

from aiohttp import web
from aiogram.types import Update

import main as bot_main


HOST = "0.0.0.0"
PORT = int(os.getenv("PORT", "10000"))


def parse_admin_ids() -> set[int]:
    """Parse Render's ADMIN_TELEGRAM_IDS robustly.

    Accepts commas, semicolons and newlines and ignores accidental spaces.
    """
    raw = os.getenv("ADMIN_TELEGRAM_IDS", "")
    result: set[int] = set()
    for value in raw.replace(";", ",").replace("\n", ",").split(","):
        value = value.strip()
        if value.startswith("+"):
            value = value[1:].strip()
        if value.lstrip("-").isdigit():
            result.add(int(value))
    return result


# Override the parsed set from main.py with the same Render environment value,
# but using the more tolerant parser above.
bot_main.ADMIN_IDS.clear()
bot_main.ADMIN_IDS.update(parse_admin_ids())


async def prepare_app(app: bot_main.App) -> None:
    if not bot_main.BOT_TOKEN:
        raise RuntimeError("BOT_TOKEN не задан")

    await app.db.connect()
    await app.repo.seed_defaults()
    await app.repo.ensure_seed_model()
    await app.ai.start()

    # Recover unfinished jobs after a restart.
    await app.db.execute(
        "UPDATE jobs SET state=?, finished_at=? WHERE state=?",
        (bot_main.STATUS_INTERRUPTED, bot_main.iso_now(), bot_main.STATUS_RUNNING),
    )

    persisted_jobs = await app.db.fetchall(
        "SELECT * FROM jobs WHERE state=? ORDER BY priority ASC, created_at ASC",
        (bot_main.STATUS_QUEUED,),
    )
    for job in persisted_jobs:
        payload = bot_main.safe_json_loads(job["payload_json"], {}) or {}
        await app.queue.enqueue(
            job["id"],
            job["job_type"],
            payload,
            int(job["priority"]),
        )

    await app.queue.start()


async def cleanup_app(app: bot_main.App) -> None:
    try:
        await app.queue.stop()
    finally:
        try:
            await app.ai.close()
        finally:
            try:
                await app.db.close()
            finally:
                await app.bot.session.close()


async def create_http_app(app: bot_main.App) -> web.Application:
    http_app = web.Application(client_max_size=50 * 1024 * 1024)

    secret = os.getenv("WEBHOOK_SECRET", "").strip()
    if not secret:
        # Stable secret derived from the bot token; it is not printed.
        secret = hashlib.sha256(bot_main.BOT_TOKEN.encode("utf-8")).hexdigest()[:48]

    external_url = os.getenv("RENDER_EXTERNAL_URL", "").strip().rstrip("/")
    if not external_url:
        raise RuntimeError(
            "RENDER_EXTERNAL_URL не найден. В Render укажите публичный URL сервиса "
            "в переменной RENDER_EXTERNAL_URL или задайте WEBHOOK_URL вручную."
        )

    webhook_url = os.getenv("WEBHOOK_URL", "").strip().rstrip("/")
    if not webhook_url:
        webhook_url = f"{external_url}/telegram/webhook/{secret}"

    async def health(request: web.Request) -> web.Response:
        return web.json_response({
            "ok": True,
            "service": "aitg",
            "admins_configured": len(bot_main.ADMIN_IDS),
        })

    async def webhook(request: web.Request) -> web.Response:
        # Check both the secret path and Telegram's official secret header.
        if request.match_info.get("secret") != secret:
            return web.Response(status=404, text="Not found")

        header_secret = request.headers.get("X-Telegram-Bot-Api-Secret-Token", "")
        if header_secret != secret:
            return web.Response(status=403, text="Forbidden")

        try:
            data = await request.json()
            update = Update.model_validate(data)
            await app.dp.feed_update(app.bot, update)
            return web.json_response({"ok": True})
        except Exception:
            # Telegram only needs a normal HTTP response. Details stay in logs.
            bot_main.log.exception("Webhook update failed")
            return web.json_response({"ok": False}, status=500)

    http_app.router.add_get("/", health)
    http_app.router.add_get("/health", health)
    http_app.router.add_post("/telegram/webhook/{secret}", webhook)

    await app.bot.set_webhook(
        webhook_url,
        secret_token=secret,
        drop_pending_updates=False,
        allowed_updates=app.dp.resolve_used_update_types(),
    )

    bot_main.log.info(
        "Render webhook started; admins=%s; port=%s; webhook configured",
        sorted(bot_main.ADMIN_IDS),
        PORT,
    )

    return http_app


async def main_async() -> None:
    app = bot_main.App()
    try:
        await prepare_app(app)
        http_app = await create_http_app(app)
        runner = web.AppRunner(http_app)
        await runner.setup()
        site = web.TCPSite(runner, HOST, PORT)
        await site.start()

        bot_main.log.info("HTTP server listening on %s:%s", HOST, PORT)

        stop_event = asyncio.Event()
        try:
            await stop_event.wait()
        finally:
            await runner.cleanup()
    finally:
        await cleanup_app(app)


if __name__ == "__main__":
    asyncio.run(main_async())
