#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""Render Web Service entrypoint for the simplified Telegram AI bot."""

from __future__ import annotations

import asyncio
import hashlib
import os

from aiohttp import web
from aiogram.types import Update

import main as bot_main


HOST = "0.0.0.0"
PORT = int(os.getenv("PORT", "10000"))


def parse_admin_ids() -> set[int]:
    """Parse ADMIN_TELEGRAM_IDS from Render environment."""
    raw = os.getenv("ADMIN_TELEGRAM_IDS", "")

    result: set[int] = set()

    for value in raw.replace(";", ",").replace("\n", ",").split(","):
        value = value.strip()

        if value.startswith("+"):
            value = value[1:].strip()

        if value.lstrip("-").isdigit():
            result.add(int(value))

    return result


# Sync admin IDs with Render environment.
bot_main.ADMIN_IDS.clear()
bot_main.ADMIN_IDS.update(parse_admin_ids())


async def prepare_app(app: bot_main.BotApp) -> None:
    """Prepare the simplified application."""

    if not bot_main.BOT_TOKEN:
        raise RuntimeError("BOT_TOKEN не задан")

    if not bot_main.AI_BASE_URL:
        raise RuntimeError("AI_BASE_URL не задан")

    if not bot_main.AI_API_KEY:
        raise RuntimeError("AI_API_KEY не задан")

    if not bot_main.AI_MODEL_ID:
        raise RuntimeError("AI_MODEL_ID не задан")

    bot_main.log.info(
        "AI configuration loaded: model=%s, model_id=%s",
        bot_main.AI_MODEL_NAME,
        bot_main.AI_MODEL_ID,
    )

    bot_main.log.info(
        "Admins configured: %s",
        sorted(bot_main.ADMIN_IDS),
    )


async def cleanup_app(app: bot_main.BotApp) -> None:
    """Cleanly close bot resources."""

    try:
        await app.ai.close()
    finally:
        await app.bot.session.close()


async def create_http_app(app: bot_main.BotApp) -> web.Application:
    """Create Render HTTP application and Telegram webhook."""

    http_app = web.Application(
        client_max_size=50 * 1024 * 1024
    )

    secret = os.getenv("WEBHOOK_SECRET", "").strip()

    if not secret:
        # Stable secret derived from BOT_TOKEN.
        secret = hashlib.sha256(
            bot_main.BOT_TOKEN.encode("utf-8")
        ).hexdigest()[:48]

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
                "RENDER_EXTERNAL_URL не найден. "
                "Render обычно предоставляет его автоматически. "
                "Либо задайте WEBHOOK_URL вручную."
            )

        webhook_url = (
            f"{external_url}/telegram/webhook/{secret}"
        )

    async def health(request: web.Request) -> web.Response:
        return web.json_response(
            {
                "ok": True,
                "service": "aitg",
                "model": bot_main.AI_MODEL_NAME,
                "model_id": bot_main.AI_MODEL_ID,
                "admins_configured": len(bot_main.ADMIN_IDS),
            }
        )

    async def webhook(request: web.Request) -> web.Response:
        # Check URL secret.
        if request.match_info.get("secret") != secret:
            return web.Response(
                status=404,
                text="Not found",
            )

        # Check Telegram secret header.
        header_secret = request.headers.get(
            "X-Telegram-Bot-Api-Secret-Token",
            "",
        )

        if header_secret != secret:
            return web.Response(
                status=403,
                text="Forbidden",
            )

        try:
            data = await request.json()

            update = Update.model_validate(data)

            await app.dp.feed_update(
                app.bot,
                update,
            )

            return web.json_response(
                {"ok": True}
            )

        except Exception:
            bot_main.log.exception(
                "Webhook update failed"
            )

            return web.json_response(
                {"ok": False},
                status=500,
            )

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

    # Configure Telegram webhook.
    await app.bot.set_webhook(
        webhook_url,
        secret_token=secret,
        drop_pending_updates=False,
        allowed_updates=app.dp.resolve_used_update_types(),
    )

    bot_main.log.info(
        "Telegram webhook configured: %s",
        webhook_url,
    )

    bot_main.log.info(
        "Render HTTP server will listen on %s:%s",
        HOST,
        PORT,
    )

    return http_app


async def main_async() -> None:
    app = bot_main.BotApp()

    try:
        await prepare_app(app)

        http_app = await create_http_app(app)

        runner = web.AppRunner(http_app)

        await runner.setup()

        site = web.TCPSite(
            runner,
            HOST,
            PORT,
        )

        await site.start()

        bot_main.log.info(
            "HTTP server listening on %s:%s",
            HOST,
            PORT,
        )

        stop_event = asyncio.Event()

        try:
            await stop_event.wait()

        finally:
            await runner.cleanup()

    finally:
        await cleanup_app(app)


if __name__ == "__main__":
    asyncio.run(main_async())
