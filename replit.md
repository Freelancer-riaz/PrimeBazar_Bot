# Prime Bazar Bot

## Overview
A Telegram shop bot ("Prime Bazar Bot") built with:
- `pyTelegramBotAPI` (telebot) — main customer-facing bot (polling)
- Flask — lightweight keep-alive/admin web server
- MongoDB — persistent storage
- Pyrogram (optional userbot) — bridges to a third-party VPN supplier bot to
  automate balance checks, navigation, and credential delivery (see `_userbot`,
  `_pyro_main` in `main.py`)

Settings (menus, texts, prices, toggles) are driven by JSON files
(`settings.json`, `texts.json`, `market_data.json`, `coupons.json`) — most
behavior changes don't require code edits.

## Production deployment
**The bot runs in production on Railway, not on Replit.** This repl is the
development/editing environment only.

## User preferences
- Do not start/run the bot workflow on Replit by default — the user tests
  changes themselves on Railway. Only make code edits here and push to
  GitHub (`origin`); the user pulls/redeploys on Railway and reports results.
  Running the bot here risks colliding with the live Railway instance (same
  `BOT_TOKEN`/`USER_SESSION_STRING` — simultaneous polling causes Telegram
  409 conflicts, and simultaneous Pyrogram userbot login can invalidate the
  session with `AUTH_KEY_DUPLICATED`).
- If a live Replit test run is ever needed (with explicit user consent), stop
  the workflow again afterward instead of leaving it running.
