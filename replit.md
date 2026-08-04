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
**বট এখন যেকোনো VPS সার্ভারে চলে (Railway বন্ধ করা হয়েছে)।**
Replit শুধু code editing এর জন্য ব্যবহার হয়।

### VPS Deploy করার ফাইলগুলো:
- `main.py` — মূল বট কোড
- `mongo_db.py` — MongoDB helper
- `config.py` — **সব secrets এখানে** (VPS-এ আপলোডের আগে values বসান)
- `requirements.txt` — Python packages
- `settings.json`, `texts.json`, `market_data.json`, `coupons.json` — data files

বিস্তারিত: **`VPS_DEPLOY_GUIDE.md`** দেখুন।

## User preferences
- Do not start/run the bot workflow on Replit by default — code edits only.
  Running locally risks 409 conflict with the live VPS instance (same BOT_TOKEN
  / USER_SESSION_STRING).
- User communicates in Bengali (বাংলা).
- Secrets are managed via `config.py` for VPS deployment (no env-var system on
  Telegram VPS bots). Replit/Railway fallback still works via os.environ.
