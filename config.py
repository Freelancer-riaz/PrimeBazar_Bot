"""
config.py — Prime Bazar Bot Configuration
==========================================
VPS-তে deploy করার জন্য নিচের values গুলো সরাসরি এখানে লিখুন।
Replit / Railway-তে environment variable থাকলে সেটা auto-ব্যবহার হবে।

👇 VPS-এ আপলোড করার আগে এই ফাইলে আপনার real values বসান:
"""
import os

# ── Telegram Bot Token ────────────────────────────────────────────────────────
# BotFather থেকে পাওয়া token: https://t.me/BotFather
BOT_TOKEN = os.environ.get("BOT_TOKEN", "YOUR_BOT_TOKEN_HERE")

# ── Admin Telegram User ID ────────────────────────────────────────────────────
# @userinfobot দিয়ে আপনার Telegram ID বের করুন
ADMIN_ID = int(os.environ.get("ADMIN_ID", "7522357347"))

# ── Pyrogram Userbot (ঐচ্ছিক — supplier bot bridge এর জন্য) ─────────────────
# https://my.telegram.org থেকে API_ID ও API_HASH নিন
USER_API_ID   = os.environ.get("USER_API_ID",         "YOUR_API_ID_HERE")
USER_API_HASH = os.environ.get("USER_API_HASH",        "YOUR_API_HASH_HERE")
# Pyrogram session string (StringSession) — userbot login এর জন্য
USER_SESSION_STRING = os.environ.get("USER_SESSION_STRING", "YOUR_SESSION_STRING_HERE")

# ── MongoDB Atlas URI ─────────────────────────────────────────────────────────
# https://cloud.mongodb.com → Connect → Python driver URI
# উদাহরণ: mongodb+srv://user:pass@cluster.mongodb.net/prime_bazar
MONGODB_URI = (
    os.environ.get("MONGODB_URI")
    or os.environ.get("MONGO_URI")
    or "YOUR_MONGODB_URI_HERE"
)

# ── Flask Port (VPS-এ সাধারণত দরকার নেই, default 3000) ──────────────────────
PORT = int(os.environ.get("PORT", 3000))

# ── App Domain (optional — webhook/OTP feature এর জন্য) ──────────────────────
# VPS-এ যদি domain না থাকে তাহলে খালি রাখুন ""
APP_DOMAIN = (
    os.environ.get("APP_DOMAIN")
    or (os.environ.get("REPLIT_DOMAINS", "") or os.environ.get("REPLIT_DEV_DOMAIN", "")).split(",")[0]
    or ""
).strip()
