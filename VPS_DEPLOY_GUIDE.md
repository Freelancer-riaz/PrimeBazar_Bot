# Prime Bazar Bot — VPS Deployment Guide
## (Telegram VPS Bot / যেকোনো VPS সার্ভারে চালানোর নিয়ম)

---

## 📁 আপলোড করার ফাইলগুলো

VPS-এ নিচের ফাইলগুলো আপলোড করতে হবে:

| ফাইল | বাধ্যতামূলক? | বিবরণ |
|------|-------------|-------|
| `main.py` | ✅ হ্যাঁ | মূল বট কোড |
| `mongo_db.py` | ✅ হ্যাঁ | MongoDB helper |
| `config.py` | ✅ হ্যাঁ | **আপনার secrets এখানে** |
| `requirements.txt` | ✅ হ্যাঁ | Python packages |
| `settings.json` | ✅ হ্যাঁ | বট settings |
| `texts.json` | ✅ হ্যাঁ | বট texts |
| `market_data.json` | ✅ হ্যাঁ | পণ্যের তথ্য |
| `coupons.json` | ✅ হ্যাঁ | কুপন তথ্য |

> **Telegram VPS Bot** সাধারণত একটি `.py` ফাইল + `requirements.txt` + data files নেয়।  
> যদি শুধু একটি ফাইল নেয়, নিচের "Single File" অংশ দেখুন।

---

## ⚙️ Step 1 — config.py তে আপনার Values বসান

`config.py` ফাইলটি খুলুন এবং placeholder গুলো বদলান:

```python
BOT_TOKEN = "1234567890:AAxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
ADMIN_ID  = 7522357347          # আপনার Telegram numeric ID

USER_API_ID   = "12345678"      # my.telegram.org থেকে
USER_API_HASH = "abcdef1234..." # my.telegram.org থেকে
USER_SESSION_STRING = "BQG..."  # নিচে দেখুন কিভাবে পাবেন

MONGODB_URI = "mongodb+srv://user:pass@cluster.mongodb.net/prime_bazar"
```

### 🔑 USER_SESSION_STRING কিভাবে পাবেন?

আপনার PC/Replit-এ একবার এই code চালান:

```python
from pyrogram import Client

app = Client(
    "my_session",
    api_id=YOUR_API_ID,
    api_hash="YOUR_API_HASH"
)

with app:
    print(app.export_session_string())
```

যে long string print হবে সেটাই `USER_SESSION_STRING`।

---

## 📦 Step 2 — requirements.txt

```
pyTelegramBotAPI==4.34.0
pyrogram==2.0.106
requests==2.34.2
pandas==3.0.3
Flask==3.1.3
Pillow==12.3.0
openpyxl==3.1.5
pymongo==4.10.1
TgCrypto
```

> **নোট:** `TgCrypto` যোগ করলে Pyrogram দ্রুত চলে।

---

## 🚀 Step 3 — VPS-এ চালানো

### Telegram VPS Bot (যেমন All VPS BOT) এর ক্ষেত্রে:
1. "Upload Python File" → `main.py` আপলোড করুন
2. "Upload requirements.txt" → `requirements.txt` আপলোড করুন
3. "Upload Data File" → `config.py` আপলোড করুন (secrets সহ!)
4. আবার "Upload Data File" → `mongo_db.py` আপলোড করুন
5. আবার "Upload Data File" → `settings.json`, `texts.json`, `market_data.json`, `coupons.json` একে একে আপলোড করুন
6. "Deploy Bot" চাপুন

### সাধারণ Linux VPS (SSH access আছে) এর ক্ষেত্রে:
```bash
# সব ফাইল আপলোড করুন (SCP/SFTP দিয়ে)
scp main.py mongo_db.py config.py requirements.txt *.json user@your-vps-ip:/home/bot/

# VPS-এ SSH করুন
ssh user@your-vps-ip
cd /home/bot

# Dependencies install করুন
pip install -r requirements.txt

# বট চালু করুন
python main.py

# Background-এ চালাতে চাইলে:
nohup python main.py > bot.log 2>&1 &
# অথবা screen ব্যবহার করুন:
screen -S primebazar
python main.py
# Ctrl+A, D দিয়ে detach
```

---

## ⚠️ গুরুত্বপূর্ণ সতর্কতা

- **একসাথে দুইটি instance চালাবেন না** — Railway বন্ধ করার পর VPS চালু করুন।
- `config.py`-তে real secrets আছে — এটি কাউকে share করবেন না, GitHub-এ push করবেন না।
- MongoDB Atlas-এ VPS-এর IP whitelist করুন (Network Access → Add IP)।

---

## 🛠️ সমস্যা হলে

| সমস্যা | সমাধান |
|--------|--------|
| `BOT_TOKEN সেট করা নেই` | config.py-তে BOT_TOKEN ঠিকমতো বসানো হয়েছে কি দেখুন |
| `MONGODB_URI not set` | config.py-তে MONGODB_URI চেক করুন |
| `409 Conflict` | Railway বা অন্য instance বন্ধ করুন |
| `AUTH_KEY_DUPLICATED` | Pyrogram session একটি জায়গা থেকেই ব্যবহার করুন |
| MongoDB connection error | Atlas-এ VPS IP whitelist করুন |
