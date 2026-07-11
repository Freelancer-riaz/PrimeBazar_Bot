---
name: Pyrogram utf-16-le surrogate slicing bug
description: Why Pyrogram Message.text/.caption slicing can crash with UnicodeDecodeError on emoji-heavy text, and how to guard against it.
---

Pyrogram's `Message.text`/`.caption` are a custom `Str` subclass whose `__getitem__`
re-encodes the whole string to UTF-16 surrogate pairs and slices/decodes it on every
subscript or slice operation (this matches Telegram's UTF-16-based entity offsets).
If a slice boundary lands in the middle of a surrogate pair (very likely with
emoji-heavy supplier/bot messages), it raises:
`UnicodeDecodeError: 'utf-16-le' codec can't decode bytes...`

This is not a bug in project code — it reproduces with plain `some_message.text[:200]`
whenever the text contains supplementary-plane emoji near the slice boundary.

**Why:** Debugged a live case (Prime Bazar Bot's Pyrogram userbot) where balance-check
and VPN-credential extraction crashed intermittently with this exact error, traced to
`raw[:200]` / `response[:200]` debug-log slicing on Pyrogram message text.

**How to apply:** Immediately convert any Pyrogram `.text`/`.caption` to a plain built-in
`str` (`str(message.text)`) before slicing, regexing, or logging it — this drops the
custom `__getitem__` and makes all further string ops safe. Also wrap the initial
attribute access in try/except `UnicodeDecodeError` since Pyrogram can fail to build
`.text` at all for certain malformed entity data.
