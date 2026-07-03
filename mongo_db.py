"""
mongo_db.py — MongoDB backend for Prime Bazar Bot.
All JSON-file data operations replaced with MongoDB equivalents.
In-memory caching is preserved; MongoDB is the persistence layer.
"""
import os
from pymongo import MongoClient, UpdateOne
from pymongo.errors import PyMongoError

MONGODB_URI = os.environ.get("MONGODB_URI", "") or os.environ.get("MONGO_URI", "")

_client = None
_db_instance = None


def _db():
    global _client, _db_instance
    if _db_instance is None:
        if not MONGODB_URI:
            raise RuntimeError("MONGODB_URI environment variable is not set.")
        _client = MongoClient(MONGODB_URI, serverSelectionTimeoutMS=8000)
        _db_instance = _client["prime_bazar"]
    return _db_instance


# ─── Generic single-document helpers ─────────────────────────────────────────

def _get_doc(collection: str, doc_id: str = "main") -> dict:
    try:
        doc = _db()[collection].find_one({"_id": doc_id})
        if doc:
            doc.pop("_id", None)
            return doc
    except PyMongoError:
        pass
    return {}


def _set_doc(collection: str, data: dict, doc_id: str = "main"):
    try:
        _db()[collection].replace_one(
            {"_id": doc_id}, {"_id": doc_id, **data}, upsert=True
        )
    except PyMongoError:
        pass


# ─── Settings ─────────────────────────────────────────────────────────────────

def db_load_settings(defaults: dict) -> dict:
    data = _get_doc("settings")
    if not data:
        return dict(defaults)
    for k, v in defaults.items():
        data.setdefault(k, v)
    return data


def db_save_settings(data: dict):
    _set_doc("settings", data)


# ─── Texts ────────────────────────────────────────────────────────────────────

def db_load_texts(fallback: dict) -> dict:
    data = _get_doc("texts")
    return data if data else fallback


def db_save_texts(data: dict):
    _set_doc("texts", data)


# ─── Market Data ──────────────────────────────────────────────────────────────

def db_load_market(default_vpn_durs: dict) -> dict:
    data = _get_doc("market_data")
    if not data:
        return {"categories": {}, "products": {}, "vpn_durations": dict(default_vpn_durs)}
    data.setdefault("vpn_durations", default_vpn_durs)
    data.setdefault("categories", {})
    data.setdefault("products", {})
    return data


def db_save_market(data: dict):
    _set_doc("market_data", data)


# ─── Coupons ──────────────────────────────────────────────────────────────────

def db_load_coupons() -> dict:
    data = _get_doc("coupons")
    return data if data else {}


def db_save_coupons(data: dict):
    _set_doc("coupons", data)


# ─── Used TRX IDs ─────────────────────────────────────────────────────────────

def db_load_trxids() -> set:
    try:
        doc = _db()["used_trxids"].find_one({"_id": "main"})
        if doc:
            return set(doc.get("ids", []))
    except PyMongoError:
        pass
    return set()


def db_save_trxids(data: set):
    try:
        _db()["used_trxids"].replace_one(
            {"_id": "main"}, {"_id": "main", "ids": list(data)}, upsert=True
        )
    except PyMongoError:
        pass


# ─── Pending Deposits ─────────────────────────────────────────────────────────

def db_load_pending_deps() -> dict:
    try:
        doc = _db()["pending_deposits"].find_one({"_id": "main"})
        return doc.get("data", {}) if doc else {}
    except PyMongoError:
        return {}


def db_save_pending_deps(data: dict):
    try:
        _db()["pending_deposits"].replace_one(
            {"_id": "main"}, {"_id": "main", "data": data}, upsert=True
        )
    except PyMongoError:
        pass


# ─── Rejected Deposits ────────────────────────────────────────────────────────

def db_load_rejected_deps() -> list:
    try:
        doc = _db()["rejected_deposits"].find_one({"_id": "main"})
        return doc.get("data", []) if doc else []
    except PyMongoError:
        return []


def db_save_rejected_deps(data: list):
    try:
        _db()["rejected_deposits"].replace_one(
            {"_id": "main"}, {"_id": "main", "data": data}, upsert=True
        )
    except PyMongoError:
        pass


# ─── Pending Manual Orders ────────────────────────────────────────────────────

def db_load_manual_orders() -> dict:
    try:
        doc = _db()["pending_manual_orders"].find_one({"_id": "main"})
        return doc.get("data", {}) if doc else {}
    except PyMongoError:
        return {}


def db_save_manual_orders(data: dict):
    try:
        _db()["pending_manual_orders"].replace_one(
            {"_id": "main"}, {"_id": "main", "data": data}, upsert=True
        )
    except PyMongoError:
        pass


# ─── Users ────────────────────────────────────────────────────────────────────

def db_load_all_users(defaults: dict) -> dict:
    """Load ALL users from MongoDB into an in-memory dict at startup."""
    try:
        result = {}
        for doc in _db()["users"].find():
            uid = str(doc.pop("_id"))
            for k, v in defaults.items():
                doc.setdefault(k, v)
            if not doc.get("language"):
                doc["language"] = "bn"
            result[uid] = doc
        return result
    except PyMongoError:
        return {}


def db_save_one_user(uid: str, data: dict):
    """Save a single user document (fast — called on every user change)."""
    try:
        _db()["users"].replace_one(
            {"_id": uid}, {"_id": uid, **data}, upsert=True
        )
    except PyMongoError:
        pass


def db_save_all_users(all_users: dict):
    """Bulk-write all users (used for full backup restore / import)."""
    if not all_users:
        return
    try:
        ops = [
            UpdateOne({"_id": uid}, {"$set": {**data}}, upsert=True)
            for uid, data in all_users.items()
        ]
        _db()["users"].bulk_write(ops, ordered=False)
    except PyMongoError:
        pass


# ─── Mail Stock (accounts uploaded via xlsx) ──────────────────────────────────
# Each product's remaining accounts are stored as a list of row-dicts under
# collection "mail_stock", one document per product (_id = product name).
# This replaces local .xlsx files so stock survives Railway restarts/redeploys.

def db_load_stock(p_name: str) -> list:
    """Load all remaining stock rows for a product."""
    try:
        doc = _db()["mail_stock"].find_one({"_id": p_name})
        return doc.get("rows", []) if doc else []
    except PyMongoError:
        return []


def db_save_stock(p_name: str, rows: list):
    """Replace all stock rows for a product (used after a purchase consumes
    some rows, leaving the remainder)."""
    try:
        _db()["mail_stock"].replace_one(
            {"_id": p_name}, {"_id": p_name, "rows": rows}, upsert=True
        )
    except PyMongoError:
        pass


def db_append_stock(p_name: str, rows: list):
    """Append newly uploaded rows to a product's existing stock, keeping any
    unsold accounts from before."""
    if not rows:
        return
    try:
        _db()["mail_stock"].update_one(
            {"_id": p_name},
            {"$push": {"rows": {"$each": rows}}},
            upsert=True,
        )
    except PyMongoError:
        pass


def db_stock_count(p_name: str) -> int:
    try:
        doc = _db()["mail_stock"].find_one({"_id": p_name}, {"rows": 1})
        return len(doc.get("rows", [])) if doc else 0
    except PyMongoError:
        return 0


def db_delete_stock(p_name: str):
    try:
        _db()["mail_stock"].delete_one({"_id": p_name})
    except PyMongoError:
        pass
