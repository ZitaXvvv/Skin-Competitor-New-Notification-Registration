"""
管理员账号验证（供 dashboard.py 的管理模式登录使用）。

账号数据存在 src/admins.json（不进 git），每个管理员独立用户名+密码，
密码用 PBKDF2-HMAC-SHA256 加盐哈希存储，不存明文。

用 manage_admins.py 命令行工具增删管理员账号：
    python src/manage_admins.py add <username>
    python src/manage_admins.py remove <username>
    python src/manage_admins.py list
"""
import hashlib
import json
import os
from pathlib import Path

ADMINS_FILE = Path(__file__).parent / "admins.json"
_PBKDF2_ITERS = 200_000


def _hash(password: str, salt: bytes) -> str:
    return hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, _PBKDF2_ITERS).hex()


def make_password_record(password: str) -> dict:
    salt = os.urandom(16)
    return {"salt": salt.hex(), "hash": _hash(password, salt)}


def load_admins() -> dict:
    if not ADMINS_FILE.exists():
        return {}
    try:
        return json.loads(ADMINS_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_admins(admins: dict):
    ADMINS_FILE.write_text(json.dumps(admins, ensure_ascii=False, indent=2), encoding="utf-8")


def verify_login(username: str, password: str) -> bool:
    admins = load_admins()
    rec = admins.get(username)
    if not rec:
        return False
    salt = bytes.fromhex(rec["salt"])
    return _hash(password, salt) == rec["hash"]


def list_admin_usernames() -> list[str]:
    return sorted(load_admins().keys())
