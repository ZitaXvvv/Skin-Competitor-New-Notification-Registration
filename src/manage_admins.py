"""
管理员账号管理命令行工具。

用法：
    python src/manage_admins.py add <username>       # 新增/重置管理员密码（交互式输入，不回显）
    python src/manage_admins.py remove <username>     # 删除管理员账号
    python src/manage_admins.py list                  # 列出所有管理员用户名
"""
import getpass
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from auth import load_admins, save_admins, make_password_record, list_admin_usernames


def cmd_add(username: str):
    pw1 = getpass.getpass(f"设置 {username} 的密码: ")
    pw2 = getpass.getpass("再次输入确认: ")
    if pw1 != pw2:
        print("❌ 两次输入不一致，取消")
        return
    if not pw1:
        print("❌ 密码不能为空")
        return
    admins = load_admins()
    admins[username] = make_password_record(pw1)
    save_admins(admins)
    print(f"✅ 已保存管理员账号: {username}")


def cmd_remove(username: str):
    admins = load_admins()
    if username not in admins:
        print(f"⚠️ 账号不存在: {username}")
        return
    del admins[username]
    save_admins(admins)
    print(f"✅ 已删除管理员账号: {username}")


def cmd_list():
    names = list_admin_usernames()
    if not names:
        print("（当前没有任何管理员账号）")
        return
    for n in names:
        print(f"  - {n}")


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return
    action = sys.argv[1]
    if action == "add" and len(sys.argv) >= 3:
        cmd_add(sys.argv[2])
    elif action == "remove" and len(sys.argv) >= 3:
        cmd_remove(sys.argv[2])
    elif action == "list":
        cmd_list()
    else:
        print(__doc__)


if __name__ == "__main__":
    main()
