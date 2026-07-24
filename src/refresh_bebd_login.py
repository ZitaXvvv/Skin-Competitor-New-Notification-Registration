"""
refresh_bebd_login.py
交互式 BEBD Cookie 刷新工具。
打开有头浏览器让管理员手动登录，登录后按 Enter 保存 Cookie。
由 scripts/refresh_bebd_login.ps1 调用，不建议直接运行。
"""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from config import BEBD_URL, COOKIES_FILE

def main():
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("错误：未安装 playwright，请先运行 pip install playwright && playwright install chromium")
        sys.exit(1)

    print(f"正在打开浏览器，请登录 {BEBD_URL} ...")
    with sync_playwright() as pw:
        browser = pw.chromium.launch(
            headless=False,   # 有头浏览器，管理员可以操作
            args=["--start-maximized"],
        )
        ctx  = browser.new_context(no_viewport=True)
        page = ctx.new_page()

        # 尝试加载已有 Cookie
        if COOKIES_FILE.exists():
            try:
                existing = json.loads(COOKIES_FILE.read_text(encoding="utf-8"))
                ctx.add_cookies(existing)
                print("  已加载现有 Cookie，正在验证...")
            except Exception:
                pass

        page.goto(BEBD_URL, timeout=30000)

        # 检查是否已登录
        page.wait_for_timeout(3000)
        content = page.content()
        if "退出" in content or "个人中心" in content or "我的" in content:
            print("  当前 Cookie 仍然有效，已自动登录 ✅")
        else:
            print()
            print("  ⚠️  请在浏览器里完成登录操作（账号密码或扫码）...")

        print()
        input("  登录完成后，在此终端按 Enter 保存 Cookie → ")

        # 保存 Cookie
        cookies = ctx.cookies()
        COOKIES_FILE.parent.mkdir(parents=True, exist_ok=True)
        COOKIES_FILE.write_text(
            json.dumps(cookies, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        print(f"  Cookie 已保存到 {COOKIES_FILE}（共 {len(cookies)} 条）")

        browser.close()

    print()
    print("完成。此 Cookie 将被凌晨的计划任务自动复用。")
    print("建议在月底最后一个工作日的下班前执行一次此操作以确保夜间任务正常运行。")


if __name__ == "__main__":
    main()
