"""
用真实产品名称，逐条查询 hzpba.nmpa.gov.cn 补全 730 天历史抓取里
"仅恢复了名称+照片但丢了备案号/日期"的 64 个普通备案产品的备案号和备案日期。

结果增量写入 checkpoint JSON，防止中途失败丢失已查到的数据。

可移植运行说明（例如打包到其他服务器重跑）：
1. 只需要同目录下的 config.py 和 _recover_targets.json 两个文件，不需要拷贝本地 PDF
   （本脚本不再扫描本地 Windows 路径，只从 JSON 里读品牌/产品名）。
2. 依赖：pip install playwright openpyxl （不需要 pywin32/pymupdf，那些是Windows专用/不再用到），
   然后执行 playwright install chromium（Linux 上用 playwright install --with-deps chromium）。
3. 运行 python _recover_hzpba_dates.py，结果增量写入 _recover_checkpoint.json，跑完后把这个文件拷回来即可。
4. 如果中途中断，直接重跑同一命令会自动跳过 checkpoint 里已有结果的条目，继续剩余的。
"""
import json
import re
import sys
from pathlib import Path

from playwright.sync_api import sync_playwright

sys.path.insert(0, str(Path(__file__).parent))
from config import HZPBA_SEARCH_URL, HZPBA_IMPORT_URL

TARGETS_FILE = Path(__file__).parent / "_recover_targets.json"
CKPT = Path(__file__).parent / "_recover_checkpoint.json"

REG_PAT = re.compile(r"(国妆特字|国妆特进字|卫妆特字|妆网备字|国妆备字|国妆备进字)[0-9A-Za-z（）()]*")
DATE_PAT = re.compile(r"(20\d{2})[-/年](\d{1,2})[-/月](\d{1,2})")


def collect_targets():
    return json.loads(TARGETS_FILE.read_text(encoding="utf-8"))


def load_ckpt():
    if CKPT.exists():
        return json.loads(CKPT.read_text(encoding="utf-8"))
    return {}


def save_ckpt(ckpt):
    CKPT.write_text(json.dumps(ckpt, ensure_ascii=False, indent=2), encoding="utf-8")


def parse_date(s: str):
    m = DATE_PAT.search(s)
    if not m:
        return None
    y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
    try:
        return f"{mo:02d}/{d:02d}/{y}"
    except ValueError:
        return None


def search_one(page, name: str, is_imported: bool):
    url = HZPBA_IMPORT_URL if is_imported else HZPBA_SEARCH_URL
    for attempt in range(3):
        try:
            page.goto(url, timeout=30000)
            page.wait_for_timeout(3000)
            if page.locator("body").inner_text().strip():
                break
        except Exception:
            pass
        page.wait_for_timeout(3000)

    box = page.locator(
        "input[placeholder*='产品名称'], input[placeholder*='请输入'], .el-input__inner"
    ).first
    box.wait_for(timeout=20000)
    box.fill(name)
    page.locator(
        "button:has-text('查询'), button:has-text('搜索'), .el-button--primary"
    ).first.click()
    page.wait_for_timeout(2500)

    no_data = page.locator("text=暂无数据, text=未查到, .el-table__empty-block, .no-data").first
    if no_data.count() > 0:
        return []

    rows = page.locator("tr.el-table__row, .el-table__body tr").all()
    out = []
    for row in rows:
        try:
            text = row.inner_text()
        except Exception:
            continue
        if text.strip():
            out.append(text)
    return out


def main():
    targets = collect_targets()
    ckpt = load_ckpt()
    print(f"待查询: {len(targets)} 条，已有进度: {len(ckpt)} 条")

    with sync_playwright() as pw:
        browser = pw.chromium.launch(
            headless=True,
            args=[
                "--lang=zh-CN",
                "--disable-blink-features=AutomationControlled",
                "--no-sandbox",
            ],
        )
        ctx = browser.new_context(
            locale="zh-CN",
            ignore_https_errors=True,
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
        )
        page = ctx.new_page()

        for i, t in enumerate(targets):
            key = f"{t['brand']}::{t['name']}"
            if key in ckpt:
                continue
            print(f"[{i+1}/{len(targets)}] {t['brand']} / {t['name']}")
            result = {"brand": t["brand"], "name": t["name"], "reg_num": None,
                      "date": None, "status": "not_found", "raw": None}
            try:
                rows = search_one(page, t["name"], is_imported=False)
                if not rows:
                    rows = search_one(page, t["name"], is_imported=True)
                    if rows:
                        result["imported"] = True
                if rows:
                    best = rows[0]
                    m_reg = REG_PAT.search(best)
                    d = parse_date(best)
                    result["reg_num"] = m_reg.group(0) if m_reg else None
                    result["date"] = d
                    result["raw"] = best[:300]
                    result["status"] = "ok" if (m_reg and d) else "partial"
                    result["n_rows"] = len(rows)
                    print(f"    -> {result['status']}: reg={result['reg_num']} date={result['date']} (共{len(rows)}行匹配)")
                else:
                    print("    -> 无结果")
            except Exception as exc:
                result["status"] = "error"
                result["error"] = str(exc)[:200]
                print(f"    -> 出错: {exc}")

            ckpt[key] = result
            save_ckpt(ckpt)

        browser.close()

    print("完成，结果已保存到", CKPT)


if __name__ == "__main__":
    main()
