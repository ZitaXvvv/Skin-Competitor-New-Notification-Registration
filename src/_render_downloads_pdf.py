"""
_render_downloads_pdf.py
检查今天 Downloads 里的所有产品 PDF，对尚未在 image_map.json 里的产品
渲染首页为 JPEG 缩略图，压缩到最长边 ≤320px 后写入 image_map。

用法: python src/_render_downloads_pdf.py [--dry-run]
"""
import argparse, json, shutil
from datetime import date
from pathlib import Path

import fitz          # PyMuPDF
from PIL import Image

WORKSPACE   = Path(__file__).parent.parent
DOWNLOADS   = Path.home() / "Downloads"
MAP_JSON    = WORKSPACE / "res" / "product_images" / "image_map.json"
OUT_DIR     = WORKSPACE / "res" / "product_images" / "_pdf_snapshot"

# 品牌前缀 → brand_en（按长度降序，优先匹配最长前缀）
BRAND_PREFIX = [
    ("雅诗兰黛",   "ESTEE LAUDER"),
    ("修丽可",     "SKIN CEUTICALS"),
    ("欧诗漫",     "OSM"),
    ("自然堂",     "Chando"),
    ("薇诺娜",     "Winona"),
    ("百雀羚",     "BQL"),
    ("珀莱雅",     "PROYA"),
    ("谷雨",       "GUYU"),
    ("韩束",       "Kans"),
    ("兰蔻",       "Lancome"),
    ("娇韵诗",     "Clains"),
    ("科颜氏",     "Kiehls"),
    ("契尔氏",     "Kiehls"),
    ("欧莱雅",     "LOREAL"),
]

MAX_DIM_PX = 320   # 产品卡片缩略图最大边长
JPEG_Q     = 82


def brand_from_name(name: str):
    for prefix, brand_en in BRAND_PREFIX:
        if name.startswith(prefix):
            return brand_en, name  # keep full name including brand prefix
    return None, name


def has_image(brand_en, name, img_map):
    brand_imgs = img_map.get(brand_en, {})
    if not brand_imgs:
        return False
    if name in brand_imgs:
        return True
    # substring match (same as dashboard.py find_prod_img logic)
    for key in brand_imgs:
        clean_key = key
        for pfx, _ in BRAND_PREFIX:
            if key.startswith(pfx):
                clean_key = key[len(pfx):]
                break
        clean_name = name
        for pfx, _ in BRAND_PREFIX:
            if name.startswith(pfx):
                clean_name = name[len(pfx):]
                break
        if len(clean_key) >= 4 and (clean_key in clean_name or clean_name in clean_key):
            return True
    return False


def render_and_compress(pdf_path: Path, out_jpg: Path) -> bool:
    try:
        doc = fitz.open(str(pdf_path))
        if doc.page_count == 0:
            return False
        page = doc[0]
        rect = page.rect
        page_max = max(rect.width, rect.height, 1)
        scale = MAX_DIM_PX / page_max
        scale = max(min(scale, 3.0), 0.05)
        pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale))
        # save to temp png, then re-compress as JPEG
        tmp_png = out_jpg.with_suffix('.tmp.png')
        pix.save(str(tmp_png))
        img = Image.open(tmp_png).convert('RGB')
        # ensure max dim
        w, h = img.size
        if max(w, h) > MAX_DIM_PX:
            ratio = MAX_DIM_PX / max(w, h)
            img = img.resize((int(w*ratio), int(h*ratio)), Image.LANCZOS)
        img.save(str(out_jpg), 'JPEG', quality=JPEG_Q, optimize=True)
        tmp_png.unlink(missing_ok=True)
        return True
    except Exception as e:
        print(f"  ✗ 渲染失败 {pdf_path.name}: {e}")
        return False


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--dry-run', action='store_true', help='只检查，不渲染')
    args = ap.parse_args()

    # 读 image_map
    img_map: dict = json.loads(MAP_JSON.read_text(encoding='utf-8')) if MAP_JSON.exists() else {}

    # 扫描 Downloads 里今天的 PDF（也包括之前几天的，文件名=产品名）
    today = date.today()
    pdfs = sorted(DOWNLOADS.glob("*.pdf"))
    print(f"Downloads 共 {len(pdfs)} 个 PDF")

    has_img, need_render, unknown_brand = [], [], []

    for pdf in pdfs:
        stem = pdf.stem  # 文件名去掉 .pdf
        brand_en, prod_name = brand_from_name(stem)
        if brand_en is None:
            unknown_brand.append(stem)
            continue
        if has_image(brand_en, prod_name, img_map):
            has_img.append((brand_en, prod_name))
        else:
            need_render.append((brand_en, prod_name, pdf))

    print(f"\n✅ 已有图片: {len(has_img)}")
    print(f"⚠️  需要渲染: {len(need_render)}")
    if unknown_brand:
        print(f"❓ 未识别品牌 ({len(unknown_brand)}): {unknown_brand}")

    if args.dry_run:
        print("\n[dry-run] 需要渲染的产品:")
        for brand_en, prod_name, _ in need_render:
            print(f"  [{brand_en}] {prod_name}")
        return

    ok, fail = 0, 0
    for brand_en, prod_name, pdf_path in need_render:
        brand_dir = OUT_DIR / brand_en
        brand_dir.mkdir(parents=True, exist_ok=True)
        # clean filename
        safe = prod_name.replace('/', '_').replace('\\', '_')[:80]
        out_jpg = brand_dir / f"{safe}.jpg"
        print(f"  渲染 [{brand_en}] {prod_name} ...", end=' ')
        if render_and_compress(pdf_path, out_jpg):
            img_map.setdefault(brand_en, {})[prod_name] = str(out_jpg).replace('\\', '/')
            print("✓")
            ok += 1
        else:
            fail += 1

    if ok:
        MAP_JSON.write_text(json.dumps(img_map, ensure_ascii=False, indent=2), encoding='utf-8')
        print(f"\n已写入 image_map.json，新增 {ok} 张图片")

    print(f"\n汇总: 已有={len(has_img)}, 新渲染={ok}, 失败={fail}, 未识别品牌={len(unknown_brand)}")

    # 最终覆盖情况
    print("\n--- 各品牌覆盖 ---")
    by_brand: dict = {}
    for brand_en, prod_name in has_img:
        by_brand.setdefault(brand_en, {'ok': 0, 'miss': 0})['ok'] += 1
    for brand_en, prod_name, _ in need_render:
        by_brand.setdefault(brand_en, {'ok': 0, 'miss': 0})
        if ok:  # if we just rendered
            by_brand[brand_en]['ok'] += 1
        else:
            by_brand[brand_en]['miss'] += 1
    for b, v in sorted(by_brand.items()):
        print(f"  {b:18s} ✅{v['ok']}  ❌{v['miss']}")


if __name__ == '__main__':
    main()
