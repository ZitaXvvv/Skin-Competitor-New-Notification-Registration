"""
_parse_word_tables.py
解析 Word 文档的表格结构，把"China Name"列和"Regular Size Pic"列的图片顺序对应
"""
import json, re, zipfile, io, sys
import xml.etree.ElementTree as ET
from pathlib import Path
from PIL import Image as PILImage

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

DOCX = r"C:\Users\xie.x.3\AppData\Local\Temp\CI_Menu_Mar26.docx"
WORKSPACE = Path(__file__).parent.parent
OUT_DIR   = WORKSPACE / "res" / "product_images" / "_word"
MAP_JSON  = WORKSPACE / "res" / "product_images" / "image_map.json"

OUT_DIR.mkdir(parents=True, exist_ok=True)

NS_W  = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_A  = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
R_EMBED = f"{{{NS_R}}}embed"

BRAND_MAP = {
    "珀莱雅": "PROYA",   "谷雨": "GUYU",     "欧诗漫": "OSM",
    "兰蔻":   "Lancome", "欧莱雅": "LOREAL", "雅诗兰黛": "ESTEE LAUDER",
    "修丽可": "SKIN CEUTICALS", "百雀羚": "BQL",
    "韩束":   "Kans",    "自然堂": "Chando", "薇诺娜": "Winona",
    "妮维雅": "Nivea",   "资生堂": "Shiseido","科颜氏": "Kiehls",
}

def detect_brand(text: str) -> str:
    for zh, en in BRAND_MAP.items():
        if zh in text:
            return en
    return ""

def safe_name(s: str) -> str:
    return re.sub(r'[\\/:*?"<>|]', "_", s.strip())

def cell_text(cell_el) -> str:
    texts = [t.text or "" for t in cell_el.findall(f".//{{{NS_W}}}t")]
    return "".join(texts).strip()

def cell_rids(cell_el) -> list[str]:
    blips = cell_el.findall(f".//{{{NS_A}}}blip")
    return [b.get(R_EMBED) for b in blips if b.get(R_EMBED)]

def img_to_png(zf, media_path: str, out_path: Path) -> bool:
    if out_path.exists():
        return True
    if media_path not in zf.namelist():
        return False
    try:
        raw = zf.read(media_path)
        tmp = Path(r"C:\Users\xie.x.3\AppData\Local\Temp\_tmp_tbl.bin")
        tmp.write_bytes(raw)
        img = PILImage.open(str(tmp))
        img.save(str(out_path))
        return True
    except Exception as e:
        sys.stdout.write(f"  转换失败 {media_path}: {e}\n")
        return False


NAME_LABELS  = ("产品中文名", "中文名", "产品名称", "ChineseName", "ChinaName")
IMG_LABELS_PRIMARY  = ("正装图片", "RegularSizePic")
IMG_LABELS_FALLBACK = ("小样照片", "系列图", "MiniPic")
BAD_NAME_RE = re.compile(r"^(Size|是否有样品|N/?A|\d+m?l.*)$", re.IGNORECASE)


def norm_label(s: str) -> str:
    """去除所有空白字符，便于'正装 图片'与'正装图片'这类变体统一匹配"""
    return re.sub(r"\s+", "", s)


def main():
    with zipfile.ZipFile(DOCX) as zf:
        doc_xml  = zf.read("word/document.xml").decode("utf-8")
        rels_xml = zf.read("word/_rels/document.xml.rels").decode("utf-8")

        rels_root = ET.fromstring(rels_xml)
        rid_to_media: dict[str, str] = {}
        for rel in rels_root:
            tgt = rel.get("Target", "")
            if "media" in tgt:
                rid_to_media[rel.get("Id")] = "word/" + tgt.replace("../", "")

        root   = ET.fromstring(doc_xml)
        tables = root.findall(f".//{{{NS_W}}}tbl")
        sys.stdout.write(f"Total tables in document: {len(tables)}\n")

        brand_img_map: dict[str, dict[str, str]] = {}
        matched_total = 0
        skipped = 0

        BRAND_RE = re.compile("|".join(BRAND_MAP.keys()))

        for tbl_idx, tbl in enumerate(tables):
            rows = tbl.findall(f".//{{{NS_W}}}tr")
            if len(rows) < 2:
                continue

            # 每行第一列是"行标签"（产品中文名/正装图片/小样照片/Size等），
            # 之后每一列(c1..cN)对应【一个产品】，这是列式(每列一个SKU)布局。
            tbl_data = []  # list of rows, each row = list of (text, rids)
            for row in rows:
                cells = row.findall(f"./{{{NS_W}}}tc")
                tbl_data.append([(cell_text(c), cell_rids(c)) for c in cells])

            # 按行标签(c0 文本，去空格后)建索引
            row_by_label: dict[str, list] = {}
            for row in tbl_data:
                if not row:
                    continue
                label = norm_label(row[0][0].strip())
                if label:
                    row_by_label.setdefault(label, row)

            name_row = None
            for lbl in NAME_LABELS:
                if lbl in row_by_label:
                    name_row = row_by_label[lbl]
                    break
            if name_row is None:
                skipped += 1
                continue

            img_row_primary  = None
            for lbl in IMG_LABELS_PRIMARY:
                if lbl in row_by_label:
                    img_row_primary = row_by_label[lbl]
                    break
            img_row_fallback = None
            for lbl in IMG_LABELS_FALLBACK:
                if lbl in row_by_label:
                    img_row_fallback = row_by_label[lbl]
                    break

            if img_row_primary is None and img_row_fallback is None:
                skipped += 1
                continue

            full_text = " ".join(t for row in tbl_data for ct, _ in row for t in [ct])
            table_brand = detect_brand(full_text)

            ncols = len(name_row)
            table_matched = 0
            for ci in range(1, ncols):
                prod_name = name_row[ci][0].strip()
                if not prod_name or len(prod_name) < 4 or BAD_NAME_RE.match(prod_name):
                    continue
                rids = []
                if img_row_primary and ci < len(img_row_primary):
                    rids = img_row_primary[ci][1]
                if not rids and img_row_fallback and ci < len(img_row_fallback):
                    rids = img_row_fallback[ci][1]
                if not rids:
                    continue
                brand_en = detect_brand(prod_name) or table_brand
                if not brand_en:
                    continue
                rid = rids[0]
                media_path = rid_to_media.get(rid, "")
                if not media_path:
                    continue
                out_path = OUT_DIR / f"{safe_name(prod_name)}.png"
                if img_to_png(zf, media_path, out_path):
                    brand_img_map.setdefault(brand_en, {})
                    brand_img_map[brand_en].setdefault(prod_name, str(out_path).replace("\\", "/"))
                    matched_total += 1
                    table_matched += 1
            if table_matched:
                sys.stdout.write(f"Table {tbl_idx}: 配对 {table_matched} 个产品\n")

        sys.stdout.write(f"\n表格解析完成，跳过 {skipped} 个非产品表，共配对 {matched_total} 个产品图片\n")

    # 合并到 image_map.json
    existing: dict = {}
    if MAP_JSON.exists():
        existing = json.loads(MAP_JSON.read_text(encoding="utf-8"))

    for brand_en, mapping in brand_img_map.items():
        existing.setdefault(brand_en, {})
        for name, path in mapping.items():
            existing[brand_en].setdefault(name, path)

    MAP_JSON.write_text(json.dumps(existing, ensure_ascii=False, indent=2), encoding="utf-8")
    total = sum(len(v) for v in existing.values())
    sys.stdout.write(f"✅ image_map.json 已更新，合计 {total} 条记录\n")
    sys.stdout.flush()


if __name__ == "__main__":
    main()
