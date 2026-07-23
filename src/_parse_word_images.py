"""_parse_word_images.py
从 Word 文档提取产品图片并建立产品名→PNG 的映射
"""
import json, re, zipfile, sys, io
import xml.etree.ElementTree as ET
from pathlib import Path
from PIL import Image as PILImage

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

DOCX = r"C:\Users\xie.x.3\AppData\Local\Temp\CI_Menu_Mar26.docx"
WORKSPACE = Path(__file__).parent.parent
OUT_DIR   = WORKSPACE / "res" / "product_images" / "_word"
MAP_JSON  = WORKSPACE / "res" / "product_images" / "image_map.json"

NS = {
    "w":  "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "a":  "http://schemas.openxmlformats.org/drawingml/2006/main",
}
R_EMBED = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"

BRAND_KW = re.compile(
    r"珀莱雅|谷雨|欧诗漫|兰蔻|欧莱雅|雅诗兰黛|修丽可|百雀羚|韩束|自然堂|薇诺娜|妮维雅|资生堂|科颜氏"
)
# 品牌名 → brand_en
BRAND_MAP = {
    "珀莱雅": "PROYA",   "谷雨": "GUYU",     "欧诗漫": "OSM",
    "兰蔻":   "Lancome", "欧莱雅": "LOREAL", "雅诗兰黛": "ESTEE LAUDER",
    "修丽可": "SKIN CEUTICALS", "百雀羚": "BQL",
    "韩束":   "Kans",    "自然堂": "Chando", "薇诺娜": "Winona",
    "妮维雅": "Nivea",   "资生堂": "Shiseido","科颜氏": "Kiehls",
}

OUT_DIR.mkdir(parents=True, exist_ok=True)


def img_to_png(zf: zipfile.ZipFile, media_path: str, out_path: Path) -> bool:
    if out_path.exists():
        return True
    if media_path not in zf.namelist():
        return False
    try:
        raw = zf.read(media_path)
        tmp = Path(r"C:\Users\xie.x.3\AppData\Local\Temp\_tmp_word_img.bin")
        tmp.write_bytes(raw)
        img = PILImage.open(str(tmp))
        img.save(str(out_path))
        return True
    except Exception as e:
        print(f"  转换失败 {media_path}: {e}")
        return False


def detect_brand(text: str) -> str:
    for zh, en in BRAND_MAP.items():
        if zh in text:
            return en
    return ""


def safe_name(s: str) -> str:
    return re.sub(r'[\\/:*?"<>|]', "_", s.strip())


def main():
    with zipfile.ZipFile(DOCX) as zf:
        doc_xml  = zf.read("word/document.xml").decode("utf-8")
        rels_xml = zf.read("word/_rels/document.xml.rels").decode("utf-8")

        rels_root = ET.fromstring(rels_xml)
        rid_to_img: dict[str, str] = {}
        for rel in rels_root:
            tgt = rel.get("Target", "")
            if "media" in tgt:
                rid_to_img[rel.get("Id")] = "word/" + tgt.replace("../", "")

        root  = ET.fromstring(doc_xml)
        paras = root.findall(".//w:p", NS)

        entries = []
        for p in paras:
            texts = [t.text or "" for t in p.findall(".//w:t", NS)]
            text  = "".join(texts).strip()
            blips = p.findall(".//a:blip", NS)
            rids  = [b.get(R_EMBED) for b in blips if b.get(R_EMBED)]
            if text or rids:
                entries.append({"text": text, "rids": rids})

        # ── 策略：扫描段落，提取图片-产品名对（含/不含品牌前缀）──
        BRAND_KW   = re.compile("|".join(BRAND_MAP.keys()))
        BRAND_SECT = re.compile(
            r"^(珀莱雅|谷雨|欧诗漫|兰蔻|欧莱雅|雅诗兰黛|修丽可|百雀羚|韩束|自然堂|薇诺娜"
            r"|Proya|Guyu|OSM|L.Oreal|Lancome|Estee|Kiehl|Skinceuticals|BQL|Kans|Chando)"
        )
        HAS_ZH = re.compile(r"[\u4e00-\u9fff]{4,}")
        NOT_HEADER = re.compile(
            r"(系列|formula|P6M|share|->|--|NI|Reg|Size|Sample|Have|Name"
            r"|Toner|Emulsion|Cream|Essence|Mask|Cleanser|Eye|Spray|Others"
            r"|Regular|Mini|Pic|UV|\bAA\b|Tone|Hydra)"
        )

        name_to_img: dict[str, list[str]] = {}
        current_brand_en = ""

        i = 0
        while i < len(entries):
            e = entries[i]

            # 更新当前品牌章节
            if e["text"] and not e["rids"] and BRAND_SECT.match(e["text"]):
                det = detect_brand(e["text"])
                if det:
                    current_brand_en = det
                i += 1
                continue

            # 情况1: 图片段本身含带品牌关键词的产品名
            if e["rids"] and BRAND_KW.search(e["text"]):
                prod_name = e["text"].strip()
                half = len(prod_name) // 2
                if half >= 4 and prod_name[:half] == prod_name[half:]:
                    prod_name = prod_name[:half]
                if prod_name and prod_name not in name_to_img:
                    name_to_img[prod_name] = [rid_to_img[r] for r in e["rids"] if r in rid_to_img]
                i += 1
                continue

            # 情况2: 图片段后1~3段有产品名文本
            if e["rids"] and not e["text"]:
                for j in range(i + 1, min(i + 4, len(entries))):
                    nxt = entries[j]
                    if nxt["rids"]:
                        break
                    txt = nxt["text"].strip()
                    if not txt or NOT_HEADER.search(txt):
                        continue
                    half = len(txt) // 2
                    if half >= 4 and txt[:half] == txt[half:]:
                        txt = txt[:half]
                    if len(txt) < 4:
                        continue

                    if BRAND_KW.search(txt):
                        # 带品牌名
                        if txt not in name_to_img:
                            name_to_img[txt] = [rid_to_img[r] for r in e["rids"] if r in rid_to_img]
                        break
                    elif HAS_ZH.search(txt) and current_brand_en:
                        # 不带品牌名：存带品牌前缀版 + 无前缀版
                        zh_brand = next((zh for zh, en in BRAND_MAP.items() if en == current_brand_en), "")
                        full_name = zh_brand + txt if zh_brand else txt
                        for nm in ([full_name, txt] if full_name != txt else [txt]):
                            if nm and nm not in name_to_img:
                                name_to_img[nm] = [rid_to_img[r] for r in e["rids"] if r in rid_to_img]
                        break
            i += 1

        sys.stdout.write(f"找到 {len(name_to_img)} 个产品名-图片对\n")

        # 提取图片，每个产品取第一张
        brand_img_map: dict[str, dict[str, str]] = {}

        for prod_name, img_paths in name_to_img.items():
            if not img_paths:
                continue
            brand_en = detect_brand(prod_name)
            # 无法从名字识别品牌时，尝试用当前章节品牌（不适用，因 detect_brand 已基于名字）
            if not brand_en:
                continue

            # 取第一张图片
            media_path = img_paths[0]
            ext = media_path.rsplit(".", 1)[-1].lower()
            out_path = OUT_DIR / f"{safe_name(prod_name)}.png"

            if img_to_png(zf, media_path, out_path):
                brand_img_map.setdefault(brand_en, {})[prod_name] = str(out_path).replace("\\", "/")
                print(f"  [{brand_en}] {prod_name[:30]} → {out_path.name}")

    # ── 合并到已有的 image_map.json ──
    existing: dict = {}
    if MAP_JSON.exists():
        existing = json.loads(MAP_JSON.read_text(encoding="utf-8"))

    for brand_en, mapping in brand_img_map.items():
        if brand_en not in existing:
            existing[brand_en] = {}
        # Word 里的图片补充进去（不覆盖 Excel 来的图片）
        for name, path in mapping.items():
            existing[brand_en].setdefault(name, path)

    MAP_JSON.write_text(json.dumps(existing, ensure_ascii=False, indent=2), encoding="utf-8")
    total = sum(len(v) for v in existing.values())
    sys.stdout.write(f"\n✅ image_map.json 已更新，合计 {total} 条记录\n")
    sys.stdout.flush()


if __name__ == "__main__":
    main()

import json, re, zipfile, sys, io
import xml.etree.ElementTree as ET
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

DOCX = r"C:\Users\xie.x.3\AppData\Local\Temp\CI_Menu_Mar26.docx"
OUT_DIR = Path(__file__).parent.parent / "res" / "product_images" / "_word"

NS = {
    "w":  "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "a":  "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r":  "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
}
R_EMBED = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"

with zipfile.ZipFile(DOCX) as zf:
    doc_xml  = zf.read("word/document.xml").decode("utf-8")
    rels_xml = zf.read("word/_rels/document.xml.rels").decode("utf-8")

    # rId → media path
    rels_root  = ET.fromstring(rels_xml)
    rid_to_img = {}
    for rel in rels_root:
        tgt = rel.get("Target", "")
        if "media" in tgt:
            rid_to_img[rel.get("Id")] = "word/" + tgt.replace("../", "")

    root  = ET.fromstring(doc_xml)
    paras = root.findall(".//w:p", NS)
    print(f"Total paragraphs: {len(paras)}, total images: {len(rid_to_img)}")

    # 构建段落列表：每段 = {"text": ..., "rids": [...]}
    entries = []
    for p in paras:
        texts = [t.text or "" for t in p.findall(".//w:t", NS)]
        text  = "".join(texts).strip()
        blips = p.findall(".//a:blip", NS)
        rids  = [b.get(R_EMBED) for b in blips if b.get(R_EMBED)]
        if text or rids:
            entries.append({"text": text, "rids": rids})

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    out_txt = OUT_DIR.parent / "_word_structure.txt"

    # 找 "By Brand" 章节起始位置
    start_idx = 0
    for i, e in enumerate(entries):
        if "By Brand" in e["text"] and "Line Up" in e["text"]:
            start_idx = i
            break

    print(f"'By Brand By Line Up' section starts at entry index: {start_idx}")

    with open(str(out_txt), "w", encoding="utf-8") as f:
        f.write(f"Total paragraphs: {len(paras)}, images: {len(rid_to_img)}\n")
        f.write(f"'By Brand' section at entry index: {start_idx}\n\n")
        f.write("=== OVERVIEW (first 60 entries) ===\n")
        for e in entries[:60]:
            if e["rids"]:
                imgs = [rid_to_img.get(r, r)[-30:] for r in e["rids"]]
                f.write(f'[IMG] {imgs}  txt="{e["text"][:60]}"\n')
            elif e["text"]:
                f.write(f'[TXT] "{e["text"][:80]}"\n')
        f.write(f"\n=== BY BRAND SECTION (entries {start_idx} to {min(start_idx+200, len(entries))}) ===\n")
        for e in entries[start_idx: start_idx + 200]:
            if e["rids"]:
                imgs = [rid_to_img.get(r, r)[-30:] for r in e["rids"]]
                f.write(f'[IMG] {imgs}  txt="{e["text"][:60]}"\n')
            elif e["text"]:
                f.write(f'[TXT] "{e["text"][:100]}"\n')

    print(f"Written to: {out_txt}")
