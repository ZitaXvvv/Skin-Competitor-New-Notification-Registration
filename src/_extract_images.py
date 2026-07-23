"""
_extract_images.py
------------------
从各品牌 Excel 文件中提取嵌入图片，按产品名保存为 PNG，
并生成一个 JSON 映射：{品牌: {产品名: 图片路径}}

输出目录: res/product_images/{brand_en}/
映射文件: res/product_images/image_map.json
"""

import json
import re
import sys
import zipfile
from pathlib import Path

from PIL import Image

# ── 路径 ───────────────────────────────────────────────
WORKSPACE = Path(__file__).parent.parent
TEMP      = Path(r"C:\Users\xie.x.3\AppData\Local\Temp\1")
OUT_DIR   = WORKSPACE / "res" / "product_images"

# 品牌文件名 → (brand_en, sheet_name)
BRAND_FILES = {
    "百雀羚.xlsx":  "BQL",
    "谷雨.xlsx":    "GUYU",
    "韩束.xlsx":    "Kans",
    "兰蔻.xlsx":    "Lancome",
    "欧诗漫.xlsx":  "OSM",
    "修丽可.xlsx":  "SKIN CEUTICALS",
    "自然堂.xlsx":  "Chando",
}


def parse_vml_map(vml_bytes: bytes) -> dict[str, int]:
    """
    解析 VML drawing XML，返回 {rId: 0-based-row} 映射。
    如 VML 不含 x:Row，回退到顺序映射（第N个shape=第N行）。
    """
    vml = vml_bytes.decode("utf-8", errors="replace")
    shapes = re.findall(r"<v:shape\b.*?</v:shape>", vml, re.DOTALL)
    rid_to_row: dict[str, int] = {}
    for idx, shape in enumerate(shapes):
        rid_m = re.search(r'relid=["\x27](rId\d+)["\x27]', shape)
        row_m = re.search(r"<x:Row>(\d+)</x:Row>", shape)
        if rid_m:
            rid = rid_m.group(1)
            row = int(row_m.group(1)) if row_m else idx
            rid_to_row[rid] = row
    return rid_to_row


def parse_rels(rels_bytes: bytes) -> dict[str, str]:
    """解析 .rels 文件，返回 {rId: 图片相对路径}"""
    import xml.etree.ElementTree as ET
    root = ET.fromstring(rels_bytes.decode("utf-8"))
    ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    rid_to_file: dict[str, str] = {}
    for rel in root:
        tag = rel.tag.replace(f"{{{ns}}}", "")
        if tag == "Relationship":
            rid_to_file[rel.get("Id")] = rel.get("Target", "")
    return rid_to_file


def extract_brand_images(xlsx_path: Path, brand_en: str) -> dict[str, str]:
    """
    从单个品牌 Excel 提取图片，转为 PNG，
    返回 {产品名: PNG 绝对路径}
    """
    out_brand = OUT_DIR / brand_en
    out_brand.mkdir(parents=True, exist_ok=True)

    result: dict[str, str] = {}

    with zipfile.ZipFile(xlsx_path) as zf:
        # 1. 读取产品名列表（sheet "护肤"，col A = 商品名称，从第2行开始）
        import openpyxl
        wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
        ws = wb.active  # 护肤 sheet
        product_names: list[str] = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            name = str(row[0] or "").strip()
            if name:
                product_names.append(name)
        wb.close()

        # 2. 找 VML 和 rels 文件
        vml_files = [n for n in zf.namelist() if "vml" in n.lower() and n.endswith(".vml")]
        rels_files = [n for n in zf.namelist() if "drawings" in n.lower() and ".rels" in n]

        if not vml_files or not rels_files:
            print(f"  ⚠  {xlsx_path.name}: 找不到 VML/rels，跳过")
            return result

        vml_data  = zf.read(vml_files[0])
        rels_data = zf.read(rels_files[0])

        rid_to_row  = parse_vml_map(vml_data)    # {rId: 0-based-row}
        rid_to_file = parse_rels(rels_data)       # {rId: ../media/imageN.emf}

        # 3. row → 产品名（row 0 = 标题行，row 1 = 第1个产品）
        row_to_name: dict[int, str] = {}
        for i, name in enumerate(product_names):
            row_to_name[i + 1] = name   # 数据从 row 1 开始（0 = 标题）

        # 如果 VML 解析失败（rid_to_row 为空），按顺序映射
        if not rid_to_row:
            all_rids = sorted(rid_to_file.keys(), key=lambda r: int(r[3:]))
            for seq, rid in enumerate(all_rids):
                row_to_name_seq = product_names[seq] if seq < len(product_names) else None
                if row_to_name_seq:
                    row_to_name[seq + 1] = row_to_name_seq
            rid_to_row = {rid: i + 1 for i, rid in enumerate(all_rids)}

        # 4. 提取并转换图片
        for rid, row in rid_to_row.items():
            prod_name = row_to_name.get(row)
            if not prod_name:
                continue
            rel_path = rid_to_file.get(rid, "")
            # ../media/imageN.emf → xl/media/imageN.emf
            media_path = "xl/media/" + rel_path.replace("../media/", "")
            if media_path not in zf.namelist():
                continue

            # 安全文件名（去掉非法字符）
            safe_name = re.sub(r'[\\/:*?"<>|]', "_", prod_name)
            png_path = out_brand / f"{safe_name}.png"

            if not png_path.exists():
                raw_emf = zf.read(media_path)
                # 写临时 EMF
                tmp = Path(r"C:\Users\xie.x.3\AppData\Local\Temp\_tmp_img.emf")
                tmp.write_bytes(raw_emf)
                try:
                    img = Image.open(str(tmp))
                    img.save(str(png_path))
                except Exception as e:
                    print(f"    转换失败 {prod_name}: {e}")
                    continue

            result[prod_name] = str(png_path)

    print(f"  {xlsx_path.name} → {brand_en}: 提取 {len(result)} 张图")
    return result


def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    image_map: dict[str, dict[str, str]] = {}

    for fname, brand_en in BRAND_FILES.items():
        src = TEMP / fname
        if not src.exists():
            print(f"⚠  {fname} 不存在，跳过")
            continue
        mapping = extract_brand_images(src, brand_en)
        image_map[brand_en] = mapping

    # 保存映射 JSON（路径转为正斜杠以便跨平台）
    map_clean = {
        brand: {name: path.replace("\\", "/") for name, path in m.items()}
        for brand, m in image_map.items()
    }
    out_json = OUT_DIR / "image_map.json"
    out_json.write_text(json.dumps(map_clean, ensure_ascii=False, indent=2), encoding="utf-8")
    total = sum(len(v) for v in image_map.values())
    print(f"\n✅ 共提取 {total} 张图片，映射保存到 {out_json}")


if __name__ == "__main__":
    main()
