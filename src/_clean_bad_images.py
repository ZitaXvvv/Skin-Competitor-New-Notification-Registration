"""
一次性修复脚本：清理 image_map.json 中被误判为产品照片的 OLE 对象图标截图
(固定尺寸 191x129，实际内容是 Excel 里插入的"以图标显示"的 PDF 附件截图，
文字显示的是被嵌入 PDF 的文件名 —— 有时与所在行的产品名完全不同，等于是
"张冠李戴"的错误图片)，以及扫描全部图片文件、剔除已损坏（无法被 PIL 打开）
的文件。清理后从 image_map.json 中删除对应条目，并删除磁盘上的坏文件，
让 _generate_missing_img_from_pdf.py 后续可以用真实 PDF 链接重新生成图片。

用法：python src/_clean_bad_images.py
"""
import json
from pathlib import Path
from PIL import Image

WORKSPACE = Path(__file__).parent.parent
MAP_JSON = WORKSPACE / "res" / "product_images" / "image_map.json"
BAD_ICON_SIZE = (191, 129)


def main():
    m = json.loads(MAP_JSON.read_text(encoding="utf-8"))
    removed = []
    for brand, mp in list(m.items()):
        for name, path in list(mp.items()):
            p = Path(path)
            if not p.exists():
                continue
            bad = False
            reason = ""
            try:
                im = Image.open(p)
                if im.size == BAD_ICON_SIZE:
                    bad = True
                    reason = "OLE图标误判"
            except Exception as e:
                bad = True
                reason = f"文件损坏: {e}"
            if bad:
                removed.append((brand, name, str(p), reason))
                del mp[name]
                try:
                    p.unlink()
                except Exception:
                    pass

    print(f"共清理 {len(removed)} 条：")
    for brand, name, path, reason in removed:
        print(f"  [{brand}] {name} ({reason})")

    MAP_JSON.write_text(json.dumps(m, ensure_ascii=False, indent=2), encoding="utf-8")
    total = sum(len(v) for v in m.values())
    print(f"\n清理后剩余 {total} 条记录，已保存 image_map.json")


if __name__ == "__main__":
    main()
