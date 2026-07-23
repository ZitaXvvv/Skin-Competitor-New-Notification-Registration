"""
一次性修复脚本：把 image_map.json 引用的全部产品图片统一压缩为小尺寸JPEG缩略图。

背景：dashboard.py 把 image_map.json 里所有图片一次性 base64 编码后内嵌进
一个巨大的 HTML 字符串，通过 st.iframe() 整体发给浏览器。之前只对
_pdf_snapshot 目录里体积异常大(>100MB)的文件做了1000px上限的缩放，但
615张图片累计仍有 ~110MB 原始体积（base64后~150MB），作为单次 iframe
payload 发送时会导致浏览器端渲染/测量高度失败或 websocket 发送失败，
表现为"页面头部/筛选栏正常，日历卡片内容整体空白/加载不出来"。

而实际展示时每张产品图在卡片里只是一个很小的正方形缩略图（CSS
aspect-ratio:1 的卡片格子），完全不需要现在动辄 700~1000px 的分辨率。

本脚本把每张图片统一缩放到最长边 <= MAX_DIM 像素，转成 RGB 后用 JPEG
(quality=82) 重新保存（PNG对照片类内容压缩率远不如JPEG），并更新
image_map.json 里的路径。

用法：python src/_shrink_all_product_images.py
"""
import json
from pathlib import Path
from PIL import Image

WORKSPACE = Path(__file__).parent.parent
MAP_JSON = WORKSPACE / "res" / "product_images" / "image_map.json"
MAX_DIM = 320
JPEG_QUALITY = 82


def main():
    m = json.loads(MAP_JSON.read_text(encoding="utf-8"))
    total_before = 0
    total_after = 0
    converted = 0
    for brand, mapping in m.items():
        for name, path in list(mapping.items()):
            p = Path(path)
            if not p.exists():
                continue
            size_before = p.stat().st_size
            total_before += size_before
            try:
                img = Image.open(p)
                img = img.convert("RGB")
                w, h = img.size
                scale = min(1.0, MAX_DIM / max(w, h, 1))
                if scale < 1.0:
                    img = img.resize((max(1, int(w * scale)), max(1, int(h * scale))), Image.LANCZOS)
                new_path = p.with_suffix(".jpg")
                img.save(new_path, "JPEG", quality=JPEG_QUALITY, optimize=True)
                size_after = new_path.stat().st_size
                total_after += size_after
                if new_path != p:
                    mapping[name] = str(new_path).replace("\\", "/")
                    p.unlink(missing_ok=True)
                converted += 1
            except Exception as e:
                print(f"  跳过 {name}: {e}")
                total_after += size_before

    MAP_JSON.write_text(json.dumps(m, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"转换 {converted} 张图片")
    print(f"总大小：{total_before/1e6:.1f}MB -> {total_after/1e6:.1f}MB")


if __name__ == "__main__":
    main()
