"""
一次性修复脚本：把 res/product_images/_pdf_snapshot 下已生成的、体积异常大的
PNG（部分单张超过100MB，原因是渲染时 dpi=110 对某些物理页面尺寸异常大的PDF
产生了巨幅像素图）原地缩放到最长边 <= 1000px，避免 dashboard.py 把它们
base64 内嵌进同一个HTML字符串时导致 MemoryError。

用法：python src/_shrink_pdf_snapshots.py
"""
from pathlib import Path
from PIL import Image

DIR = Path(__file__).parent.parent / "res" / "product_images" / "_pdf_snapshot"
MAX_DIM = 1000

def main():
    files = list(DIR.rglob("*.png"))
    print(f"共 {len(files)} 个文件")
    shrunk = 0
    total_before = 0
    total_after = 0
    for f in files:
        size_before = f.stat().st_size
        total_before += size_before
        try:
            img = Image.open(f)
            w, h = img.size
            m = max(w, h)
            if m > MAX_DIM or size_before > 1_000_000:
                scale = MAX_DIM / m
                new_size = (max(1, int(w * scale)), max(1, int(h * scale)))
                img = img.convert("RGB").resize(new_size, Image.LANCZOS)
                img.save(f, "PNG", optimize=True)
                size_after = f.stat().st_size
                total_after += size_after
                shrunk += 1
                print(f"  {f.name}: {size_before/1e6:.1f}MB -> {size_after/1e6:.2f}MB ({w}x{h} -> {new_size[0]}x{new_size[1]})")
            else:
                total_after += size_before
        except Exception as e:
            print(f"  跳过 {f.name}: {e}")
            total_after += size_before

    print(f"\n处理完成：缩放了 {shrunk}/{len(files)} 个文件")
    print(f"总大小：{total_before/1e6:.1f}MB -> {total_after/1e6:.1f}MB")


if __name__ == "__main__":
    main()
