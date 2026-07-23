import fitz

for i in range(1, 7):
    path = rf"C:\Users\xie.x.3\AppData\Local\Temp\_extracted_ole{i}.pdf"
    doc = fitz.open(path)
    for pno in range(doc.page_count):
        page = doc[pno]
        imgs = page.get_images(full=True)
        print(f"ole{i} page{pno}: {len(imgs)} images")
        for idx, img in enumerate(imgs):
            xref = img[0]
            base = doc.extract_image(xref)
            w, h = base.get("width"), base.get("height")
            print(f"    img{idx} xref={xref} {w}x{h} ext={base.get('ext')}")
    doc.close()
