import fitz  # PyMuPDF

for i in range(1, 7):
    path = rf"C:\Users\xie.x.3\AppData\Local\Temp\_extracted_ole{i}.pdf"
    doc = fitz.open(path)
    print(f"ole{i}: pages={doc.page_count}")
    page = doc[0]
    pix = page.get_pixmap(dpi=150)
    out = rf"C:\Users\xie.x.3\AppData\Local\Temp\_ole{i}_page1.png"
    pix.save(out)
    print("  saved", out, pix.width, pix.height)
    doc.close()
