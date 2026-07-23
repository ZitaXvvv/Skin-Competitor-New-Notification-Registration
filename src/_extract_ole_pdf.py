import zipfile, olefile

XLSX = r"C:\Users\xie.x.3\AppData\Local\Temp\1\韩束.xlsx"
z = zipfile.ZipFile(XLSX)

OUT = r"C:\Users\xie.x.3\AppData\Local\Temp"

for i in range(1, 7):
    data = z.read(f"xl/embeddings/oleObject{i}.bin")
    tmp = rf"{OUT}\_ole{i}.bin"
    with open(tmp, "wb") as f:
        f.write(data)
    ole = olefile.OleFileIO(tmp)
    native = ole.openstream("\x01Ole10Native").read()
    ole.close()
    print(f"oleObject{i}: native stream size={len(native)}  header bytes={native[:80]}")
    idx = native.find(b"%PDF")
    print("  %PDF found at offset:", idx)
    if idx >= 0:
        pdf_bytes = native[idx:]
        out_pdf = rf"{OUT}\_extracted_ole{i}.pdf"
        with open(out_pdf, "wb") as f:
            f.write(pdf_bytes)
        print("  wrote", out_pdf, "size", len(pdf_bytes))
