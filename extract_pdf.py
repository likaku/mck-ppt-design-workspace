import pymupdf
import os

pdf_path = os.path.expanduser("~/iCloud Drive (Archive)/McKinsey & Company/McKinsey Powerpoint template 2023.pdf")
if not os.path.exists(pdf_path):
    for alt in [
        os.path.expanduser("~/Library/Mobile Documents/com~apple~CloudDocs/McKinsey & Company/McKinsey Powerpoint template 2023.pdf"),
        os.path.expanduser("~/iCloud Drive/McKinsey & Company/McKinsey Powerpoint template 2023.pdf"),
    ]:
        if os.path.exists(alt):
            pdf_path = alt
            break

print(f"PDF path: {pdf_path}")
print(f"Exists: {os.path.exists(pdf_path)}")

if os.path.exists(pdf_path):
    doc = pymupdf.open(pdf_path)
    print(f"Total pages: {doc.page_count}")
    
    for i in range(doc.page_count):
        page = doc[i]
        text = page.get_text().strip()
        if text:
            preview = text[:400].replace('\n', ' | ')
            print(f"\n--- Page {i+1} ---")
            print(preview)
    doc.close()
else:
    print("PDF file not found!")
    base = os.path.expanduser("~/iCloud Drive (Archive)")
    if os.path.exists(base):
        for item in os.listdir(base):
            print(f"  {item}")
