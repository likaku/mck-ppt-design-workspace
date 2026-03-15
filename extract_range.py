import pymupdf
import os

pdf_path = os.path.expanduser("~/iCloud Drive (Archive)/McKinsey & Company/McKinsey Powerpoint template 2023.pdf")
doc = pymupdf.open(pdf_path)

# Focus on pages 85-200 and 600-679 for chart/image/data heavy content
ranges = list(range(85, 200)) + list(range(200, 350)) + list(range(600, 680))

for pg in ranges:
    if pg <= doc.page_count:
        page = doc[pg-1]
        text = page.get_text().strip()
        if text:
            t = text[:300].replace('\n', ' | ')
            print(f"P{pg}: {t}")

doc.close()
