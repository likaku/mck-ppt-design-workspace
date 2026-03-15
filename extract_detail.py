import pymupdf
import os

pdf_path = os.path.expanduser("~/iCloud Drive (Archive)/McKinsey & Company/McKinsey Powerpoint template 2023.pdf")
doc = pymupdf.open(pdf_path)

# Extract detailed text for key chart/image/data pages
key_pages = [
    # Chart pages
    55, 56, 57, 58, 59, 82, 83, 86, 87, 89, 91, 92, 93, 94,
    # More chart/data pages
    101, 102, 103, 104, 105, 108, 110, 111, 112, 113, 114,
    # Image pages  
    95, 96, 117, 118, 119, 120, 
    # Also check pages around image keywords
    133, 134, 135,
    # Check nearby pages for context
    60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70,
    71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81,
    84, 85, 88, 90, 97, 98, 99, 100,
    106, 107, 109, 115, 116,
    # Image-heavy pages from the end
    675, 676, 677, 678, 679,
    # Check for infographic/icon pages
    121, 122, 123, 124, 125, 126, 127, 128, 129, 130,
    131, 132,
    # First 50 pages structure
    14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25,
    26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40,
    41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54,
]

key_pages = sorted(set(key_pages))

for pg in key_pages:
    if pg <= doc.page_count:
        page = doc[pg-1]
        text = page.get_text().strip()
        if text:
            print(f"\n===== PAGE {pg} =====")
            print(text[:500])

doc.close()
