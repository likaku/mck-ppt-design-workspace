import pymupdf
import os

pdf_path = os.path.expanduser("~/iCloud Drive (Archive)/McKinsey & Company/McKinsey Powerpoint template 2023.pdf")
doc = pymupdf.open(pdf_path)

# Key pages representing NEW layout types for v1.8
key_pages = [
    # Pie/Donut charts
    140, 353, 354, 355, 357,
    # Data viz combos
    91, 92, 93, 94, 95, 96,
    # Dashboard pages
    85, 86, 87, 88, 89, 90,
    # Area charts / process with area
    186, 187, 188, 189, 190, 191, 192, 193, 194, 195,
    # Callout layouts
    255, 299, 300, 301,
    # Icon-based layouts
    294, 295, 296, 297, 298,
    # Image layouts
    152, 377, 471, 599, 675, 676, 677, 678,
    # Journey/experience maps  
    348, 349, 350, 371, 372,
    # Column layouts / frameworks
    97, 98, 99, 100, 137, 138,
    # Flow diagrams
    344, 345, 347, 415, 416, 424, 425,
    # Gauge
    89,
    # Factoid big stat
    603, 655,
    # Process special
    362, 363, 364, 369,
    # Maps
    409, 410, 411,
    # Bump chart
    567,
    # Placeholder images
    389, 392, 394, 395,
    # Big stat Georgia
    28, 33, 34, 37, 39,
]

key_pages = sorted(set(key_pages))

for pg in key_pages:
    if pg <= doc.page_count:
        page = doc[pg-1]
        text = page.get_text().strip()
        if text:
            print(f"\n===== PAGE {pg} =====")
            print(text[:400])

doc.close()
