import pymupdf
import os

pdf_path = os.path.expanduser("~/iCloud Drive (Archive)/McKinsey & Company/McKinsey Powerpoint template 2023.pdf")
doc = pymupdf.open(pdf_path)

os.makedirs("pdf_screenshots", exist_ok=True)

# Export key pages as images to visually inspect layouts
key_pages = [
    37,   # 4 objectives with icons and image
    55,   # Pareto chart
    85,   # Dashboard with thinkcell + quotes
    86,   # Dashboard with spider chart
    87,   # Dashboard with factoids + table
    89,   # Dashboard with gauges
    91,   # Data viz chart of 2
    92,   # Data viz chart of 3
    93,   # Data viz chart of 4 with icons
    94,   # Data viz chart of 5 with factoid
    95,   # Data viz chart with callout
    96,   # Risk matrix
    97,   # Column layout
    140,  # Pie chart
    152,  # Full bleed image
    186,  # Area chart / process
    255,  # Callout layout
    294,  # Icon layout
    348,  # Journey map
    353,  # Pie chart variant
    377,  # With image layout
    389,  # Placeholder image
    409,  # Map layout
    567,  # Bump chart
    599,  # Image layout
    603,  # Case study big stat
    655,  # Factoid quote
    676,  # Goals with image
    677,  # Big data with image 1
    678,  # Big data with 3 images
]

for pg in key_pages:
    if pg <= doc.page_count:
        page = doc[pg-1]
        pix = page.get_pixmap(matrix=pymupdf.Matrix(1.5, 1.5))
        pix.save(f"pdf_screenshots/page_{pg:03d}.png")
        print(f"Exported page {pg}")

doc.close()
print("Done!")
