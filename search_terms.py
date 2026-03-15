import pymupdf
import os

pdf_path = os.path.expanduser("~/iCloud Drive (Archive)/McKinsey & Company/McKinsey Powerpoint template 2023.pdf")
doc = pymupdf.open(pdf_path)

# Specifically look for: image layouts, icon layouts, infographic layouts, 
# pie/donut charts, waterfall, area charts, line charts, combination charts
focus_terms = {
    'pie': [], 'donut': [], 'waterfall': [], 'area': [], 'line chart': [],
    'combination': [], 'with image': [], 'full bleed': [], 'half bleed': [],
    'split': [], 'icon': [], 'infographic': [], 'callout': [],
    'photo': [], 'picture': [], 'placeholder': [],
    'gauge': [], 'speedometer': [], 'thermometer': [],
    'grid': [], 'mosaic': [], 'gallery': [],
    'big stat': [], 'factoid': [],
    'process': [], 'flow': [], 'journey': [], 'map': [],
    'Harvey': [], 'harvey': [],
    'stacked': [], 'grouped': [], 'column': [],
    'bump': [], 'lollipop': [], 'butterfly': [],
    'dumbbell': [], 'dot plot': [],
    'sankey': [], 'Marimekko': [], 'marimekko': [], 'mekko': [],
}

for i in range(doc.page_count):
    page = doc[i]
    text = page.get_text().strip().lower()
    if not text:
        continue
    for term in focus_terms:
        if term.lower() in text:
            focus_terms[term].append(i+1)

doc.close()

for term, pages in sorted(focus_terms.items()):
    if pages:
        print(f"{term}: pages {pages[:20]} ({len(pages)} total)")
