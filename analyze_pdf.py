import pymupdf
import os, json

pdf_path = os.path.expanduser("~/iCloud Drive (Archive)/McKinsey & Company/McKinsey Powerpoint template 2023.pdf")
doc = pymupdf.open(pdf_path)

# Keywords for chart/image/data visualization layouts
chart_keywords = ['chart', 'graph', 'bar', 'pie', 'donut', 'line chart', 'area chart', 'waterfall',
                  'bubble', 'scatter', 'histogram', 'sparkline', 'gauge', 'meter', 'thermometer',
                  'heat map', 'heatmap', 'treemap', 'radar', 'spider']
image_keywords = ['image', 'photo', 'picture', 'placeholder', 'full bleed', 'background image',
                  'with image', 'photography', 'visual', 'icon', 'illustration']
data_keywords = ['data', 'dashboard', 'KPI', 'metric', 'stats', 'statistic', 'number',
                 'percent', 'progress', 'trend', 'comparison', 'benchmark']
layout_keywords = ['template', 'layout', 'slide', 'divider', 'cover', 'agenda', 'summary',
                   'grid', 'column', 'matrix', 'framework', 'process', 'timeline',
                   'infographic', 'callout', 'quote']

results = {'chart': [], 'image': [], 'data': [], 'layout': []}

for i in range(doc.page_count):
    page = doc[i]
    text = page.get_text().lower().strip()
    if not text:
        continue
    
    for kw in chart_keywords:
        if kw in text:
            results['chart'].append((i+1, text[:200].replace('\n', ' | ')))
            break
    
    for kw in image_keywords:
        if kw in text:
            results['image'].append((i+1, text[:200].replace('\n', ' | ')))
            break
    
    for kw in data_keywords:
        if kw in text:
            results['data'].append((i+1, text[:200].replace('\n', ' | ')))
            break

doc.close()

print(f"=== CHART pages ({len(results['chart'])}) ===")
for pg, txt in results['chart']:
    print(f"  Page {pg}: {txt[:150]}")

print(f"\n=== IMAGE pages ({len(results['image'])}) ===")
for pg, txt in results['image']:
    print(f"  Page {pg}: {txt[:150]}")

print(f"\n=== DATA pages ({len(results['data'])}) ===")
for pg, txt in results['data'][:50]:
    print(f"  Page {pg}: {txt[:150]}")
