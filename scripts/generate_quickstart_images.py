from PIL import Image, ImageDraw, ImageFont
from pathlib import Path

root = Path(__file__).resolve().parent.parent / 'docs' / 'images'
root.mkdir(parents=True, exist_ok=True)

try:
    font_bold = ImageFont.truetype('arialbd.ttf', 28)
    font_regular = ImageFont.truetype('arial.ttf', 18)
    font_small = ImageFont.truetype('arial.ttf', 16)
except Exception:
    font_bold = ImageFont.load_default()
    font_regular = ImageFont.load_default()
    font_small = ImageFont.load_default()

steps = [
    ('Open BOM of data', 'Start met het laden van je projectgegevens en controleer de input.'),
    ('Controleer instellingen', 'Bekijk de export- en PDF-instellingen voordat je verdergaat.'),
    ('Genereer documenten', 'Klik op de exportknop om bestelbonnen of PDF dossiers te maken.'),
    ('Controleer output', 'Open de gegenereerde bestanden in de exportmap en controleer ze.'),
]

for index, (title, caption) in enumerate(steps, start=1):
    width, height = 1200, 700
    image = Image.new('RGB', (width, height), '#F8FAFC')
    draw = ImageDraw.Draw(image)

    # Header
    draw.rectangle([0, 0, width, 110], fill='#1D4ED8')
    draw.text((32, 32), 'Filehopper', font=font_bold, fill='white')
    draw.text((32, 68), f'Stap {index}: {title}', font=font_regular, fill='white')

    # Main screenshot area
    margin = 40
    top = 140
    draw.rectangle([margin, top, width-margin, height-margin], fill='white', outline='#CBD5E1', width=2)

    block_y = top + 24
    draw.rectangle([margin+24, block_y, width-margin-24, block_y+180], fill='#E2E8F0', outline='#94A3B8', width=2)
    draw.text((margin+40, block_y+18), 'Menu / Tabblad overzicht', font=font_bold, fill='#0F172A')
    draw.text((margin+40, block_y+58), 'Selecteer je tabblad om naar de juiste workflow te gaan.', font=font_regular, fill='#334155')

    block_y += 220
    draw.rectangle([margin+24, block_y, width-margin-24, block_y+240], fill='#F1F5F9', outline='#94A3B8', width=1)
    draw.text((margin+40, block_y+18), 'Actiegebied', font=font_bold, fill='#0F172A')
    draw.text((margin+40, block_y+58), caption, font=font_regular, fill='#334155')
    draw.rectangle([margin+60, block_y+110, margin+280, block_y+150], fill='#2563EB')
    draw.text((margin+76, block_y+118), 'Laad BOM', font=font_regular, fill='white')
    draw.rectangle([margin+300, block_y+110, margin+520, block_y+150], fill='#10B981')
    draw.text((margin+316, block_y+118), 'Start export', font=font_regular, fill='white')

    side_x = width - margin - 300
    draw.rectangle([side_x, block_y, side_x+260, block_y+220], fill='#DBEAFE', outline='#93C5FD', width=1)
    draw.text((side_x+16, block_y+18), 'Tip:', font=font_bold, fill='#1D4ED8')
    draw.text((side_x+16, block_y+50), 'Werk met een huidige dataset en controleer labels.', font=font_small, fill='#1E293B')

    path = root / f'quickstart_step_{index}.png'
    image.save(path)
    print(f'Saved {path}')
