import pandas as pd
import requests
import os
from PIL import Image
from datetime import datetime

print("Script starting...")

# Läs in Excel-filen
df = pd.read_excel('din_excel_fil.xlsx', engine='openpyxl')

downloaded_images = 0
converted_images = 0
failed_articles = []  # Lista med dictionaries för 'Artikelnummer' och 'Anledning'

# Skapa en mapp för att lagra bilderna
current_time = datetime.now().strftime('%Y%m%d_%H%M%S')
folder_name = f"bilder_{current_time}"
os.makedirs(folder_name)

# Kolla varje rad i DataFrame
for index, row in df.iterrows():
    url = row['Bildlänk']
    art_nr = row['Artikelnummer']
    
    # Initiera variabel för anledningen till fel
    reason = ''
    
    # Kontrollera om URL är ogiltig eller indikerar ingen bild
    if pd.isnull(url) or url == 'No picture available' or not str(url).startswith(('http://', 'https://')):
        if pd.isnull(url):
            reason = 'Ingen bildlänk angiven'
        elif url == 'No picture available':
            reason = 'Ingen bild tillgänglig'
        else:
            reason = 'Ogiltig URL'
        failed_articles.append({'Artikelnummer': art_nr, 'Anledning': reason})
        continue

    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()  # Kontrollera om förfrågan lyckades

        # Bestäm filändelsen baserat på innehållstyp
        content_type = response.headers.get('Content-Type', '')
        if 'image/jpeg' in content_type:
            file_extension = '.jpeg'
        elif 'image/png' in content_type:
            file_extension = '.png'
        elif 'image/webp' in content_type:
            file_extension = '.webp'
        else:
            file_extension = '.jpeg'  # Standard till JPEG om okänt

        file_name = f"{folder_name}/{art_nr}{file_extension}"
        with open(file_name, 'wb') as f:
            f.write(response.content)
            downloaded_images += 1

        # Konvertera till WebP och ta bort originalbilden
        with Image.open(file_name) as img:
            webp_name = f"{folder_name}/{art_nr}.webp"
            img.save(webp_name, "WEBP")
            os.remove(file_name)
            converted_images += 1
    except Exception as e:
        print(f"Misslyckades med att bearbeta artikel {art_nr}: {e}")
        reason = str(e)
        failed_articles.append({'Artikelnummer': art_nr, 'Anledning': reason})
        continue

# Skapa en rapport
with open(f"{folder_name}/rapport.txt", "w") as report:
    report.write(f"Totalt antal rader i Excel: {len(df)}\n")
    report.write(f"Totalt antal bilder nedladdade: {downloaded_images}\n")
    report.write(f"Totalt antal bilder konverterade till webp: {converted_images}\n\n")
    report.write("Artiklar där nedladdning misslyckades:\n")
    for entry in failed_articles:
        article = entry['Artikelnummer']
        reason = entry['Anledning']
        report.write(f"Artikelnummer: {article}, Anledning: {reason}\n")

print(f"Totalt antal rader i Excel: {len(df)}")
print(f"Totalt antal bilder nedladdade: {downloaded_images}")
print(f"Totalt antal bilder konverterade till webp: {converted_images}")

print("Script finished.")
