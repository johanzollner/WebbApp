import streamlit as st
import pandas as pd
import requests
import os
from PIL import Image
from datetime import datetime

def process_excel(df):
    """
    Funktion som tar in en DataFrame och utför nedladdning samt konvertering av bilder.
    Returnerar en map för rapport och den skapade mappens namn.
    """
    downloaded_images = 0
    converted_images = 0
    failed_articles = []  # Lista med dictionaries för 'Artikelnummer' och 'Anledning'

    # Skapa en mapp för att lagra bilderna
    current_time = datetime.now().strftime('%Y%m%d_%H%M%S')
    folder_name = f"bilder_{current_time}"
    os.makedirs(folder_name, exist_ok=True)

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
            reason = str(e)
            failed_articles.append({'Artikelnummer': art_nr, 'Anledning': reason})
            continue

    # Skapa en rapport som en sträng
    total_rows = len(df)
    report_lines = []
    report_lines.append(f"Totalt antal rader i Excel: {total_rows}")
    report_lines.append(f"Totalt antal bilder nedladdade: {downloaded_images}")
    report_lines.append(f"Totalt antal bilder konverterade till webp: {converted_images}")
    report_lines.append("")
    report_lines.append("Artiklar där nedladdning/konvertering misslyckades:")
    for entry in failed_articles:
        article = entry['Artikelnummer']
        reason = entry['Anledning']
        report_lines.append(f"Artikelnummer: {article}, Anledning: {reason}")

    # Skapa en rapport som fil
    report_path = os.path.join(folder_name, "rapport.txt")
    with open(report_path, "w", encoding="utf-8") as r:
        for line in report_lines:
            r.write(line + "\n")

    # Returnera rapporten (som sträng) och sökvägen till mappen
    return "\n".join(report_lines), folder_name

def main():
    st.title("Bildhämtning och konvertering till WebP")

    # Ladda upp Excel-filen
    uploaded_file = st.file_uploader("Ladda upp en Excel-fil", type=["xlsx", "xls"])

    if uploaded_file is not None:
        # När filen är uppladdad, läs in den i en DataFrame
        try:
            df = pd.read_excel(uploaded_file, engine="openpyxl")
            st.success(f"Excel-fil uppladdad. Antal rader: {len(df)}")

            # Visa knapp för att köra process
            if st.button("Kör bildhämtning och konvertering"):
                with st.spinner("Bearbetar..."):
                    report, folder_name = process_excel(df)
                st.success("Bearbetning klar!")
                
                # Visa rapport
                st.subheader("Rapport")
                st.text(report)

                # Möjlighet att ladda ned rapportfilen
                with open(os.path.join(folder_name, "rapport.txt"), "rb") as f:
                    st.download_button(
                        label="Ladda ner rapport",
                        data=f,
                        file_name="rapport.txt",
                        mime="text/plain"
                    )

        except Exception as e:
            st.error(f"Ett fel uppstod vid inläsning: {e}")

if __name__ == "__main__":
    main()
