import os
import shutil
import pandas as pd
import hashlib
import zipfile
from fpdf import FPDF
from datetime import date
import logging

# logging instellen
logging.basicConfig(
    filename='forensic_analysis.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

logging.info("Script started")

# Oorspronkelijke folder met bestanden
usb_pad = "C:/Users/jelle/OneDrive - Hogeschool Leiden/ifipop/Toets/files"
hash_vergelijking = "hashes.txt"

# Maak een tijdelijke map voor kopieën
tijdelijk_mapje = "temp_files"
os.makedirs(tijdelijk_mapje, exist_ok=True)

# Kopieer alle bestanden uit file_path naar de tijdelijke map
for file in os.listdir(usb_pad):
    try:
        shutil.copy(os.path.join(usb_pad, file), os.path.join(tijdelijk_mapje, file))
        logging.info(f"Copied file {file} to temporary folder")
    except Exception as e:
        logging.error(f"Error copying file {file}: {e}")

# DataFrame om resultaten op te slaan
df = pd.DataFrame(columns=["bestand_naam", "foto_naam", "foto_hash"])

# Functie om bestandstype en pad naar afbeeldingen te bepalen
def zoek_bestand_types(file):
    if file.endswith(".docx") or file.endswith(".doc"):
        logging.info(f"File {file} identified as Word document.")
        return "word", "docx", "word/media/"
    elif file.endswith(".pptx") or file.endswith(".ppt"):
        logging.info(f"File {file} identified as PowerPoint document.")
        return "powerpoint", "pptx", "ppt/media/"
    elif file.endswith(".xlsx") or file.endswith(".xls"):
        logging.info(f"File {file} identified as Excel document.")
        return "excel", "xlsx", "xl/media/"
    else:
        logging.warning(f"Unsupported file type for file: {file}")
        return None

# Functie om afbeeldingen uit MS Office-bestanden te halen en hashen
def fetch_fotos_van_mso_bestanden(bestand, pad, tijdelijke_folder):
    global df
    elke_extensie = [".png", ".jpg", ".jpeg", ".gif", ".RAW"]
    try:
        with zipfile.ZipFile(f"{tijdelijke_folder}/{bestand}") as extractie:
            logging.info(f"Opened {bestand} for image extraction.")
            for i in range(len(extractie.namelist())):
                for extensie in elke_extensie:
                    try:
                        image_path = pad + f"image{str(i)}{extensie}"
                        extractie.extract(image_path, "extracted")
                        hash = hashlib.md5(extractie.read(image_path)).hexdigest()
                        nieuwe_kolom_in_df = pd.DataFrame({"file_name": bestand, "image_name": f"image{i + 1}{extensie}", "image_hash": hash}, index=[0])
                        df = pd.concat([df, nieuwe_kolom_in_df], ignore_index=True)
                        logging.info(f"Found and hashed image {image_path} in {bestand} with hash {hash}.")
                    except KeyError:
                        logging.warning(f"Image {image_path} not found in {bestand}. Continuing to next image.")
                        break
    except zipfile.BadZipFile:
        logging.error(f"{bestand} is een leeg of corrupt bestand")

# Loop door bestanden in de tijdelijke map
for bestand in os.listdir(tijdelijk_mapje):
    info = zoek_bestand_types(bestand)
    if info is not None:
        path = info[2]
        fetch_fotos_van_mso_bestanden(bestand, path, tijdelijk_mapje)

# Functie om PDF op te stellen met de DataFrame-tabel
class PDF(FPDF):
    def header(self):
        if self.page_no() > 1:  # Geen header op de eerste pagina
            self.set_font("Arial", "B", 12)
            self.cell(0, 10, "Forensisch rapport", 0, 1, "C")
            self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font("Arial", "I", 8)
        self.cell(0, 10, f'bladzijde {self.page_no()}', 0, 0, 'C')

    def add_cover_page(self):
        self.add_page()
        self.set_font("Arial", "B", 24)
        self.cell(0, 10, "Forenscih MS office rapport", 0, 1, "C")
        self.ln(20)
        self.set_font("Arial", size=16)
        self.cell(0, 10, "Generated by: Jelle's geweldige forensische tool", 0, 1, "C")
        self.cell(0, 10, f"Date: {date.today()}", 0, 1, "C")  # Pas de datum aan naar de huidige datum indien nodig
        self.ln(30)
        self.set_font("Arial", size=12)
        self.multi_cell(0, 10, "Dit rapport bevat de analyse van de MS office bestanden en de afbeeldingen die daar in staan, inclusief een vergelijking van afbeeldingshashes. Elke sectie biedt inzicht in de inhoud van de bestanden en mogelijke hash-overeenkomsten.", 0, "C")
        self.ln(30)
        self.cell(0, 10, "Kwalificatie: GEHEIM", 0, 1, "C")
        logging.info("Cover page added to PDF report.")

    def add_files_with_images(self, df):
        self.add_page()
        self.set_font("Arial", "B", 12)
        self.cell(200, 10, txt="Bestanden met foto's", ln=True, align="L")
        self.set_font("Arial", size=10)
        files_with_image = df['file_name'].unique()
        for file in files_with_image:
            self.cell(200, 10, txt=file, ln=True, align="L")
        self.ln(10)
        logging.info("Section 'Files with Images' added to PDF report.")

    def add_files_image_count(self, df):
        self.set_font("Arial", "B", 12)
        self.cell(200, 10, txt="Bestanden met aantal foto's", ln=True, align="L")
        self.set_font("Arial", size=10)
        count = df.groupby('file_name').size().reset_index(name='count')
        for _, row in count.iterrows():
            self.cell(200, 10, txt=f"{row['file_name']} heeft {row['count']} foto's", ln=True, align="L")
        self.ln(10)
        logging.info("Section 'Files and Image Count' added to PDF report.")

    def add_table(self, df):
        self.add_page()
        self.set_font("Arial", "B", 10)
        col_widths = [50, 50, 90]
        headers = ["Bestand naam", "Foto naam", "Foto hash(md5)"]
        for header, width in zip(headers, col_widths):
            self.cell(width, 10, header, border=1, align="C")
        self.ln()

        self.set_font("Arial", size=10)
        for index, row in df.iterrows():
            self.cell(col_widths[0], 10, row['file_name'], border=1)
            self.cell(col_widths[1], 10, row['image_name'], border=1)
            self.cell(col_widths[2], 10, row['image_hash'], border=1)
            self.ln()
        logging.info("Detailed table of image hashes added to PDF report.")

    def add_hash_comparison(self, df, hash_file_path):
        self.add_page()
        self.set_font("Arial", "B", 12)
        self.cell(200, 10, txt="Hash vergelijking", ln=True, align="L")
        self.set_font("Arial", size=10)
        try:
            with open(hash_file_path, 'r') as hash_file:
                for line in hash_file:
                    image_name, file_hash = line.strip().split(": ")
                    if file_hash in df['image_hash'].values:
                        matching_row = df[df['image_hash'] == file_hash].iloc[0]
                        self.cell(200, 10, txt=f"{image_name} heeft hash {file_hash} en zit in de data.", ln=True, align="L")
                        self.cell(200, 10, txt=f"Details - Bestand: {matching_row['file_name']}, Foto naam: {matching_row['image_name']}", ln=True, align="L")
                        logging.info(f"Hash match found for {image_name} in {matching_row['file_name']}")
                    else:
                        self.cell(200, 10, txt=f"{image_name} heeft hash {file_hash} en zit niet in de data.", ln=True, align="L")
                        logging.info(f"No hash match found for {image_name}")
        except Exception as e:
            logging.error(f"Error reading hash file: {e}")

# PDF genereren
pdf = PDF()
pdf.add_cover_page()  # Voeg een voorpagina toe
pdf.add_files_with_images(df)  # Sectie 1: Bestanden met afbeeldingen
pdf.add_files_image_count(df)  # Sectie 2: Bestanden en aantal afbeeldingen
pdf.add_table(df)  # Sectie 3: Gedetailleerde tabel met hashes
pdf.add_hash_comparison(df, hash_vergelijking)  # Sectie 4: Hash vergelijking

# PDF opslaan
pdf.output("forensic_report.pdf")
logging.info("PDF report generated successfully")
logging.info("Script ended")