import os
import shutil
import pandas as pd
import hashlib
import zipfile

# Oorspronkelijke folder met bestanden
file_path = "path/to/files"

# Maak een tijdelijke map voor kopieën
temp_folder = "temp_files"
os.makedirs(temp_folder, exist_ok=True)

# Kopieer alle bestanden uit file_path naar de tijdelijke map
for file in os.listdir(file_path):
    shutil.copy(os.path.join(file_path, file), os.path.join(temp_folder, file))

# DataFrame om resultaten op te slaan
df = pd.DataFrame(columns=["file_name", "image_name", "image_hash"])

# Functie om bestandstype en pad naar afbeeldingen te bepalen
def get_file_type(file):
    if ".docx" in file or ".doc" in file:
        file_name = "word"
        file_extension = "docx"
        path = "word/media/"
        return file_name, file_extension, path
    elif ".pptx" in file or ".ppt" in file:
        file_name = "powerpoint"
        file_extension = "pptx"
        path = "ppt/media/"
        return file_name, file_extension, path
    elif ".xlsx" in file or ".xls" in file:
        file_name = "excel"
        file_extension = "xlsx"
        path = "xl/media/"
        return file_name, file_extension, path

# Functie om afbeeldingen uit MS Office-bestanden te halen en hashen
def fetch_images_mso(files, path, temp_folder):
    global df
    extensions_list = [".png", ".jpg", ".jpeg"]
    try:
        with zipfile.ZipFile(f"{temp_folder}/{files}") as extraction:
            for i in range(len(extraction.namelist())):
                for extensions in extensions_list:
                    try:
                        extraction.extract(path + "image" + str(i + 1) + f"{extensions}", "extracted")
                        hash = hashlib.sha256(extraction.read(path + "image" + str(i + 1) + f"{extensions}")).hexdigest()
                        new_column = pd.DataFrame({"file_name": files, "image_name": f"image{i + 1}{extensions}", "image_hash": hash}, index=[0])
                        df = pd.concat([df, new_column], ignore_index=True)
                    except KeyError:
                        break
    except zipfile.BadZipFile:
        print(f"{files} is een leeg bestand")

# Loop door bestanden in de tijdelijke map
for files in os.listdir(temp_folder):
    info = get_file_type(files)
    if info is not None:
        path = info[2]
        fetch_images_mso(files, path, temp_folder)

# Resultaten opslaan in CSV
print(df)
df.to_csv('output1.txt', sep='\t', index=False)

# Rapport maken
with open('report.txt', 'w') as report:
    #rapporteer bestanden die afbeeldingen hebben
    report.write("Files with images:\n")
    files_with_image = df.file_name.unique()
    for file in files_with_image:
        report.write(file + "\n")

    # Rapporteer bestanden met aantal afbeeldingen
    report.write("\n\nFiles with image count:\n")
    count = df.groupby('file_name').size().reset_index(name='count')
    for _, row in count.iterrows():
        report.write(f"{row['file_name']} has {row['count']} images linked\n")

    # rapporteer hashes die overeenkomen
    report.write("\n\nHash Comparison:\n")
    with open("hashes.txt", 'r') as hash_file:  # Bestand met te vergelijken hashes
        for line in hash_file:
            image_name, file_hash = line.strip().split(": ")

            # Controleer of de hash bestaat in de dataframe
            if file_hash in df['image_hash'].values:
                # Zoek de bijbehorende rij in de dataframe
                matching_row = df[df['image_hash'] == file_hash].iloc[0]
                report.write(f"{image_name} with hash {file_hash} exists in the DataFrame.\n")
                report.write(f"Details \nFile: {matching_row['file_name']}, Image Name: {matching_row['image_name']}\n")
            else:
                report.write(f"{image_name} with hash {file_hash} does NOT exist in the DataFrame.\n")

# shutil.rmtree(temp_folder)
