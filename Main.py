import os
import pandas as pd
import hashlib
import zipfile

df = pd.DataFrame(columns=["file_name", "image_name", "image_hash"])

# vindt de bestand type om te gebruiken in data en foto's extracten
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
    # elif ".odt" in file:
    #     file_name = "libreoffice_text"
    #     file_extension = "odt"
    #     path = "Pictures/"
    #     return file_name, file_extension, path
    # elif ".ods" in file:
    #     file_name = "libreoffice_sheet"
    #     file_extension = "ods"
    #     path = "Pictures/"
    #     return file_name, file_extension, path
    # elif ".odp" in file:
    #     file_name = "libreoffice_presentation"
    #     file_extension = "odp"
    #     path = "Pictures/"
    #     return file_name, file_extension, path
    # return None

def fetch_images_mso(files, path):
    global df
    extensions_list = [".png", ".jpg", ".jpeg"]
    try:
        with zipfile.ZipFile(f"files/{files}") as extraction:
            # extraction.extractall("extracted")
            # full_path = os.path.join("extracted", path)

            for i in range(len(extraction.namelist())):
                for extensions in extensions_list:
                    try:
                        # for files in os.path.join("extracted", full_path):
                        #     print(files)
                        extraction.extract(path + "image" + str(i + 1) + f"{extensions}", "extracted")
                        hash = hashlib.sha256(extraction.read(path + "image" + str(i + 1) + f"{extensions}")).hexdigest()
                        new_column = pd.DataFrame({"file_name": files, "image_name": f"image{i + 1}{extensions}", "image_hash": hash}, index=[0])
                        df = pd.concat([df, new_column], ignore_index=True)
                    except KeyError:
                        break
    except zipfile.BadZipFile:
        print(f"{files} is een leeg bestand")




for files in os.listdir("files"):
    # print(files)
    info = get_file_type(files)
    if info is not None:
        # print(info)
        path = info[2]
        fetch_images_mso(files, path)
print(df)
df.to_csv('output1.txt', sep='\t', index=False)

report = open('report.txt', 'w')
report.write("Files with images:\n")
files_with_image = df.file_name.unique()
for file in files_with_image:
    report.write(file + "\n")

report.write("\n\nFiles with image count:\n")
count = df.groupby('file_name').size().reset_index(name='count')
for _, row in count.iterrows():
    report.write(f"{row['file_name']} has {row['count']} images linked\n")

