import os
import pandas as pd
import hashlib
import zipfile
from docx import Document

def fetch_images():
    with zipfile.ZipFile('files/Portfolio.docx', 'r') as docx:
        path = "word/media"
        for i in range(200):
            try:
                docx.extract(path + "/image" + str(i + 1) + ".png", "extracted")
            except KeyError:
                break

df_files = pd.DataFrame(columns=["file_id", "file_name"])
df_images = pd.DataFrame(columns=["image_id", "file_id", "image_name", "image_hash(sha256)"])

def hash_file(df_files):
    df_images = pd.DataFrame(columns=["image_id", "file_id", "image_name", "image_hash(sha256)"])

    images = os.listdir("extracted/word/media")

    files_sorted = sorted(images, key=lambda x: int(x.split('image')[1].split('.png')[0]))

    print(files_sorted)

    for file in files_sorted:
        file_path = f"extracted/word/media/{file}"
        with open(file_path, "rb") as f:
            file_hash = hashlib.sha256(f.read()).hexdigest()
        new_row = pd.DataFrame({"image_id": [len(df_images) + 1],"file_id": df_files["file_id"], "image_name": [file], "image_hash(sha256)": file_hash})
        df_images = pd.concat([df_images, new_row], ignore_index=True)

    return df_images

def fetch_file_name():
    df_files = pd.DataFrame(columns=["file_id", "file_name"])

    files = os.listdir("files")
    new_row = pd.DataFrame({"file_id": [len(df_files) + 1], "file_name": [files[0]]})
    df_files = pd.concat([df_files, new_row], ignore_index=True)
    return df_files

fetch_images()
df_files = fetch_file_name()
df_images = hash_file(df_files)

print(df_files)
print(df_images)