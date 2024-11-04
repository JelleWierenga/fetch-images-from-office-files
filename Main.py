import hashlib
import zipfile
from docx import Document
from PIL import Image
import pandas as pd

df = pd.DataFrame(columns=["id", "hash", "foto naam"])

with zipfile.ZipFile('files/Portfolio.docx', 'r') as docx:
    path = "word/media"
    for i in range(200):
        try:
            docx.extract(path + "/image" + str(i+1) + ".png", "extracted")
        except KeyError:
            break

import hashlib
import zipfile
from docx import Document
from PIL import Image
import pandas as pd

df = pd.DataFrame(columns=["id", "hash", "foto naam"])

with zipfile.ZipFile('files/Portfolio.docx', 'r') as docx:
    path = "word/media"
    for i in range(200):
        try:
            docx.extract(path + "/image" + str(i+1) + ".png", "extracted")
        except KeyError:
            break

file = open("Hashes.txt", "w")
for x in range(1, 200):
    try:
        folder = f"extracted/word/media/image{x}.png"
        with open(folder, "rb") as f:
            hash = hashlib.sha256(f.read()).hexdigest()
            print(hash)
            file.write(f"Portfolio.docx: {hash}\n")
    except FileNotFoundError:
        break



file.close()


file.close()