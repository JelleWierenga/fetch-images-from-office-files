# fetch-images-from-office-files
The name says it all—this tool can fetch images from MS Office files, hash them, and generate a forensic report for detailed analysis.

## What does it do?
Ever needed to extract images from Office files for forensic purposes? Me too. That’s why I built this tool. It:

* Identifies MS Office file types (Word, Excel, PowerPoint) and processes them accordingly.
* Extracts images embedded in these files and hashes them using MD5 for easy identification.
* Generates a PDF report containing:
  * A summary of processed files.
  * A count of extracted images per file.
  * A detailed table of images, filenames, and their hashes.
  * A hash comparison section to check extracted hashes against a provided list.

## How does it work?
1. Copy Files:
  It starts by copying all files from a given folder into a temporary directory to ensure the originals are untouched.

2. Process MS Office Files:
* Determines the file type (Word, Excel, or PowerPoint).
* Extracts images from the corresponding media folders inside the file.

## Hash Images:
Each image is hashed using MD5, making it easy to track and compare during investigations.

## Generate a Forensic Report:
* The report includes extracted data and any matching hashes from a provided list.
* Outputs the analysis in a well-structured PDF format for forensic documentation.

## Features
* Logging: Tracks every step of the process, logging successes, warnings, and errors into `forensic_analysis.log`.
* PDF Report: A professionally formatted PDF, complete with sections for summary, hash comparison, and detailed tables.
* Error Handling: Handles unsupported file types and corrupt files gracefully.
* Customizable: Easily change paths or add support for additional file types and hash algorithms.

## How to Use
1. Clone this repo.
2. Install dependencies:
```bash
pip install pandas fpdf
```
3. Update the following paths in the script:

`usb_pad`: The folder containing your MS Office files.

`hash_vergelijking`: A text file with hash values to compare against.
4. Run the script:
```bash
python Main.py
```
5. Check the output:
* Extracted images will be in the `extracted` folder.
* The PDF report will be named `forensic_report.pdf`.

## Why did I build it?
I dont really know, but the answer would be something like below:

Extracting images from MS Office files is a common task in forensic investigations, but I needed a streamlined way to hash these images and document the results. This tool ensures accuracy and speeds up the process.

Initially, I built a simpler version of this for a test at school, but I decided to expand on it to make it more comprehensive and usable for real forensic cases.