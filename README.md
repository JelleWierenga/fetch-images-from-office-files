# fetch-images-from-office-files
the name says it all, this tool can fetch images from office files and store them in a db and hash them for later forensisc research


Did you ever need to extract images from office files?
Well i did, and i needed to do it in a way that i can hash the images and store them in a database for later forensic research.
The tool needs to be able to identify the office file and handle it accordingly.
It stores the images in a database with a sha256 hash of the image as the key.
and a seperate database with the file names and a id that is linked to the images so you can always see whice image came from what file.
