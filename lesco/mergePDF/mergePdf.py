#import OS
import os
from PyPDF2 import PdfFileMerger, PdfFileReader

# list of files in folder
filesList = list()
for x in os.listdir():
    if x.endswith(".pdf"):
        # Prints only text file present in My Folder
        filesList.append(x)
y = (filesList[0])

# merge Files
# Call the PdfFileMerger
mergedObject = PdfFileMerger()
for i in filesList:
    mergedObject.append(PdfFileReader(i, 'rb'))
 
# Write all the files into a file which is named as shown below
mergedObject.write(f'{y[0:8]}-bills.pdf')