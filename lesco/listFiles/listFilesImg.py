
#import OS
import os
from csv import writer

def append_referenceNo_as_row(file_name, referenceNo):
    # Open file in append mode
    with open(file_name, 'a+', newline='') as write_obj:
        # Create a writer object from csv module
        csv_writer = writer(write_obj)
        # Add contents of list as last row in the csv file
        csv_writer.writerow(referenceNo)
 
for filesNames in os.listdir():
    if filesNames.endswith("1E.jpg"):
        fileNamesExtensionRemoved = filesNames.replace("1E.jpg","U")
        fileNamesYearMonthRemoved = fileNamesExtensionRemoved.replace(fileNamesExtensionRemoved[0:7],fileNamesExtensionRemoved[6])
        BatchNo = fileNamesYearMonthRemoved[0:2]
        append_referenceNo_as_row(f'{BatchNo}.csv', [fileNamesYearMonthRemoved])
    else:
        pass
print("****************************************************************************************")
print("***                                                                                  ***")
print("***   CSV file created containg All reference numbers. Kindly cross Check with CP.   ***")
print("***                                                                                  ***")
print("****************************************************************************************")