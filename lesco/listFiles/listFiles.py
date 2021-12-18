
#import OS
import os
 
for x in os.listdir():
    if x.endswith(".docx"):
        # Prints only text file present in My Folder
        print(x)