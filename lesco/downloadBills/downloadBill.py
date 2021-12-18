#!/usr/bin/env python
# coding: utf-8

# In[ ]:

### Imports ###
import csv
import pdfkit

## Initilization
path_wkhtmltopdf = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'
config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)
options = {
    'page-size': 'A4'
}

##########################################################
##########################################################
##########################################################
##########################################################
##########################################################
##########################################################
### Enter Reference No. ###
print("**********Are you going to enter the Reference No. 'Manually' or going to use 'CSV file'**********")
print('Enter "M" for manual')
print('Enter "C" for CSV file')
e_m_mode = 0
R = list()
while e_m_mode == 0:
    entry_mode = input('Entry Mode ? ')
    if entry_mode.upper() == "M" or entry_mode.upper() == "C":
        if entry_mode.upper() == "M":
            check_R_m = 0
            while check_R_m == 0:
                rn = input('Enter 15 Digit Reference No. ')
                if len(rn) == 15:
                    R.append(rn)
                    check_R_m = 1
                else:
                    continue        
        if entry_mode.upper() == "C":            
            check_R_c = 0
            while check_R_c == 0:
                csv_file_name = input("Enter CSV file name: ")
                try:
                    with open(f"{csv_file_name}.csv") as f:
                        contents_of_f = csv.reader(f)
                        for each_line in contents_of_f:
                            R += each_line
                    if len(R) >= 1:
                        check_R_c = 1
                    else:
                        print("Entered file contains no entry, Kindly Enter valid file name.")
                except:
                    check_R_c = 0            
        e_m_mode = 1
    else:
        print("Kindly enter 'M' or 'C'")
R_ok = list()
R_error = list()
for elem in R:
    if len(elem) == 15:
        R_ok.append(elem)
    else:
        R_error.append(elem)
executioncount = 0
filecount =0
reference_number_causing_errors = list()
for elem_R in R_ok:
    executioncount += 1
### Process Reference No. ###
    reference_no = elem_R
    BatchNo = reference_no[0:2]
    SubDiv = reference_no[2:7]
    RefNo = reference_no[7:14]
    RU = reference_no[14]
    RU = RU.upper()
    try:
        ### Read URL data ###
        print(f"PDF created for : {BatchNo}-{SubDiv}-{RefNo}")
        url = f"http://lesco.gov.pk/Modules/CustomerBill/BillPrintMDI.asp?nBatchNo={BatchNo}&nSubDiv={SubDiv}&nRefNo={RefNo}&strRU={RU}"
        pdfkit.from_url(url, f"{BatchNo}-{SubDiv}-{RefNo}.pdf", configuration=config,options=options)
        filecount += 1
    
    except:
        reference_number_causing_errors.append(reference_no)


print(" ")
print(f"Total reference nos. entered or in file : {len(R)}")
print(f"Total Number of records processed : {executioncount}")      
print(f"Number of files created : {filecount}") 
if len(R_error) > 0:
    n = 1
    print(" ")
    print("Following Reference Nos. in the CSV are invalid kindly re-check them manually")
    for elem in R_error:
        print (f"{n} --- {elem}")
        n += 1
if len(reference_number_causing_errors) > 0:
    n = 1
    print(" ")
    print("Following Reference Nos. are generating no result kindly re-check them manually")
    for elem in reference_number_causing_errors:
        print (f"{n} --- {elem}")
        n += 1

