#!/usr/bin/env python
# coding: utf-8

# In[ ]:


### Imports ###
import urllib.request, urllib.parse, urllib.error
from bs4 import BeautifulSoup
import ssl
import docx
from docx.shared import Inches, Cm
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import csv
from csv import writer

### Ignoring Errors ###
ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

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
def append_list_as_row(file_name, list_of_elem):
    # Open file in append mode
    with open(file_name, 'a+', newline='') as write_obj:
        # Create a writer object from csv module
        csv_writer = writer(write_obj)
        # Add contents of list as last row in the csv file
        csv_writer.writerow(list_of_elem)
executioncount = 0
filecount = 0
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
        url = f"http://lesco.gov.pk/Modules/CustomerBill/BillPrintMDI.asp?nBatchNo={BatchNo}&nSubDiv={SubDiv}&nRefNo={RefNo}&strRU={RU}"
        html = urllib.request.urlopen(url, context=ctx).read()
        soup = BeautifulSoup(html, 'html.parser')

        ### Reading Reference No., Sanctioned Load and Current Month ###
        tags = soup('div')
        
        n = 0
        m = 0
        for tag in tags:
            a = tag.contents[0]
            try:
                a = a.lstrip()
                a = a.rstrip()
                if len(a) <= 0:
                    continue
                else:
                    n += 1
                    if a == "OLD A/C No.":
                        b = n+1
                    elif n == b:
                        reference_no = a
                    elif n == b+1:
                        tarrif = a    
                    elif n == b+2:
                        sanc_load = a
            except:
                continue

        ### Reading Name, Address and Current MDI ###
        tags = soup('td')
        n = 0
        for tag in tags:
            a = tag.contents[0]
            try:
                a = a.lstrip()
                a = a.rstrip()

                if len(a) <= 0:
                    continue
                else:
                    # print(a)
                    n = n+1
                    if a == "NAME & ADDRESS":
                        b = n+1
                    if n == b:
                        name = a
                    if n == b+1:
                        address = a
                    if a == "SUB-DIVISION":
                        d = n+23
                    if a == "FEEDER":
                        f = n+1
                    if n == f:
                        feeder = a
                    if n == d:
                        MF = a
                    if a[0:7] == "CHARGED":
                        c = n+1
                    if n == c:
                        current_mdi = a
            except:
                continue
        CRN = reference_no[0:2]+'-'+reference_no[3:8]+'-'+reference_no[9:16]

        print("**********************************************************************************************")
        print("Kindly verify the data")
        print(f"Reference No. : {reference_no}")
        print(f"Name : {name}")
        print(f"Address : {address}")
        print(f"Sanctioned Load : {sanc_load}-KW")
        print(f"TARIFF : {tarrif}")
        print(f"MF : {MF}")
        print(f"Feeder : {feeder}")
        print("**********************************************************************************************")
        row_contents = [CRN,name,address,sanc_load,tarrif,MF,feeder]
        append_list_as_row(f'{BatchNo}-consumer-details.csv', row_contents)
    
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

