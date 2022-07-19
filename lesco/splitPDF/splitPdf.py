from PyPDF2 import PdfFileWriter, PdfFileReader
fileName = input("Enter file name to split : ")
fileName1 = fileName + ".pdf"
input_pdf = PdfFileReader(fileName1)
for i in range(input_pdf.getNumPages()):
    output = PdfFileWriter()
    output.addPage(input_pdf.getPage(i))
    with open(f"{fileName}_{i+1}.pdf", "wb") as output_stream:
        output.write(output_stream)
