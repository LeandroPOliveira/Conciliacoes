
import win32com.client
from fpdf import FPDF
from reportlab.pdfgen import canvas
from PyPDF2 import PdfFileWriter, PdfFileReader
import os
import glob

# Import Module
from win32com import client

# Open Microsoft Excel
excel = client.Dispatch("Excel.Application")

# Read Excel File
sheets = excel.Workbooks.Open('C:\\Users\loliveira\PycharmProjects\Excel\Conta 112421.xlsx')
work_sheets = sheets.Worksheets[0]

# Convert into PDF File
work_sheets.ExportAsFixedFormat(0, 'C:\\Users\loliveira\PycharmProjects\Excel\Conta 112.pdf')

# Create the watermark from an image
c = canvas.Canvas('watermark.pdf')
# Draw the image at x, y. I positioned the x,y to be where i like here
c.drawImage('Paulo.png', 440, 30, 100, 60,
            mask='auto')
c.save()
# Get the watermark file you just created
watermark = PdfFileReader(open("watermark.pdf", "rb"))
# Get our files ready


output = PdfFileWriter()

input = PdfFileReader(open("Conta 112.pdf", "rb"))
number_of_pages = input.getNumPages()

for current_page_number in range(number_of_pages):
    page = input.getPage(current_page_number)
    if page.extractText() != "":
        output.addPage(page)


page_count = output.getNumPages()
# Go through all the input file pages to add a watermark to them
for page_number in range(page_count):
    input_page = output.getPage(page_number)
    if page_number == page_count - 1:
        input_page.mergePage(watermark.getPage(0))
    output2 = PdfFileWriter()
    output2.addPage(input_page)

    # # dir = os.getcwd()
    # path = 'G:\GECOT\Análise Contábil_Tributária_Licitações\\2021'
    # os.chdir(path)
    # file = glob.glob(str(arquivo[21:32]) + '*')
    # file = ''.join(file)
    # try:
    #     os.chdir(file)
    # except:
    #     os.chdir(path)

    # finally, write "output" to document-output.pdf
    with open('Conta_nova112.pdf', "wb") as outputStream:
        output2.write(outputStream)

