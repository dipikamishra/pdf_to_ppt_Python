import os
import win32com.client
import subprocess
from PyPDF2 import PdfFileReader

from PyPDF2 import PdfReader


# Set the path to the PDF file you want to convert

pdf_path = 'EScore-REAlignment20230503_Dipika.pdf'

# Use PyPDF2 to read the PDF file
#pdf_reader = PdfFileReader(open(pdf_path, 'rb'))
pdf_reader = PdfReader(open(pdf_path, 'rb'))


# Get the number of pages in the PDF file
#num_pages = pdf_reader.getNumPages()
num_pages = len(pdf_reader.pages)

# Create a list of page numbers to convert
page_nums = ','.join([str(i) for i in range(num_pages)])

# Set the output file path
output_path = 'output.pptx'

# Use subprocess to call the libreoffice command line utility to convert the PDF file to PPT
#subprocess.call(['libreoffice', '--headless', '--convert-to', 'pptx', '--outdir', os.path.dirname(output_path), '--convert-images-to-jpg', '--page-range', page_nums, pdf_path])

def convert_pdf_to_ppt(pdf_path, output_path, page_nums):
    powerpoint = win32com.client.Dispatch("Powerpoint.Application")
    deck = powerpoint.Presentations.Add()
    deck.SaveAs(output_path, 24)
    deck.Close()

    powerpoint.Quit()

convert_pdf_to_ppt(pdf_path,output_path,page_nums)