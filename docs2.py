import sys, os, subprocess, win32api, win32print
from PyPDF2 import PdfFileReader, PdfFileWriter
import time
import convertapi

def convert(in_file, timeout=None):
	convertapi.api_secret = ' '
	upload_io = convertapi.UploadIO(open(in_file, 'rb'))
	saved_files = convertapi.convert('pdf', {'File': upload_io}).save_files(os.path.realpath(__file__).replace(os.path.basename(__file__),""))
	print("The PDF saved to %s" % saved_files)

def cut():
	infile = PdfFileReader("document.pdf", 'rb')
	output = PdfFileWriter()
	p = infile.getPage(infile.getNumPages()-1)
	output.addPage(p)
	with open("print.pdf", 'wb') as f:
		output.write(f)
	os.remove("document.pdf")
	print("Cutting complete.")
	
def printer():
	hPrinter = win32print.OpenPrinter(win32print.GetDefaultPrinter())
	os.startfile(os.path.realpath(__file__).replace(os.path.basename(__file__),"")+"print.pdf", "print")
	time.sleep(5)
	os.system("TASKKILL /F /IM AcroRD32.exe") 
	os.remove("print.pdf")
	print("Printing complete.")

if __name__ == '__main__':
	in_file = os.path.realpath(__file__).replace(os.path.basename(__file__),"")+"document.docx"
	convert(in_file)
	cut()
	printer()