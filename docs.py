import sys, os, subprocess, win32api, win32print
from PyPDF2 import PdfFileReader, PdfFileWriter
import comtypes.client

def convert(in_file, out_file, timeout=None):
	wdFormatPDF = 17
	word = comtypes.client.CreateObject('Word.Application')
	doc = word.Documents.Open(in_file)
	doc.SaveAs(out_file, FileFormat=wdFormatPDF)
	doc.Close()
	word.Quit()
	print("convertion complete.")

def cut():
	infile = PdfFileReader("document.pdf", 'rb')
	output = PdfFileWriter()
	p = infile.getPage(infile.getNumPages()-1)
	output.addPage(p)
	with open("print.pdf", 'wb') as f:
		output.write(f)
	os.remove("document.pdf")

if __name__ == '__main__':
	in_file = os.path.realpath(__file__).replace(os.path.basename(__file__),"")+"document.docx"
	out_file = os.path.realpath(__file__).replace(os.path.basename(__file__),"")+"document.pdf"
	convert(in_file, out_file)

	bb = b''
	with open("print.pdf", "rb") as f:
		byte = f.read(1)
		while byte:
			bb += byte
			byte = f.read(1)
	os.remove("print.pdf")

	hPrinter = win32print.OpenPrinter(win32print.GetDefaultPrinter())
	try:
		hJob = win32print.StartDocPrinter (hPrinter, 1, ("test of raw data", None, "RAW"))
		try:
			win32print.StartPagePrinter(hPrinter)
			win32print.WritePrinter(hPrinter, bb)
			win32print.EndPagePrinter(hPrinter)
		finally:
			win32print.EndDocPrinter(hPrinter)
	finally:
		win32print.ClosePrinter(hPrinter)