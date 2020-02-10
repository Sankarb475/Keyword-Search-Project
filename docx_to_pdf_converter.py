import sys
import os
import comtypes.client

wdFormatPDF = 17

in_file = r"C:\Users\sb512911\Desktop\All\Applications\VDR\VDR data\dum\rtf.doc"
out_file = os.path.abspath(r"C:\Users\sb512911\Desktop\All\Applications\VDR\VDR data\dum\rtf1.pdf")

word = comtypes.client.CreateObject('Word.Application')
doc = word.Documents.Open(in_file)
doc.SaveAs(out_file, FileFormat=wdFormatPDF)
doc.Close()
word.Quit()

