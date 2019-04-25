import sys
import os
import comtypes.client

wdFormatPdf = 17

in_file = os.path.abspath('./Lorem ipsum.docx')
out_file = os.path.abspath('./Lorem ipsum.pdf')

word = comtypes.client.CreateObject('Word.Application')

doc = word.Documents.Open(in_file)
doc.SaveAs(out_file, FileFormat=wdFormatPdf)
doc.Close();

word.Quit();
