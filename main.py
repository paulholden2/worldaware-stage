import sys
import os
import comtypes.client
import extract_msg

wdFormatPdf = 17

class Stager:
    def __init__(self):
        self.word = comtypes.client.CreateObject('Word.Application')

    def stage_file(self, source_path, dest_dir, filetype):
        doc = word.Documents.Open(source_path)
        doc.SaveAs(os.path.join(dest_dir, os.path.basename(source_path)), FileFormat=wdFormatPdf)
        doc.Close()

    def extract_attachments(self, source_path, dest_dir):
        msg = extract_msg.Message(source_path)
        for att in msg.attachments:
            att.save(useFileName=True, customPath=dest_dir)

    def cleanup():
        self.word.Quit();

##

stager = Stager()
stager.cleanup()
