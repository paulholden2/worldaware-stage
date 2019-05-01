import sys
import os
import csv
import magic
import mimetypes
import openxmllib
import comtypes.client
import extract_msg
import distutils.dir_util
from shutil import copy, copyfile

wdFormatPdf = 17

class Stager:
    def __init__(self, dest_dir, source_dir):
        self.dest_dir = dest_dir
        self.source_dir = source_dir
        self.word = comtypes.client.CreateObject('Word.Application')

    def is_already_staged(self, source_path, file_type):
        if not source_path.startswith(self.source_dir):
            raise Exception('File ' + source_path + ' not under source_dir')
        rel_path = os.path.relpath(source_path, self.source_dir)
        dir_name = os.path.dirname(rel_path)
        base_name = os.path.basename(source_path)
        file_name, ext_name = os.path.splitext(base_name)
        if file_type == 'msg':
            msg = extract_msg.Message(source_path)
            for att in msg.attachments:
                attachment_name = file_name + '-' + att.longFilename
                if not self.is_already_staged(os.path.join(self.dest_dir, dir_name, attachment_name), 'pdf'):
                    return False
            return True
        else:
            search_path = os.path.join(self.dest_dir, dir_name, file_name + '.pdf')
            return os.path.isfile(search_path)

    def stage_file(self, source_path, dest_dir, file_type):
        file_name = self.source_file_name(source_path)
        if file_type == 'docx' or file_type == 'doc':
            doc = self.word.Documents.Open(source_path)
            file_path = self.source_to_dest(source_path)
            distutils.dir_util.mkpath(file_path)
            doc.SaveAs(os.path.join(file_path, file_name + '.pdf'), FileFormat=wdFormatPdf)
            doc.Close()
        elif file_type == 'msg':
            self.extract_attachments(source_path, dest_dir)
        elif file_type == 'pdf':
            if source_path.endswith('.pdf'):
                copy(source_path, dest_dir)
            else:
                file_name = self.source_file_name(source_path)
                copyfile(source_path, os.path.join(dest_dir, file_name + '.pdf'))
        else:
            print('Cannot stage: ' + source_path)

    def stage_file_if_missing(self, source_path, file_type):
        if not self.is_already_staged(source_path, file_type):
            self.stage_file(source_path, self.dest_dir, file_type)

    def guess_file_type(self, file_path):
        res = magic.from_file(file_path, mime=True)
        ext = mimetypes.guess_extension(res, strict=False)
        if ext is None:
            return None
        return ext[1:]

    def stage_files_from_csv(self, csv_path):
        with open(csv_path, 'r') as csv_file:
            reader = csv.DictReader(csv_file)
            for row in reader:
                file_path = row['Full UNC Path']
                file_ext = row['FileExt'].lower()
                if file_ext == '':
                    file_ext = self.guess_file_type(file_path)
                if file_ext is None:
                    print('Unable to determine file type of' + file_path)
                    continue
                print('staging as ' + file_ext)
                self.stage_file_if_missing(file_path, file_ext)

    def extract_attachments(self, source_path, dest_dir):
        msg = extract_msg.Message(source_path)
        file_name = self.source_file_name(source_path)
        for att in msg.attachments:
            # Prepend attachment name with the email file's name (to avoid
            # potential name conflicts)
            attachment_name = file_name + '-' + att.longFilename
            attachment_path = self.source_to_dest(source_path)
            # Create directories as necessary
            distutils.dir_util.mkpath(attachment_path)
            att.save(customFilename=attachment_name, customPath=attachment_path)

    # Get file name (no ext)
    def source_file_name(self, source_path):
        base_name = os.path.basename(source_path)
        file_name, ext_name = os.path.splitext(base_name)
        return file_name

    # Get destination path (relative to Stager's source dir) within the
    # Stager's dest dir.
    def source_to_dest(self, source_path):
        rel_path = os.path.relpath(source_path, self.source_dir)
        dir_name = os.path.dirname(rel_path)
        return os.path.join(self.dest_dir, dir_name)

    def cleanup(self):
        self.word.Quit();

##

job_folder = '\\\\stria-prod1\\CID01570 - WorldAware\\JID01215 - CaaS'
dest_dir = job_folder + '\\Staging'
source_dir = job_folder + '\\SharePoint Files'

stager = Stager(dest_dir, source_dir)
stager.stage_files_from_csv('./File List Short.csv')
stager.cleanup()
