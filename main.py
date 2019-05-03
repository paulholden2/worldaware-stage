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
        self.file_list = []
        self.problem_files = []
        self.file_count = 0
        self.files_processed = 0
        self.currently_staging = ''

    def is_already_staged(self, source_path, file_type):
        if not source_path.startswith(self.source_dir):
            raise Exception('File ' + source_path + ' not under source_dir')
        rel_path = os.path.relpath(source_path, self.source_dir)
        dir_name = os.path.dirname(rel_path)
        base_name = os.path.basename(source_path)
        file_name, ext_name = os.path.splitext(base_name)
        if file_type == 'msg':
            msg = extract_msg.Message(source_path)
            for i, att in enumerate(msg.attachments):
                att_id = att.longFilename
                if att_id is None:
                    att_id = i
                attachment_name = file_name + '-' +  str(att_id)
                if not self.is_already_staged(os.path.join(self.source_dir, dir_name, attachment_name), 'pdf'):
                    return False
            return True
        else:
            search_path = os.path.join(self.dest_dir, dir_name, file_name + '.pdf')
            return os.path.isfile(search_path)

    def stage_file(self, source_path, dest_dir, file_type):
        file_name = self.source_file_name(source_path)
        if file_type == 'docx' or file_type == 'doc':
            try:
                doc = self.word.Documents.Open(source_path)
                file_path = self.source_to_dest(source_path)
                distutils.dir_util.mkpath(file_path)
                doc.SaveAs(os.path.join(file_path, file_name + '.pdf'), FileFormat=wdFormatPdf)
                doc.Close()
            except comtypes.COMError as err:
                raise Exception('ERROR while staging: ' + source_path + ' - ' + err.args[2][0])
        elif file_type == 'msg':
            self.extract_attachments(source_path, dest_dir)
        else:
            ok_exts = [
                'bmp',
                'tif',
                'tiff',
                'jpg',
                'jpeg',
                'png',
                'gif',
                'pdf'
            ]

            try:
                idx = ok_exts.index(file_type)
                ext = '.' + ok_exts[idx]
                copy_dest = self.source_to_dest(source_path)
                distutils.dir_util.mkpath(copy_dest)
                if source_path.endswith(ext):
                    copy(source_path, copy_dest)
                else:
                    file_name = self.source_file_name(source_path)
                    copyfile(source_path, os.path.join(copy_dest, file_name + ext))
            except ValueError as e:
                self.add_problem_file(source_path, 'Invalid file type')

    def stage_file_if_missing(self, source_path, file_type):
        try:
            rel_path = os.path.relpath(source_path, self.source_dir)
            self.currently_staging = rel_path
            if not self.is_already_staged(source_path, file_type):
                self.stage_file(source_path, self.dest_dir, file_type)
        except Exception as exception:
            self.add_problem_file(source_path, str(exception))

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
                self.file_list.append(row.copy())

        self.file_count = len(self.file_list)

        print('%d files' % self.file_count)
        print('')
        print('')

        for i, row in enumerate(self.file_list):
            file_path = row['Full UNC Path']
            file_ext = row['FileExt'].lower()
            if file_ext == '':
                file_ext = self.guess_file_type(file_path)
            if file_ext is None:
                self.add_problem_file(file_path, 'Unknown file type')
            else:
                self.stage_file_if_missing(file_path, file_ext)
            self.files_processed += 1
            self.print_progress()

    def add_problem_file(self, source_path, reason):
        self.problem_files.append({ 'path': source_path, 'reason': reason })
        shortened_name = self.shorten_name(os.path.relpath(source_path, self.source_dir))
        sys.stdout.write('\r\033[F\rERR: %+70s\n\n' % shortened_name)

    def shorten_name(self, source_path):
        if source_path > 60:
            return source_path[:30] + '...' + source_path[-30:]
        else:
            return source_path

    def print_progress(self):
        bar_len = 50
        count = self.file_count
        progress = self.files_processed
        filled = int(round(bar_len * (progress / float(count))))
        percent = 100 * progress / float(count)
        bar = '=' * filled + '-' * (bar_len - filled)
        staging_name = self.shorten_name(self.currently_staging)
        error_count = len(self.problem_files)
        sys.stdout.write('\r\033[F\r%-80s\n[%s] % 4.1f%% %5d/%-5d s/e' % (staging_name, bar, percent, progress, error_count))
        sys.stdout.flush()

    def extract_attachments(self, source_path, dest_dir):
        msg = extract_msg.Message(source_path)
        file_name = self.source_file_name(source_path)
        for i, att in enumerate(msg.attachments):
            att_id = att.longFilename
            if att_id is None:
                att_id = i
            # Prepend attachment name with the email file's name (to avoid
            # potential name conflicts)
            attachment_name = file_name + '-' +  str(att_id)
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
dest_dir = job_folder + '\\Staging\\Client'
source_dir = job_folder + '\\SharePoint Files\\Client'

stager = Stager(dest_dir, source_dir)
stager.stage_files_from_csv('./File List.csv')
stager.cleanup()
