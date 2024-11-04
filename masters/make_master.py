from tempfile import gettempdir
from pathlib import Path

from masters import app, get_file_name, get_master_file
from masters.master import XLMaster
from masters.notes import Notes
from masters.email import Email


def main():
    """
    Application to generate master and note text files, generate email and transfer files to remote server.
    :return: None
    """
    import argparse

    parser = argparse.ArgumentParser(description='Generate master file')
    parser.add_argument('-c', '--config', help='config file', required=True)
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument('-master', action='store_true')
    group.add_argument('-intensives', action='store_true')
    parser.add_argument('-t', '--text_only', action='store_true')
    parser.add_argument('year', help='master file year', type=int)

    app.init(parser.parse_args())

    with XLMaster(get_master_file()) as master:
        if master.process():
            # This option (-t) added to create a txt file without notes.
            if app.args.text_only:
                master.open_file(master.make_master())
                return

            # Make note text file from docx file
            if not (docx := Path(master.folder, get_file_name(master.type, 'docx', master.year))).exists():
                master.exit(error=f'Could not find {docx}')
            path = Path(gettempdir(), get_file_name(master.type, 'notes', master.year))
            notes = Notes(docx)
            notes.save_txt(path)
            # Files to be transferred for backup
            backup = [master.path, master.format_path, master.ns_codes_path, master.media_key_path, docx]
            # Files to be transfer to cddis
            files = list(filter(None, [master.make_master(), master.make_media(), path]))
            # Send files to external server
            master.upload_files(files, 'master')
            master.upload_files(backup, 'backup')
            master.exec_commands()
            # Prepare email
            email = Email(master.type)
            email.mailto(master.year, notes)


if __name__ == '__main__':

    import sys
    sys.exit(main())
