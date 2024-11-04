from pathlib import Path
from masters import app, Base


def main():
    """
    This application make backup of specific files on a remote server
    :return: None
    """
    import argparse

    parser = argparse.ArgumentParser(description='Generate master file')
    parser.add_argument('-c', '--config', help='config file', required=True)

    app.init(parser.parse_args())
    app.config['show'] = True

    folder = Path(app.config['folder'])
    files = [path for path in folder.iterdir()
             if path.is_file() and ((path.stem.startswith(('master', 'int'))
                                     and path.suffix in ['.xlsx', '.docx'])
                                    or path.stem.startswith(('media-key', 'ns-codes', 'master-format')))]

    with Base() as process:
        process.upload_files(files, 'backup', show=True)


if __name__ == '__main__':

    import sys
    sys.exit(main())
