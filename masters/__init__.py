import os
import sys
import webbrowser

from pathlib import Path, PurePosixPath
from tempfile import gettempdir
from datetime import datetime
from typing import Optional, Dict, List, Any

import tkinter as tk
import tkinter.simpledialog
import toml
import json

from masters import app
from masters.client import Client

# Password for remote server
passwords = {}


class Base:
    """
    Base class handling errors and messages
    """
    def __init__(self):
        self.debug, self.show = app.config.get('debug', False), app.config.get('show', False)
        self.has_errors, self.messages = None, []

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.has_errors or self.show:
            self.show_status()

    # exit with error message
    def exit(self, error: Optional[str] = None) -> None:
        if error:
            self.has_errors = True
            self.messages.append({'type': 'ERROR', 'text': error})
        if self.messages:
            print(f'Terminate with {"error" if self.has_errors else "comments"}.')
            print('\n'.join([f'{msg["type"]} : {msg["text"]}' for msg in self.messages]))
        sys.exit(0)

    def add_error(self, ses: Dict[str, Any], error: str, debug: bool = False) -> None:
        """
        Add code and row information to message and append to appropriate group
        :param ses: session information stored in dict
        :param error: error string
        :param debug: true if debug mode
        :return: None
        """
        if not debug:
            self.has_errors = True
        extra = 'debug' if self.debug else ''
        msg = {'type': 'ERROR', 'text': f'{ses["CODE"]} ({ses["row"]:d}) {error} {extra}'}
        self.messages.append(msg)

    #
    def add_information(self, info: Optional[List[str]|List[str]]) -> None:
        """
        Add information to messages
        :param info: List of string messages or list of lists of string
        :return: None
        """
        for lines in info:
            for line in ([lines] if isinstance(lines, str) else lines):
                msg = {'type': 'INFO', 'text': line.replace('\n', '')}
                self.messages.append(msg)

    def open_file(self, path: Path) -> None:
        """
        Open file for specific platform
        :param path: file path
        :return: None
        """
        os.system(f'open -e {path}') if os.name == 'posix' else webbrowser.open(path)

    def show_status(self) -> None:
        """
        Open text file with messages if any
        :return:
        """
        if self.messages:
            path = Path(gettempdir(), f'msg-{datetime.now().strftime("%Y%m%d-%H%M%S")}.txt')
            with open(path, 'w+') as tmp:
                #tmp.print('')
                print('\n'.join([f'{msg["type"]} : {msg["text"]}' for msg in self.messages]), file=tmp)
            self.open_file(path)

    def upload_files(self, files: List[Path], file_type: str, listing: bool = False):
        """
        Upload list of files to remote server
        :param files: List of files
        :param file_type: 'master' or 'backup'
        :param listing: show list of remote directory
        :return: None
        """
        servers = toml.load(Path(app.config['folder'], app.config.get('servers', 'servers.toml')))
        if not (scp := app.config['scp'].get(file_type)):
            return  # Do not want to copy the files for this specific file type

        # Copy to remote server
        server_name, remote_folder = scp['server'], scp['folder']
        commands, setmode = scp.get('commands', []), scp.get('setmode', False)
        if listing and 'ls -l' not in commands:
            commands.append('ls -l')

        host = servers[server_name]
        if not host.get('id_rsa', '').strip():
            host['password'] = get_password(server_name, host['user'])

        with Client(host) as server:
            if not server.connected:
                self.exit(error=f'Could not connect to {server_name} {server.error}')

            for path in files:
                remote_path = str(PurePosixPath(remote_folder).joinpath(path.name))
                ok, msg = server.put_and_exec(str(path), remote_path, commands, setmode)
                if not ok:
                    self.exit(error=f'Could not copy {path} to {server_name} {remote_path}- {msg}')
                self.add_information(msg)

    def exec_commands(self):
        """
        execute
        :param parent:
        :return:
        """
        if actions := app.config.get('exec', None):
            servers = toml.load(Path(app.config['folder'], app.config.get('servers', 'servers.toml')))
            # Execute commands in exec
            for action in actions:
                if (command := action.get('command', '')) and (server_name := action.get('server', '')):
                    if not (host := servers[server_name]).get('id_rsa', '').strip():
                        host['password'] = get_password(server_name, host['user'])
                    with Client(host) as server:
                        if not server.connected:
                            self.exit(error=f'Could not connect {action["server"]} {server.error}')
                        self.add_information(server.exec(command))


def get_file_name(code: str, extension: str, year:str):
    """
    Get file path from config file
    :param code: code for type sessions [master, intensive, vgos]
    :param extension: file extension, [xlsx, txt, docx]
    :param year:
    :return:
    """
    return app.config[code]['filename'][extension].format(year=year, yy=year[2:])


def get_master_file(extension: str = 'xlsx') -> str:
    """
    Make path from application options
    :param extension:
    :return:
    """
    for code in ['master', 'intensives']:
        if getattr(app.args, code):
            return Path(app.config['folder'], get_file_name(code, extension, str(app.args.year)))


def get_password(server: str, user: str):
    """
    Get password for specific user
    :param server: server name
    :param user: username
    :return:
    """
    global passwords

    if not passwords.get(server):
        tk.Tk().withdraw()
        passwords[server] = tkinter.simpledialog.askstring(f"{user} password for {server}", "Enter password:", show='*')

    return passwords[server]
