import socket
import re

from typing import Dict, List, Tuple, Any
from stat import S_IMODE
from paramiko import SSHClient, SSHException, BadHostKeyException, AuthenticationException, RSAKey
from paramiko import AutoAddPolicy


class Client:
    """
    Class for ssh client
    """
    def __init__(self, server: Dict[str, Any]) -> None:
        """
        Initialize class
        :param server: dictionary of server information (see servers.toml)
        """
        self.server = server
        self.client = SSHClient()
        self.client.set_missing_host_key_policy(AutoAddPolicy())
        self.sftp = None

        self.uid = self.gid = self.error = None
        self.connected = False

    def __enter__(self):
        self.connect()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        try:
            self.sftp.close()
            self.client.close()
        except AttributeError:
            pass

    def connect(self) -> None:
        """
        Connect to remote server
        :return: None
        """
        id_rsa = self.server.get('id_rsa', '').strip()
        pw, key = (None, RSAKey.from_private_key_file(id_rsa)) if id_rsa else (self.server['password'], None)
        try:
            self.client.connect(self.server['host'], port=self.server['port'],
                                username=self.server['user'], password=pw, pkey=key)
            self.sftp = self.client.open_sftp()
            self.connected = True
        except (BadHostKeyException, AuthenticationException, SSHException, socket.error) as err:
            self.error = str(err)
            self.connected = False

    def exec(self, cmd: str) -> List[str]:
        """
        Execute a specific command on remote server
        :param cmd: string to be executed
        :return: List of output lines
        """
        lines = []
        try:
            std_in, std_out, std_err = self.client.exec_command(cmd)
            lines.extend(std_out.readlines())
            lines.extend(std_err.readlines())
        except SSHException as err:
            lines.append(f'{cmd} failed error: {str(err)}')
        return lines

    def getids(self) ->Tuple[int, int]:
        """
        Get uid and gid of user on remote server
        :return:
        """
        if not self.uid or not self.gid:
            for info in self.exec(f"id {self.server['user']}")[0].split():
                if info.startswith('uid='):
                    self.uid = int(re.sub(r"\D", '', info).strip())
                if info.startswith('groups='):
                    for group in info.split(','):
                        if self.server['group'] in group:
                            self.gid = int(re.sub(r"\D", '', group).strip())
                            break

        return self.uid, self.gid

    def chmod(self, remote_path: str, mode: str or int) -> Tuple[bool, str]:
        """
        Change mode of file on remote server
        :param remote_path: path to remote file
        :param mode: mode in int or str formats
        :return: Succes, error message
        """
        mode = int(mode, 8) if isinstance(mode, str) else mode
        uid, gid = self.getids()
        stats = self.sftp.stat(remote_path)
        # No stats so file does not exist
        if not stats:
            return False, 'file does not exist'
        # Change group
        if gid != stats.st_gid:
            self.sftp.chown(remote_path, int(uid), int(gid))
        if mode != S_IMODE(stats.st_mode):
            self.sftp.chmod(remote_path, mode)
        return True, ''

    # Copy local file to remove server
    def put(self, local_path: str, remote_path: str, setmode: bool = False) -> Tuple[bool, str]:
        """
        Copy local file to remove server
        :param local_path:
        :param remote_path:
        :param setmode:
        :return:
        """
        try:
            self.sftp.put(local_path, remote_path)
            if setmode:
                return self.chmod(remote_path, '664')
        except Exception as err:
            return False, str(err)
        return True, ''

    def put_and_exec(self, local_path: str, remote_path: str, cmds: List[str],
                     setmode: bool = False) -> Tuple[bool, List[str]]:
        """
        Copy file to remote server and execute command
        :param local_path: path of local file as string
        :param remote_path: path of remote file as string
        :param cmds: List of commands
        :param setmode: set mode for remote file
        :return: Succes, Error message
        """
        ok, err = self.put(local_path, remote_path, setmode)
        if not ok:
            return ok, err
        return True, [self.exec(f'{cmd} {remote_path}') for cmd in cmds]

    def get(self, remote_path: str, local_path: str) -> bool:
        """
        Get file from remote server
        :param remote_path: path of remote file as string
        :param local_path: path of local file as string
        :return: Success
        """
        try:
            self.sftp.get(remote_path, local_path)
        except Exception:
            return False

        return True

    def remove(self, files: List[str] or str) -> bool:
        """
        Remove files on remote server
        :param files: List of files to remove
        :return: Success
        """
        try:
            for file in ([files] if isinstance(files, str) else files):
                self.sftp.remove(file)
        except Exception:
            return False

        return True
