import os
import toml
import re
import sys
import webbrowser

from pathlib import Path
from datetime import datetime, date
from string import Formatter
from urllib.parse import quote
from typing import Dict, Sequence, Mapping, Any


from masters import app
from masters import get_master_file
from masters.master import XLMaster


def is_using_outlook() -> bool:
    """
    Test if Outlook is the default mail client for darwin OS.
    :return:
    """
    from pathlib import Path
    import plistlib

    if path := Path(Path.home(), "Library/Preferences/com.apple.LaunchServices/com.apple.launchservices.secure.plist"):
        with open(path, "rb") as fp:
            for handler in plistlib.load(fp)["LSHandlers"]:
                if handler.get("LSHandlerURLScheme") == "mailto":
                    return 'outlook' in handler["LSHandlerRoleAll"]
    return False


class ScheduleRequest:
    """
    Create email for a specific agency regarding availability of its stations
    """

    def __init__(self, agency: str, force_text: bool = False):
        """
        Initialize class
        :param agency: Agency to contact
        :param force_text: Email in text format
        """
        # Set options to master schedule
        year = app.args.year
        # Initialize class with data from agency
        self.agency, self.antennas = agency, agency['antennas']
        names = list(self.antennas.values())
        antenna_names = names[0] if len(names) == 1 else f"{', '.join(names[:-1])} and {names[-1]}"
        plural = 's' if len(names) > 1 else ''

        self.to, self.cc = self.agency.get('to', []), self.agency.get('cc', [])

        request = app.config['request']
        self.subject = request['subject'].format(year=year, antennas=antenna_names)
        self.text = request['text'].format(greeting=self.agency['greeting'], antennas=antenna_names, plural=plural)

        # Check if html could be use
        if os.name != 'posix' and not force_text:
            self.fmt = HTMLFormatter(request['header'], request['format'])
            self.show = self.show_win_outlook
        elif sys.platform == 'darwin' and not force_text and is_using_outlook():
            self.fmt = HTMLFormatter(request['header'], request['format'])
            self.show = self.show_mailto
        else:
            self.fmt = TEXTFormatter(request['header'], request['format'])
            self.show = self.show_mailto

    def build(self, master: XLMaster) -> None:
        """
        Build message with the list of sessions for agency stations
        :param master: master information
        :return: None
        """
        self.fmt.body_begin()
        self.fmt.body_text(self.text)

        for sta_id, name in self.antennas.items():
            self.fmt.antenna_begin(name)
            rec = 0
            for ses in master.sessions:
                if sta_id in ses['master']:
                    rec += 1
                    ses['rec'] = rec
                    ses['STATIONS'] = ses['master']
                    self.fmt.session(ses)
            self.fmt.antenna_end()
        self.fmt.body_end()

    def show_win_outlook(self) -> None:
        """
        Create html email for outlook users on Window and open outlook
        :return: None
        """
        from win32com.client import Dispatch

        outlook = Dispatch("Outlook.Application")

        mail = outlook.CreateItem(0)

        separator = ',' if os.name == 'posix' else ';'
        mail.to = separator.join(self.to).strip()
        mail.cc = separator.join(self.cc).strip()
        mail.subject = self.subject
        mail.HTMLBody = self.fmt.build_text()
        mail.Display(False)

    def show_mailto(self) -> None:
        """
        Create text email and open mail editor
        :return: None
        """
        subject = quote(self.subject)
        body = quote(self.fmt.build_text())

        # Create mailto message
        separator = ',' if os.name == 'posix' else ';'
        to = quote(separator.join(self.to).strip(), '@,')
        cc = quote(f'cc={separator.join(self.cc).strip()}&', '@,') if self.cc else ''

        webbrowser.open(f"mailto:{to}?{cc}subject={subject}&body={body}")


# Write body in text mode.
class TEXTFormatter(Formatter):
    """
    Text formatter
    """
    def __init__(self, header: str, format_dict: Dict[str, str]):
        """
        Initailze formatter
        :param header: Header at top of sessions list
        :param format_dict: Dictionary of formats for fields
        """
        self.header, self.fmt = header, self._build_format(format_dict)
        self.lines, self.key = [], None

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        pass

    # Build format for this formatter
    @staticmethod
    def _build_format(format_dict):
        full_fmt = ['{rec:3d} ']
        for code, fmt in format_dict.items():
            if code == 'Stat1':
                code = 'STATIONS'
            elif code.startswith('Stat'):
                continue
            if fmt := fmt.replace('{', '{{{}:'.format(code)):
                full_fmt.append(fmt)
        return ''.join(full_fmt)

    def body_begin(self) -> None:
        """
        Begin body
        :return: None
        """
        pass

    def body_end(self) -> None:
        """
        End of body
        :return: None
        """
        pass

    def body_text(self, text: str) -> None:
        """
        Add text to body
        :param text: Text to be included in body
        :return: None
        """
        self.lines.extend(text.splitlines())

    # Start antenna data
    def antenna_begin(self, name):
        """
         Include antenna data in body
         :param name: Antenna name
         :return:
         """
        self.lines.extend([f'{name}', '\n', self.header])

    def antenna_end(self) -> None:
        """
        Close the antenna information
        :return:
        """
        self.lines.append('')

    # Write line
    def session(self, ses) -> None:
        """
        Write information for session
        :param ses: Session information
        :return: None
        """
        self.lines.append(self.format(self.fmt, **ses))

    def build_text(self) -> str:
        """
        Build the text from list of lines
        :return:
        """
        return '\n'.join(self.lines)

class HTMLFormatter(TEXTFormatter):
    """
    HTML formatter
    """

    # Build format for this formatter
    @staticmethod
    def _build_format(format_dict: Dict[str, str]) -> str:
        """
        Build the format string for each line
        :param format_dict: Dictionary of formats
        :return: The format for each session
        """
        full_fmt = ['<td>{rec:d}</td>']
        for code, fmt in format_dict.items():
            if code == 'Stat1':
                code = 'STATIONS'
            elif code.startswith('Stat'):
                continue
            if fmt := fmt.replace('{', '{{{}:'.format(code)):
                full_fmt.append(f'<td>{fmt}</td>')
        full_fmt.append('</tr>')
        return ''.join(full_fmt)

    def body_begin(self) -> None:
        """
        Begin body
        :return: None
        """
        self.lines.append('<html><body>')

    def body_end(self) -> None:
        """
        End of body
        :return: None
        """
        self.lines.append('</body></html>')

    # Add text
    def body_text(self, text: str) -> None:
        """
        Add text to body
        :param text: Text to be included in body
        :return: None
        """
        # Replace end of line by <br>
        self.lines.append('<br>'.join(text.splitlines()+['']))

    def antenna_begin(self, name: str) -> None:
        """
        Include antenna data in body
        :param name: Antenna name
        :return:
        """
        self.lines.append('<h3>{}</h3>'.format(name))
        self.lines.append('<table style=\"padding-right: 10px;\">')
        line = '<thead><tr style=\"font-weight:bold\"><td style=\"text-align:right\"></td>'
        for word in self.header.split():
            if word == 'DUR':
                line += '</td><td style=\"text-align:center\"></td>'
            else:
                line += '<td>{}</td>'.format(word.capitalize())
        line += '</tr></thead><tbody>'
        self.lines.append(line)

    def antenna_end(self) -> None:
        """
        Close the antenna information
        :return:
        """
        self.lines.append('</tbody></table><br>')

    def session(self, ses: Dict[str, Any]) -> None:
        """
        Write information for session
        :param ses: Session data
        :return: None
        """
        self.lines.append(self.format(self.fmt, **ses))

    def get_field(self, field_name: str, args: Sequence, kwargs: Mapping[str, Any]) -> Any:
        """
        Get the field name for the next format_field request
        :param field_name:
        :param args:
        :param kwargs:
        :return:
        """
        self.key = field_name
        return Formatter.get_field(self, field_name, args, kwargs)

    def format_field(self, value: Any, format_spec: Dict[str, str]) -> str:
        """
        Format a value using the self.key
        :param value: value to format
        :param format_spec: dictionary of formats
        :return: Formatted string
        """
        # do special formatting for some fields
        if self.key in ['DATE', 'TIME']:
            value = value.strftime(format_spec).upper()
            format_spec = ''
        elif self.key == 'DUR' and format_spec == '%H:%M':
            hours, remainder = divmod(value, 1)
            value = f'{int(hours):2d}:{int(remainder*60):02d}'
            format_spec = ''
        elif self.key == 'STATUS':
            if isinstance(value, datetime):
                value = value.strftime(format_spec).upper()
                format_spec = ''
            else:
                format_spec = '<{:d}s'.format(len(date.today().strftime(format_spec).upper()))
        elif self.key in ['PF', 'MK4NUM'] and not isinstance(value, (int, float)):
            format_spec = '<{:d}s'.format(int(re.sub(r"\D", ' ', format_spec).strip().split()[0]))
        value = ' ' if value is None else value
        return format(value, format_spec).strip()

    def get_value(self, key: str or int, args: Sequence, kwds: Mapping[str, Any]) -> Any:
        """
        Retrive given field value
        :param key: Index of args or key of kwds
        :param args: list of positional arguments to vformat
        :param kwds: dictionary of keyword arguments
        :return: Any
        """
        self.key = key
        return kwds[key] if isinstance(key, str) else Formatter.get_value(key, args, kwds)

    def build_text(self) -> str:
        """
        Build the text from list of lines
        :return:
        """
        return ''.join(self.lines)



def main():

    import argparse

    parser = argparse.ArgumentParser(description='Generate master file')
    parser.add_argument('-c', '--config', help='config file', required=True)
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument('-master', action='store_true')
    group.add_argument('-intensives', action='store_true')
    parser.add_argument('-text', action='store_true')
    parser.add_argument('year', help='master file year', type=int)
    parser.add_argument('agency', help='agency code')

    app.init(parser.parse_args())

    # Test if agency in list of agencies
    if not (agency := toml.load(Path(app.config['folder'], app.config['agencies'])).get(app.args.agency, None)):
        print(f'{app.args.agency} not a valid agency')
        return

    with XLMaster(get_master_file()) as master:
        if master.process():
            request = ScheduleRequest(agency, force_text=app.args.text)
            request.build(master)
            request.show()


if __name__ == '__main__':

    sys.exit(main())
