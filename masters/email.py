from datetime import datetime
from urllib.parse import quote
import webbrowser
import os
from typing import List, Tuple

from masters import app
from masters.notes import Notes


class Email:
    """
    Class used to generate email and show the email
    """

    def __init__(self, code: str = 'master'):
        """
        Initialize class
        :param code: code for type of schedule (master or intensive)
        """
        info = app.config['email']
        self.to, self.cc = info[code]['to'], info[code].get('cc', [])
        self.body, self.subject = info['body'], info['subject']
        self.label = ' ' if code == 'master' else f' {code} '

    def split_comments(self, text: str) -> List[str]:
        """
        Split comment to keep each line in paragraph smaller than max_length
        :param text: Line of text
        :return: List of lines with maximum length
        """
        if len(text) <= Notes.max_length:
            return [text]
        index = text[:Notes.max_length].rfind(' ')
        return [text[:index].strip()] + self.split_comments(text[index:].strip())

    def make_notes(self, notes: Notes) -> Tuple[str, List[str]]:
        """
        Build the email messages using notes in docx format
        :param notes: Notes read from docx document
        :return: None
        """
        today = datetime.now().strftime('%B %d')
        accepted = []
        # Keep blocks with last date only
        for block in notes.blocks:
            if block['date'] == today or (accepted and not block['date']):
                accepted.append(block)
            else:
                accepted = []

        # Write notes
        email_notes = []
        for block in accepted:
            if block['date']:
                email_notes.append(block['date'])
            comments = notes.build_text_paragraph(block['text'])
            for comment in comments:
                email_notes.append(f"{' '*15}{comment}")
            email_notes.append('')
        return today, email_notes

    def mailto(self, year: str, notes: Notes) -> None:
        """
        Use the notes to generate email and start email client
        :param year: Year of master file
        :param notes: notes to be send by email
        :return: None
        """
        today, notes = self.make_notes(notes)
        updated = datetime.now().strftime('%B %d, %Y')
        body = self.body.format(updated=updated, year=year, label=self.label, date=today, notes='\n'.join(notes))
        subject = self.subject.format(year=year, label=self.label)

        self.poxis_mail(subject, body) if os.name == 'posix' else self.outlook_mail(subject, body)

    def poxis_mail(self, subject: str, body:str) -> None:
        """
        Display email for mac and linux system (poxis)
        :param subject: subject
        :param body: body
        :return: None
        """

        # Create mailto message
        to = quote(','.join([name.replace(',', '') for name in self.to]).strip(),'@,')
        cc = f"cc={quote(','.join([name.replace(',', '') for name in self.cc]).strip(),'@,')}&" if self.cc else ''
        webbrowser.open(f"mailto:{to}?{cc}subject={quote(subject, '')}&body={quote(body, '')}")

    def outlook_mail(self, subject: str, body:str) -> None:
        """
         Display email for mac and linux system (poxis)
         :param subject: subject
         :param body: body
         :return: None
         """
        from win32com.client import Dispatch

        # Create mailto message
        mail = Dispatch("Outlook.Application").CreateItem(0)
        mail.to = ';'.join([name.replace(';', ' ') for name in self.to]).strip()
        mail.cc = ';'.join([name.replace(';', ' ') for name in self.cc]).strip() if self.cc else ''
        mail.subject = subject
        mail.Body = body
        mail.Display()

    
