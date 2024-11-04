import re
import sys
from pathlib import Path
from typing import Union

from docx import Document
from docx.table import Table
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.text.paragraph import Paragraph
from docx.document import Document as _Document


class Notes:
    """
    Create master note text file from word docx file.
    """
    draft = 'DRAFT!!       DRAFT!!     DRAFT!!      NOT FINAL!!!     DRAFT!!         DRAFT!!'
    max_length = 80

    def __init__(self, path: Path):
        self.path = path
        self.max_line_length = Notes.max_length
        self.title, self.updated, self.author = '', '', 'N/A'
        self.blocks = []

        if not self.path.exists():
            print(f'{path} does not exist!')
            sys.exit(0)
        self.read_docx()

    # Read a docx document
    def read_docx(self) -> None:
        """
        Read docx document
        :return: None
        """
        try:
            doc = Document(self.path)
        except Exception as err:
            print(f'{self.path.name} is not a docx document! {str(err)}')
            sys.exit(0)
        # Read all paragraphs
        self.blocks = []
        for block in self.iter_blocks(doc):
            if isinstance(block, Paragraph):
                text = bytes(block.text, 'utf-8').decode('utf-8', 'ignore')
                if 'SCHEDULE NOTES' in text:
                    self.title = text
                elif 'Last Updated' in text:
                    self.updated = text
            elif isinstance(block, Table):
                for row in block.rows[1:]:
                    b = {'date': row.cells[0].text, 'text': []}
                    for line in row.cells[1].text.splitlines():
                        line = bytes(line, 'utf-8').decode('utf-8', 'ignore')
                        b['text'].append(line)
                    self.blocks.append(b)

        self.author = doc.core_properties.author
        for comment in doc.core_properties.comments.splitlines():
            if 'MaxLineLength' in comment:
                self.max_line_length = int(comment.split(':')[1].strip())

    def save_txt(self, path: Path) -> None:
        """
        Save new text document
        :param path: Path of note file
        :return: None
        """
        self.max_line_length -= 17
        # Open file for writing
        with open(path, 'w') as file:
            if self.title:  # Write title
                file.write(f'\n{self.title:^101s}\n\n')
            if self.updated:  # Write last updated
                file.write(f'{self.updated:^101s}\n\n')
            # Write each comments
            for block in self.blocks:
                date = block['date']
                comments = self.build_text_paragraph(block['text'])
                if len(comments) > 1 and len(comments[-1]) < 10:
                    comments[-2] += ' ' + comments[-1]
                    comments.pop(-1)
                for comment in comments:
                    if 'DRAFT!!' in comment:
                        comment = Notes.draft
                    file.write(f'{date:<17s}{comment}\n')
                    date = ' '
                file.write('\n')

    # Build paragraph so it looks ok in text file
    def build_text_paragraph(self, lines: list[str]) -> list[str]:
        """
        Build paragraph so it looks ok in text file
        :param lines: List of lines from word document
        :return: List of lines formatted for text file
        """
        comment, comments = [], []
        for line in lines:
            ok, line = self.same_paragraph(line)
            if ok:
                comment.append(line)
            else:
                if comment:
                    comments += self.split_comments(' '.join(comment))
                comments.append(line)
                comment = []
        if comment:
            comments += self.split_comments(' '.join(comment))
        return comments

    def split_comments(self, text: str) -> list[str]:
        """
        Split comment to keep each line in paragraph smaller than max_length
        :param text:
        :return: List of line with maximum length
        """
        if len(text) <= self.max_line_length:
            return [text]
        index = text[:self.max_line_length].rfind(' ')
        return [text[:index].strip()] + self.split_comments(text[index:].strip())

    def same_paragraph(self, line: str) -> tuple[bool, str]:
        """
        Try to find if a comment is the continuity of the previous comment.
        :param line: Line (string) to test
        :return: Tupple Bool (same paragraph) and line
        """
        if re.match(r'^[A-Z]\s-\s', line):
            return False, line
        if re.match(r'^[1-9]\.\s', line):
            return False, line

        line = self.clean_punctuation(line)
        text = ' '.join(line.split())
        if text == line:
            return True, line
        else:
            return False, line

    def clean_punctuation(self, line: str) -> str:
        """
        Clean end of line and remove double spaces
        :param line: Line to clean
        :return: Cleaned line
        """
        text = line.replace('.  ', '. ').replace(',  ', ', ')
        return text if text == line else self.clean_punctuation(text)

    def iter_blocks(self, parent: Document) -> Union[Paragraph, Table]:
        """
        Generate a reference to each paragraph and table child within *parent*,
        in document order. Each returned value is an instance of either Table or
        Paragraph. *parent* would most commonly be a reference to a main
        Document object, but also works for a _Cell object, which itself can
        contain paragraphs and tables.
        """
        if not isinstance(parent, _Document):
            raise ValueError(f"something's not right in {self.path.name}")

        for child in parent.element.body.iterchildren():
            if isinstance(child, CT_P):
                yield Paragraph(child, parent)
            elif isinstance(child, CT_Tbl):
                yield Table(child, parent)


def main():
    import argparse
    from tempfile import gettempdir

    from masters import app, get_master_file, get_file_name

    parser = argparse.ArgumentParser(description='Generate note text file')
    parser.add_argument('-c', '--config', help='config file', required=True)
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument('-master', action='store_true')
    group.add_argument('-intensives', action='store_true')
    group.add_argument('-vgos', action='store_true')
    parser.add_argument('year', help='year')
    args = app.init(parser.parse_args())
    print(type(args))

    path = Path(gettempdir(), get_master_file('notes').name)
    notes = Notes(get_master_file('docx'))
    notes.save_txt(path)
    print(f'Notes saved in {path}')


if __name__ == '__main__':

    sys.exit(main())
