from collections import Counter
from typing import List

from openpyxl import load_workbook
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextBoxHorizontal, LTChar

file_pdf = 'Abstract Book from the 5th World Psoriasis and Psoriatic Arthritis Conference 2018.pdf'
file_xml = 'Data Entry - 5th World Psoriasis & Psoriatic Arthritis Conference 2018 - Case format (2).xlsx'
START_PAGE = 100
END_PAGE = 104
SHEET_NAME = 'Sheet1'
START_ROW = 6


def update_xml(file: str, rec: list, sheet_name: str, start_row: int) -> None:
    """
    :param file: file xml for read and update with bunch of records
    :param rec: list of tuples aka records, we expect that every tuple consists 6 values
    :param sheet_name: name of sheet that has to be updated
    :param start_row: number of row from which we start to update cells
    """
    wb = load_workbook(file)
    ws = wb[sheet_name]
    for row in range(start_row, len(rec) + start_row):
        for col in range(1, 7):
            ws.cell(row=row, column=col).value = rec[row - 6][col - 1]

    wb.save(file)
    wb.close()


def parse_articles(file: str, start_page: int, end_page: int) -> List[tuple]:
    """
    :param file: pdf-file which we want to parse
    :param start_page: number of page where we start to parse
    :param end_page: number of page where we end to parse
    :return: list of tuples(ordered values)
    """
    # flag for breaking a loop if we reach END_PAGE
    flag = False

    # initial values
    session_name = ''
    topic_title = ''
    name = ''
    affiliation_name = ''
    persons_location = ''
    abstract = ''
    records = []

    for page_layout in extract_pages(file):
        if flag:
            break

        for element in page_layout:
            if isinstance(element, LTTextBoxHorizontal):

                for text_line in element:
                    # let's do some investigation about most frequent font and size of one pdf-line
                    font_and_size_list = [(character.fontname, character.size) for character in text_line if
                                          isinstance(character, LTChar)]
                    font_and_size = Counter(font_and_size_list).most_common(1)[0][0]

                    if font_and_size == ('OAOVWE+TimesNewRomanPS-BoldItalicMT', 9.5):
                        base_name = text_line.get_text()
                        if affiliation_name:
                            persons_location = affiliation_name.split(',')[-1]
                            affiliation_name = affiliation_name.replace(', ' + persons_location, '')

                        current_page = int(base_name[1:])
                        """if current_page < start_page:
                            break"""

                        if start_page < current_page <= end_page:
                            names = name.split(', ')
                            if names:
                                for _ in names:
                                    record = (_.replace('\n', ''), affiliation_name.replace('\n', ''),
                                              persons_location.replace('\n', ''), session_name.replace('\n', ''),
                                              topic_title.replace('\n', ''), abstract.replace('-\n', '').replace('\n', ''))
                                    records.append(record)
                        elif current_page > end_page:
                            names = name.split(', ')
                            if names:
                                for _ in names:
                                    record = (_.replace('\n', ''), affiliation_name.replace('\n', ''),
                                              persons_location.replace('\n', ''), session_name.replace('\n', ''),
                                              topic_title.replace('\n', ''), abstract.replace('-\n', '').replace('\n', ''))
                                    records.append(record)
                            flag = True
                            break

                        session_name = text_line.get_text()

                        topic_title = ''
                        name = ''
                        affiliation_name = ''
                        persons_location = ''
                        abstract = ''
                    elif font_and_size == ('OAOVWE+TimesNewRomanPS-BoldMT', 9.0):
                        # lets filter line from super-scripted numbers of references
                        line = ''.join(x.get_text() for x in text_line if isinstance(x, LTChar) and x.size == 9.0)
                        topic_title += line
                        # topic_title += text_line.get_text()
                    elif font_and_size == ('OAOVWE+TimesNewRomanPS-ItalicMT', 9.0):
                        line = ''.join(x.get_text() for x in text_line if isinstance(x, LTChar) and x.size == 9.0)
                        if ':' not in line:
                            name += line
                    elif font_and_size == ('OAOVWE+TimesNewRomanPS-ItalicMT', 8.0):
                        line = ''.join(x.get_text() for x in text_line if isinstance(x, LTChar) and x.size == 8.0)
                        if ':' not in line:
                            affiliation_name += line
                    else:
                        abstract += text_line.get_text()
    return records


update_xml(file_xml, parse_articles(file_pdf, START_PAGE, END_PAGE), SHEET_NAME, START_ROW)
