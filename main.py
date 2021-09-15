from collections import Counter
from openpyxl import load_workbook
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextBoxHorizontal, LTChar


file_pdf = 'Abstract Book from the 5th World Psoriasis and Psoriatic Arthritis Conference 2018.pdf'
file_xml = 'Data Entry - 5th World Psoriasis & Psoriatic Arthritis Conference 2018 - Case format (2).xlsx'
START_PAGE = 100
END_PAGE = 104


def update_xml(file, rec):
    """
    param re: list of tuples aka records
    param file: file xml for read and update with bunch of records
    """
    wb = load_workbook(file)
    ws = wb['Sheet1']
    for row in range(6, len(rec)+6):
        for col in range(1, 7):
            ws.cell(row=row, column=col).value = rec[row-6][col-1]

    wb.save(file)
    wb.close()


flag = False

session_name = ''
topic_title = ''
name = ''
affiliation_name = ''
persons_location = ''
abstract = ''
records = []

for page_layout in extract_pages(file_pdf):
    # flag for breaking a loop if we reach END_PAGE
    if flag:
        break

    for element in page_layout:
        if isinstance(element, LTTextBoxHorizontal):

            for text_line in element:
                # do some investigation about most frequent font and size of one pdf line
                quirk = [(character.fontname, character.size) for character in text_line if
                         isinstance(character, LTChar)]
                font_and_size = Counter(quirk).most_common(1)[0][0]

                if font_and_size == ('OAOVWE+TimesNewRomanPS-BoldItalicMT', 9.5):
                    base_name = text_line.get_text()
                    if affiliation_name:
                        persons_location = affiliation_name.split(',')[-1]
                        affiliation_name = affiliation_name.replace(persons_location, '')

                    current_page = int(base_name[1:])

                    if START_PAGE < current_page <= END_PAGE:
                        names = name.split(', ')
                        if names:
                            for x in names:
                                record = (x.replace('\n', ''), affiliation_name.replace('\n', ''), persons_location.replace('\n', ''), session_name.replace('\n', ''), topic_title.replace('\n', ''), abstract.replace('\n', ''))
                                records.append(record)
                        update_xml(file_xml, records)
                    elif current_page > END_PAGE:
                        names = name.split(', ')
                        if names:
                            for x in names:
                                record = (x.replace('\n', ''), affiliation_name.replace('\n', ''), persons_location.replace('\n', ''), session_name.replace('\n', ''), topic_title.replace('\n', ''), abstract.replace('\n', ''))
                                records.append(record)
                        update_xml(file_xml, records)
                        flag = True
                        break

                    session_name = text_line.get_text()

                    topic_title = ''
                    name = ''
                    affiliation_name = ''
                    persons_location = ''
                    abstract = ''
                elif font_and_size == ('OAOVWE+TimesNewRomanPS-BoldMT', 9.0):
                    line = ''.join(x.get_text() for x in text_line if isinstance(x, LTChar) and x.size == 9.0)
                    topic_title += line
                    # topic_title += text_line.get_text()
                elif font_and_size == ('OAOVWE+TimesNewRomanPS-ItalicMT', 9.0):
                    line = ''.join(x.get_text() for x in text_line if isinstance(x, LTChar) and x.size == 9.0)
                    if ':' not in line:
                        name += line
                        # name += text_line.get_text()
                elif font_and_size == ('OAOVWE+TimesNewRomanPS-ItalicMT', 8.0):
                    line = ''.join(x.get_text() for x in text_line if isinstance(x, LTChar) and x.size == 8.0)
                    if ':' not in line:
                        affiliation_name += line
                        # affiliation_name += text_line.get_text()
                else:
                    abstract += text_line.get_text()



