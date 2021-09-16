The task is to parse pdf-file to extract info from repetitive pattern of article which consists with blocks of fields. When I get structured data I will update xml-file.

Main point in my decision is to sort fields(or recognize) of article by searching most frequent font format(plain, italic, bold) and size; accumulate the list of records; update xml with received data from parser function.

Another possible ways to do this task are: 
- NLP(Natural Language Processing) but I need a million(my humble expectations) examples of articles;
- doing some research and find particular fields by regexp pattern;
- convert whole pdf-file to html and get data from BeautifulSoup(python library bs4) object;

I choose PyCharm IDE and installed some special libraries: pdfminer.six, openpyxl
To parse pdf I chose most popular pdf parser (for python 3 and above) pdfminer.six
To manage xml I chose openpyxl library.

My thoughts:
Pdf-file is 67 A4 pages consists of 3 different parts – part before the list of articles and part after so we need to think how can we skip redundant lines of text. Every page of the part of abstract book with articles divided by 2 columns, content of the right column continues the left column.

Let’s investigate pdf document, enlist some artifacts to recognize pattern:
every article has persistent order of blocks: article_number, title,  authors, institutions, abstract content;
every article consist with paragraphs of formatted text: I see different font: bold, plain, italic, font size, underlined text, super-scripted numbers of references, I see tables. Some paragraphs consists with different fonts . I can recognize most frequent font parameters of every text-line.

So my first decision was just print every line of text with tuple of two most frequent font-name and size. Then I manually enlisted these tuples.

I need to mention that this assignment or test case consists of two valuable parts: to extract data and then put the data in to existing xml file so my second part of work will be associated with researching of how to manage xml file with python.

I see some blockers:
- how to parse location because I can't recognize manually common pattern. Sometimes location specified, and some time no.
- not exactly understand it is necessary to split the field with article authors to different names and duplicate records in final xml-table with the same article, but different names. Or I need just write only one record for one article with comma separated list of authors?

After investigation or just playing around with parser(script) I understood that parser not working properly:
- (not solved) fix bug with location by hard-coding or something else; maybe I need just manually enlist locations from every articles and then just verify every word for inclusion in this list;
- (solved) I’m encapsulated parsing in one function and writing results to xml in another function
- (solved) I need to delete dashes at the end of lines in abstract block;
- (not solved) I missed only one word ‘References:’ or ‘Reference:’ in abstract block