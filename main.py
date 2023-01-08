from bs4 import BeautifulSoup
import requests
import urllib.parse
import openpyxl
import logging
import json
import time
from pathlib import Path
from jsonpath_ng import parse

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

c_handler = logging.StreamHandler()
c_handler.setLevel(logging.DEBUG)
c_format = logging.Formatter('%(asctime)s - %(levelname)s: %(message)s', datefmt='%H:%M:%S')
c_handler.setFormatter(c_format)
logger.addHandler(c_handler)

f_handler = logging.FileHandler('data_log.log')
f_handler.setLevel(logging.WARNING)
f_format = logging.Formatter('%(asctime)s - %(levelname)s: %(message)s', datefmt='%d-%m-%Y %H:%M:%S')
f_handler.setFormatter(f_format)
logger.addHandler(f_handler)

address = 'https://lubimyczytac.pl/'
file_path = Path('ebook.xlsx')
try:
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active
except IOError:
    logger.info('The file is missing, a new one was created.')
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.append(
        ["Author", "Title", "Folder", "Rating", "Votes", "Category", "Date of the first edition", "Date of the "
                                                                                                  "first PL "
                                                                                                  "edition",
         "Pages No", "Tags", "Description"])
    wb.save('ebook.xlsx')


def main():
    logger.info('Getting Started...\n')

    lc_headers = {'X-Csrf-Token': 'b42ee05dc19f593fb108c70c11a17e22', 'X-Requested-With': 'XMLHttpRequest'}
    counter = 1
    missing_items = []
    for i in range(row_from, row_to):
        title = sheet.cell(row=i + 1, column=2)
        author = sheet.cell(row=i + 1, column=1).value
        surname = author.split(' ', 1)[1] if author.find(' ') > 0 else author
        logger.info('Search item: %s, %s.', title.value, surname)
        uri1 = urllib.parse.quote(title.value + ' ' + surname)
        uri2 = urllib.parse.quote(title.value)

        url = 'https://lubimyczytac.pl/searcher/getsuggestions?phrase=' + uri1
        url2 = 'https://lubimyczytac.pl/searcher/getsuggestions?phrase=' + uri2

        def book_url(url):
            content = requests.get(url, headers=lc_headers).text
            soup = BeautifulSoup(content, 'html.parser')
            json_text = soup.get_text()
            json_data = json.loads(json_text)
            jsonpath_expression = parse('$.items.books.results[0].url')
            match = jsonpath_expression.find(json_data)
            return match

        try:
            try:
                book_match = book_url(url)
                title_url = book_match[0].value
                logger.info('Loading the website %s.', title_url)
            except:
                book_match = book_url(url2)
                title_url = book_match[0].value
                logger.info('Loading the website %s.', title_url)
        except:
            logger.warning('Item not found!')
            missing_items.append(title.value + ' - ' + author)
            counter += 1
            continue

        content2 = requests.get(title_url).text
        soup = BeautifulSoup(content2, 'html.parser')
        description = soup.find("div", {"class": "collapse-content"}).find('p')

        rate = soup.find(class_="big-number").get_text()
        sheet.cell(row=i + 1, column=4).value = rate[1:]

        raters_no = soup.find(class_="d-none d-lg-block book-on-shelfs__rating--1").get_text()
        detail_list = raters_no.split(' \n\n')

        logger.info('%s. %s - saving the following data:', i, title.value)

        voters = (detail_list[2].replace("\n", "")).split(' ')[0] if len(detail_list) > 2 else 0
        logger.info('   - %s;', voters)
        try:
            sheet.cell(row=i + 1, column=5).value = int(voters)
        except:
            sheet.cell(row=i + 1, column=5).value = voters

        logger.info('   - %s;', soup.find(class_='book__category d-sm-block d-none').get_text()[1:])
        sheet.cell(row=i + 1, column=6).value = soup.find(class_='book__category d-sm-block d-none').get_text()[1:]

        description_text = description.get_text().replace('\n', '') if description is not None else 'Brak opisu.'
        logger.info('   - %s(...);', description_text[:50])
        sheet.cell(row=i + 1, column=11).value = description_text

        details = soup.find(id="book-details").get_text()
        row_details = soup.find(id="book-details")

        keys = row_details.find_all('dt')
        values = row_details.find_all('dd')

        details_dict = {}
        for key in keys:
            for value in values:
                details_dict[key.get_text().replace('\n', '')] = value.get_text().replace('\n', '')
                values.remove(value)
                break

        def cell_filler(tag, column_no, convert_type='str'):
            if tag in details_dict:
                logger.info('   - %s;', details_dict[tag].replace('\n', ''))
                sheet.cell(row=i + 1, column=column_no).value = details_dict[tag].replace('\n',
                                                                                          '') if convert_type == 'str' else int(
                    details_dict[tag].replace('\n', ''))
            else:
                sheet.cell(row=i + 1, column=column_no).value = ""

        cell_filler('Data 1. wydania:', 7)
        cell_filler('Data 1. wyd. pol.:', 8)
        cell_filler('Liczba stron:', 9, 'int')
        cell_filler('Tagi:', 10)
        logger.info('Completed %s percent.\n', round((i - row_from + 1) * 100 / (row_to - row_from)))
        i = 0
        while i < 3:
            try:
                wb.save(file_path)
                break
            except PermissionError:
                i += 1
                print(f"You Don't Have Permission to Access the File, retry in 15 sec.")
                time.sleep(15)

    if (counter - 1) == 0:
        logger.info('All %s items have been found!', (row_to - row_from))
    else:
        logger.info('The following items were not found (%s of %s): %s', counter - 1, (row_to - row_from), missing_items)

while True:
    print("...::: MENU :::...")
    print("1. Find all items")
    print("2. Find selected items")
    print("3. Info")
    print("4. Exit")

    option = input("Select an option: ")

    if option == "1":
        print(f'In the active sheet {sheet.title} I write the information for {sheet.max_row - sheet.min_row} items.')
        row_from = int(sheet.min_row)
        row_to = int(sheet.max_row)
        main()
    elif option == "2":
        print(f'The active sheet {sheet.title} has rows {sheet.min_row} through {sheet.max_row} filled in.')
        row_from = int(input('Row from: '))
        row_to = int(input('Row to: '))
        main()
    elif option == "3":
        print(f'There are {sheet.max_row - sheet.min_row} items in the active sheet {sheet.title} (rows {sheet.min_row} to {sheet.max_row}).')
        print(f'The other sheets: {wb.sheetnames}')
    elif option == "4":
        break
    else:
        print("Incorrect entry, try again!")