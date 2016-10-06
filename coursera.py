import requests
import random
from lxml import etree
import requests
from bs4 import BeautifulSoup
import json
from openpyxl import Workbook
from openpyxl.styles import Font, Border,  Side, Color, PatternFill
from openpyxl.styles import colors
import os.path
import argparse
import sys


URL_TO_COURSERA_XML = 'https://www.coursera.org/sitemap~www~courses.xml'
NUMBER_ANALYZED_COURSES = 20
TEXT_FOR_LINK = 'подробнее'
COLUMNS_ORDER = {
                 'name_course': [1, 'Имя'],
                 'language_course': [2, 'Язык'],
                 'starts_course': [3, 'Дата начала'],
                 'number_weeks_course': [4, 'Продолжительность (недель)'],
                 'rating_course': [5, 'Рейтинг'],
                 'course_url': [6, 'URL']
                }


def create_parser():
    parser = argparse.ArgumentParser(description='Скрипт выполняет сбор \
                                     информации о разных курсах на Курсере.')
    parser.add_argument('-f', '--file', metavar='ФАЙЛ',
                        help='Имя файла для выгрузки информации.')
    return parser


def check_filepath(filepath):
    (dir_name, file_name) = os.path.split(os.path.abspath(filepath))
    if os.path.exists(dir_name):
        (root, ext) = os.path.splitext(filepath)
        if ext not in ['.xls', '.xlsx']:
            print('Файл должен быть формата "xls" или "xlsx".')
            return False
    else:
        print('Каталог не существует!')
        return False
    return True


def get_courses_list():
    courses_list = []
    response = requests.get(URL_TO_COURSERA_XML, stream=True)
    response.raw.decode_content = True
    tree = etree.parse(response.raw)
    root = tree.getroot()
    for num in range(NUMBER_ANALYZED_COURSES):
        num_elem = random.randrange(0, len(root)-1)
        courses_list.append(root[num_elem][0].text)
    return courses_list


def get_tag_text(tag):
    if tag:
        return tag.text
    else:
        return ''


def get_course_info(course_url):
    name_course = ''
    language_course = ''
    starts_course = ''
    number_weeks_course = ''
    rating_course = ''
    course_info = {}
    response = requests.get(course_url)
    response.encoding = 'utf-8'
    page = response.text
    soup = BeautifulSoup(page, 'lxml')
    # имя курса
    name_course = get_tag_text(soup.find('div', {'class':
                                         'title display-3-text'}))
    # рейтин курса
    rating_course = get_tag_text(soup.find('div', {'class':
                                           'ratings-text bt3-visible-xs'}))
    # количество недель
    number_weeks_course = len(soup.findAll('div', {'class': 'week'}))
    # дата начала курса (находится в данных java-script)
    text_json = get_tag_text(soup.find('script', {'type':
                                       'application/ld+json'}))
    if text_json:
        json_data = json.loads(text_json)
        starts_course = json_data['hasCourseInstance'][0]['startDate']
    # язык курса (находится в таблице)
    table = soup.find('table', {'class':
                                'basic-info-table bt3-table bt3-table-striped '
                                'bt3-table-bordered bt3-table-responsive'})
    if table:
        all_cols = []
        for row in table:
            for col in row.find_all('td'):
                all_cols.append(col.text)
        col_name_language = all_cols.index('Language')
        if col_name_language:
            language_course = all_cols[col_name_language+1]
    # наполним словарь параметров курса полученной информацией
    for name in ['name_course', 'language_course', 'starts_course',
                 'number_weeks_course', 'rating_course', 'course_url']:
        course_info[name] = eval(name)
    return course_info


def output_courses_info_to_xlsx(courses_info, filepath):
    wb = Workbook()
    sheet = wb.active
    # параметры оформления
    thin_border = Border(left=Side(style='thin'),
                         right=Side(style='thin'),
                         top=Side(style='thin'),
                         bottom=Side(style='thin'))
    darkgray_fill = PatternFill(start_color='A9A9A9',
                                end_color='A9A9A9',
                                fill_type='solid')
    lightgray_fill = PatternFill(start_color='D3D3D3',
                                 end_color='D3D3D3',
                                 fill_type='solid')
    # шапка таблицы
    for item in COLUMNS_ORDER.items():
        cell = sheet.cell(row=1, column=item[1][0])
        cell.value = item[1][1]
        cell.font = Font(bold=True)
        cell.border = thin_border
        cell.fill = darkgray_fill
    # тело таблицы
    for row, course in enumerate(courses_info, start=2):
        for item in course.items():
            cell = sheet.cell(row=row, column=COLUMNS_ORDER[item[0]][0])
            if item[0] == 'course_url':
                cell.value = '=HYPERLINK("%s","%s")' % (item[1], TEXT_FOR_LINK)
            else:
                cell.value = item[1]
            cell.border = thin_border
            if row % 2:
                cell.fill = lightgray_fill
    sheet.column_dimensions['A'].width = 50
    wb.save(filepath)
    return True


if __name__ == '__main__':
    parser = create_parser()
    namespace = parser.parse_args()
    if namespace.file:
        filepath = namespace.file
    else:
        filepath = input('Введите имя файла для выгрузки данных \
                         в формате "xls" или "xlsx":\n')
    if not check_filepath(filepath):
        sys.exit()
    courses_list = get_courses_list()
    courses_info = [get_course_info(course) for course in courses_list]
    if output_courses_info_to_xlsx(courses_info, filepath):
        print('Информация о курсах выгружена в файл "%s"' %
              os.path.abspath(filepath))
