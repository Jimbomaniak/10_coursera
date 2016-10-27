from lxml import etree
import random
import requests
import re
from openpyxl import Workbook


COURSES_NUMBER = 20
FILE_NAME_TO_SAVE = 'random_courses_info'


def get_courses_urls():
    url = 'https://www.coursera.org/sitemap~www~courses.xml'
    response = requests.get(url)
    xml = response.content
    root = etree.XML(xml)
    sites_in_string = root.xpath("string()")
    pattern = r'(https.+www.coursera.org.learn.[a-z0-9-]+)'
    all_urls_list = re.findall(pattern, sites_in_string)
    return all_urls_list


def get_course_info(course_url):
    response = requests.get(course_url)
    html_bytes = response.content
    tree = etree.HTML(html_bytes)
    title = tree.xpath('string(//h1)')
    weeks = len(tree.xpath('//div[@class="week"]'))
    info = tree.xpath('string(//script[@type="application/ld+json"])')
    dates = re.search(r'(\d{4}.\d{2}.\d{2})', info)
    start_date = dates and dates.group() or 'No specific dates'
    language = tree.xpath('//tr/td/div[@class="language-info"]/text()')[0]
    avarage_rating = tree.xpath('//div[@class="ratings-text bt3-hidden-xs"]/text()')
    return {
        'title': title,
        'weeks': weeks,
        'start_date': start_date,
        'language': language,
        'avarage_rating': avarage_rating and avarage_rating[0] or 'No rating yet'
    }

def fetch_random_courses_data(courses_urls, courses_number):
    random_courses_urls = [] 
    random_courses_data = []
    random_course_url = random.choice(courses_urls)
    while len(random_courses_urls) < courses_number:
        random_course_url = random.choice(courses_urls)
        if random_course_url not in random_courses_urls:
            random_courses_urls.append(random_course_url)
        else:
            continue
    for course_url in random_courses_urls:
        random_courses_data.append(get_course_info(course_url))
    return random_courses_data


def output_courses_info_to_xlsx(random_courses_data, name_to_save):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Courses information'
    ws.append(['Course Title', 'Weeks','Start Date',
                 'Language','Avarage Rating'])
    ws.column_dimensions['A'].width = 60
    ws.column_dimensions['B'].width = 5
    ws.column_dimensions['C'].width = 13
    ws.column_dimensions['D'].width = 15
    ws.column_dimensions['E'].width = 20
    for course in random_courses_data:
        ws.append([
            course['title'], course['weeks'],
            course['start_date'], course['language'],
            course['avarage_rating']
            ])
    wb.save("{0}.xlsx".format(name_to_save))


if __name__ == '__main__':
    courses_urls = get_courses_urls()
    random_courses_data = fetch_random_courses_data(courses_urls, COURSES_NUMBER)
    output_courses_info_to_xlsx(random_courses_data, FILE_NAME_TO_SAVE)
    print('Done. Courses info is in script folder - {0}.xlsx'.format(FILE_NAME_TO_SAVE))



