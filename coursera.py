from lxml import etree
from bs4 import BeautifulSoup
from openpyxl import Workbook

import requests
import random
import json
import datetime
import re


def get_courses_list(courses_count=20):
    data = requests.get('https://www.coursera.org/sitemap~www~courses.xml')
    root = etree.fromstring(data.content)

    return random.sample(
        [child[0].text for child in root],
        courses_count
    )


def get_course_info(course_url):
    course_data = {}
    response = requests.get(course_url)
    html_doc = BeautifulSoup(response.content, 'html.parser')
    api_template = 'https://api.coursera.org/api/courses.v1?q=slug&'\
        'slug={slug}&fields=startDate'

    title = get_by_class(html_doc, 'title display-3-text')
    language = get_by_class(html_doc, 'language-info')
    starts = None
    commitment = 0
    
    # try to get course info from html
    course_info = html_doc.find('script',type='application/ld+json')
    if course_info:
        course = json.loads(course_info.string)['hasCourseInstance'][0]
        start_date = course.get('startDate')
        if start_date:
            start_date = datetime.datetime.strptime(
                course['startDate'],
                '%Y-%m-%d'
            )
            end_date = course.get('endDate')
            if end_date:    
                end_date = datetime.datetime.strptime(
                    course['endDate'],
                    '%Y-%m-%d'
                )
                delta = end_date - start_date
                commitment = delta.days / 7
            starts = start_date
    else: # no json in html? try to use api
        match = re.search(r'/[^/]+$', course_url)
        slug = match.group(0)[1:]
        response = requests.get(api_template.format(slug=slug))
        course = response.json()['elements'][0]
        start_date = course.get('startDate')
        if start_date:
            starts = datetime.datetime.fromtimestamp(int(start_date)/1000)

    str_rating = get_by_class(html_doc, 'ratings-text bt3-visible-xs')
    rating = float(str_rating[:4]) if str_rating else 0

    course_data['course_url'] = course_url
    course_data['title'] = title
    course_data['language'] = language
    course_data['starts'] = starts
    course_data['commitment'] = commitment
    course_data['rating'] = rating
    return course_data


def get_by_class(html_doc, class_info):
    result = html_doc.find(class_=class_info)
    if not result:
        return None
    elif not result.string and result.contents:
        return result.contents[1]
    return result.string


def output_courses_info_to_xlsx(data, filepath):
    wb = Workbook()
    ws = wb.active
    keys = ['title', 'starts', 'language',
        'commitment', 'rating', 'course_url'
    ]
    for column, key in enumerate(keys, start=1):
        ws.cell(row=1, column=column, value=key)

    for row, item in enumerate(data, start=2):
        for column, key in enumerate(keys, start=1):
            ws.cell(row=row, column=column, value=item[key])

    wb.save(filepath)


if __name__ == '__main__':
    output_courses_info_to_xlsx(
        (get_course_info(course_url) for course_url in get_courses_list(20)),
        input('Enter filepath (.xlsx):')
    )
        