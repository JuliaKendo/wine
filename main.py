import os
import argparse
import datetime
import pandas
import collections
from dotenv import load_dotenv
from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape

YEAR_OF_FOUNDATION = 1920


def get_age():
    current_date = datetime.datetime.now()
    return current_date.year - YEAR_OF_FOUNDATION


def get_wines_by_category(file):
    excel_data_df = pandas.read_excel(file, na_values=' ', keep_default_na=False)
    wines = excel_data_df.to_dict(orient='record')

    wines_by_category = collections.defaultdict(list)

    for wine in wines:
        wines_by_category[wine[u'Категория']].append(wine)

    return wines_by_category


def create_parser():
    parser = argparse.ArgumentParser(description='Параметры запуска скрипта')
    parser.add_argument('-f', '--file', default='wine.xlsx', help='Файл excel с информацией для публикации на сайте')
    parser.add_argument('-t', '--template', default='template.html', help='Шаблон html')
    return parser


def main():

    load_dotenv()
    parser = create_parser()
    args = parser.parse_args()
    server_ip = os.getenv('SERVER_IP')
    server_port = int(os.getenv('SERVER_PORT'))
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    try:
        template = env.get_template(args.template)

        rendered_page = template.render(
            categories=get_wines_by_category(args.file).items(),
            age=get_age()
        )

    except (ValueError, TypeError, KeyError) as error:
        print('%s' % error)

    else:
        with open('index.html', 'w', encoding="utf8") as file:
            file.write(rendered_page)

        server = HTTPServer((server_ip, server_port), SimpleHTTPRequestHandler)
        server.serve_forever()


if __name__ == '__main__':
    main()
