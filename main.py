import datetime
from collections import defaultdict
from http.server import HTTPServer, SimpleHTTPRequestHandler
from pprint import pprint

import pandas
from jinja2 import Environment, FileSystemLoader, select_autoescape

START_YEAR = 1920
FILE_NAME = "excel_files/wine3.xlsx"
SHEET_NAME = "Лист1"


class Server:
    def __init__(self):
        self.env = Environment(
            loader=FileSystemLoader("."), autoescape=select_autoescape(["html", "xml"])
        )
        self.template = self.env.get_template("template.html")
        self.start_year = START_YEAR
        self.year_now = datetime.datetime.now().year
        self.age = self.year_now - self.start_year

    def start_server_and_create_index(self, excel):
        rendered_page = self.template.render(
            winery_age=self.age,
            grouped_wines=excel.get_grouped_wines(),
        )
        with open("index.html", "w", encoding="utf8") as file:
            file.write(rendered_page)

    def main(self):
        _server = HTTPServer(("0.0.0.0", 8001), SimpleHTTPRequestHandler)
        return _server.serve_forever()


class Excel:
    def __init__(self):
        self.excel_name = FILE_NAME
        self.sheet_name = SHEET_NAME
        self.wines = pandas.read_excel(
            self.excel_name,
            self.sheet_name,
            na_filter=False,
        ).to_dict(orient="records")
        # Меня просто учили, что дробить на части - это лучше, чем
        # делать многосоставные "предложения". Проще дебажить в случае чего
    def get_grouped_wines(self):
        wines_sorted_by_group = defaultdict(list)
        for wine in self.wines:
            wines_sorted_by_group[wine["Категория"]].append(wine)
        return wines_sorted_by_group


if __name__ == "__main__":
    server = Server()
    excel = Excel()
    server.create_index_html(excel)
    pprint(excel.get_grouped_wines())
    # server.main()
