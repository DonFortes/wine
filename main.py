import datetime
from collections import defaultdict
from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas
from jinja2 import Environment, FileSystemLoader, select_autoescape

START_YEAR = 1920
FILE_PATH = "excel_files/wines.xlsx"
SHEET_NAME = "Лист1"


class Server:

    """Class to create index.html and start the server."""

    def __init__(self, start_year):
        self.env = Environment(
            loader=FileSystemLoader(""), autoescape=select_autoescape(["html", "xml"])
        )
        self.template = self.env.get_template("template.html")
        self.start_year = start_year
        self.year_now = datetime.datetime.now().year
        self.age = self.year_now - self.start_year

    def create_index_html(self, _excel):

        """Creates index.html."""

        rendered_page = self.template.render(
            winery_age=self.age,
            grouped_wines=_excel.create_grouped_wines(),
        )
        with open("index.html", "w", encoding="utf8") as file:
            file.write(rendered_page)

    def main(self):

        """Starts the server."""

        _server = HTTPServer(("0.0.0.0", 8000), SimpleHTTPRequestHandler)
        return _server.serve_forever()


class Excel:

    """Class for work with excel files."""

    def __init__(self, excel_path, sheet_name):
        self.excel_path = excel_path
        self.sheet_name = sheet_name
        self.wines = pandas.read_excel(
            self.excel_path,
            self.sheet_name,
            na_filter=False,
        ).to_dict(orient="records")

    def create_grouped_wines(self):

        """Creates groups of drinks."""

        wines_sorted_by_group = defaultdict(list)
        for wine in self.wines:
            wines_sorted_by_group[wine["Категория"]].append(wine)
        return wines_sorted_by_group


if __name__ == "__main__":
    server = Server(START_YEAR)
    excel = Excel(FILE_PATH, SHEET_NAME)
    server.create_index_html(excel)
    server.main()
