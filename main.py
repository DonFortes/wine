import datetime
from collections import defaultdict
from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas
from jinja2 import Environment, FileSystemLoader, select_autoescape

START_YEAR = 1920
FILE_NAME = "wine3.xlsx"
SHEET_NAME = "Лист1"


class Server:
    def __init__(self):
        self.env = Environment(
            loader=FileSystemLoader("."), autoescape=select_autoescape(["html", "xml"])
        )
        self.template = self.env.get_template("template.html")
        self.start_year = START_YEAR
        self.year_now = datetime.datetime.now().year
        self.delta = self.year_now - self.start_year

    def start_server_and_create_index(self, excel):
        rendered_page = self.template.render(
            year=self.delta,
            categories=excel.create_default_dict_to_render_page(),
        )
        with open("index.html", "w", encoding="utf8") as file:
            file.write(rendered_page)
        server = HTTPServer(("0.0.0.0", 8000), SimpleHTTPRequestHandler)
        return server.serve_forever()


class Excel:
    def __init__(self):
        self.excel_name = FILE_NAME
        self.sheet_name = SHEET_NAME
        self.data_frame_object = pandas.read_excel(
            self.excel_name,
            self.sheet_name,
            na_filter=False,
        )
        self.excel_data_dict = self.data_frame_object.to_dict(orient="records")

    def create_default_dict_to_render_page(self):
        dict_of_lists = defaultdict(list)
        for wine in self.excel_data_dict:
            dict_of_lists[wine["Категория"]].append(wine)
        return dict_of_lists


if __name__ == "__main__":
    server_obj = Server()
    excel_obj = Excel()
    server_obj.start_server_and_create_index(excel_obj)
