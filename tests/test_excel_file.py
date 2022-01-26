import main

FILE_PATH = "excel_files/wines.xlsx"
SHEET_NAME = "Лист1"
excel = main.Excel(FILE_PATH, SHEET_NAME)


def test_category_in_table():
    grouped_wines_params = list(excel.create_grouped_wines().items())[0][1][0]
    assert (
        "Категория" in grouped_wines_params.keys()
    ), "Проверьте, что столбец 'Категория' есть в вашей таблице"


def test_name_in_table():
    grouped_wines_params = list(excel.create_grouped_wines().items())[0][1][0]
    assert (
        "Категория" in grouped_wines_params.keys()
    ), "Проверьте, что столбец 'Категория' есть в вашей таблице"


def test_sort_in_table():
    grouped_wines_params = list(excel.create_grouped_wines().items())[0][1][0]
    assert (
        "Сорт" in grouped_wines_params.keys()
    ), "Проверьте, что столбец 'Сорт' есть в вашей таблице"


def test_price_in_table():
    grouped_wines_params = list(excel.create_grouped_wines().items())[0][1][0]
    assert (
        "Цена" in grouped_wines_params.keys()
    ), "Проверьте, что столбец 'Цена' есть в вашей таблице"


def test_image_in_table():
    grouped_wines_params = list(excel.create_grouped_wines().items())[0][1][0]
    assert (
        "Картинка" in grouped_wines_params.keys()
    ), "Проверьте, что столбец 'Картинка' есть в вашей таблице"


def test_stock_in_table():
    grouped_wines_params = list(excel.create_grouped_wines().items())[0][1][0]
    assert (
        "Акция" in grouped_wines_params.keys()
    ), "Проверьте, что столбец 'Акция' есть в вашей таблице"
