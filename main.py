import openpyxl


# Открываем xlsx-файл
def get_input_row(filename):
    wb = openpyxl.load_workbook(filename)

    # Получаем активный лист
    sheet = wb.active

    # Читаем и выводим каждую строку
    for row in sheet.iter_rows():
        row_data = []
        for cell in row:
            row_data.append(cell.value)
        print(row_data)


def get_basic_row(row):
    row.append("")
    row.append("article")  # артикул
    row.append("title")  # название товара
    row.append("price")
    row.append("price_before_sale")
    row.append("НДС")
    row.append("")
    row.append("Commercial_type")


def get_stationery_row(input_row):
    """ Генерирует строку для записи в Канцелярские товары > Канцелярия """
    row = []
    get_basic_row(row)
    row.append("barcode")
    row.append("weight")
    row.append("width")
    row.append("height")
    row.append("length")
    row.append("main_photo")
    row.append("additional_photos")
    row.extend(["", ""])
    row.append("brand")
    row.append("title")
    row.append("count_in_pack")
    row.append("color")
    row.append("color_name")
    row.append("color_name")
    row.append("type")
    row.append("produced country")
    row.extend(['', '', '', ''])
    row.append("description")
    row.extend(['', ''])
    row.append("material")
    return row


def get_paper_row():
    """ Генерирует строку для записи в  Канцелярские товары > Бумажная продукция """
    row = []
    get_basic_row(row)
    row.append("barcode")
    row.append("weight")
    row.append("width")
    row.append("height")
    row.append("length")
    row.append("main_photo")
    row.append("additional_photos")
    row.extend(["", ""])
    row.append("brand")
    row.append("title")
    row.append("sizes_string")
    row.append("sheets count")
    row.append("color")
    row.append("color_name")
    row.append("type")
    row.append("")
    row.append("produced country")
    row.extend(['', '', '', ''])
    row.append("description")
    row.append("")
    row.append("markup")  # разметка бумаги(в клетку/линейку или без нее)
    row.append("")
    row.append("density") # плотность бумаги
    