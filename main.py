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
    row.append("barcode")
    row.append("weight")
    row.append("width")
    row.append("height")
    row.append("length")
    row.append("main_photo")
    row.append("additional_photos")


def get_stationery_row(input_row):
    """ Генерирует строку для записи в Канцелярские товары > Канцелярия """
    row = []
    get_basic_row(row)
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


def get_paper_products_row():
    """ Генерирует строку для записи в  Канцелярские товары > Бумажная продукция """
    row = []
    get_basic_row(row)
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
    row.append("density")  # плотность бумаги


def get_folders_files_row():
    """ Генерирует строку для записи в  Канцелярские товары > Папки и файлы """
    row = []
    get_basic_row(row)
    row.extend(["", ""])
    row.append("brand")
    row.append("title")
    row.append("color")
    row.append("color_name")
    row.append("type")
    row.append("sizes_string")
    row.append("")
    row.append("produced country")
    row.extend(['', '', '', ''])
    row.append("description")


def get_medical_devices_row():
    """ Генерирует строку для записи в  Аптека > Медицинские изделия """
    row = []
    get_basic_row(row)
    row.extend(["", ""])
    row.append("brand")
    row.append("title")
    row.append("count_in_pack")
    row.append("color")
    row.append("type")
    row.append("")
    row.append("produced country")
    row.extend(['', '', '', '', '', '', '', '', ''])
    row.append("description")


def get_medical_supplies_row():
    """ Генерирует строку для записи в  Аптека > Медицинские расходные материалы """
    row = []
    get_basic_row(row)
    row.extend(["", ""])
    row.append("brand")
    row.append("title")
    row.append("sizes_string")
    row.append("color")
    row.append("count_in_pack")  # в упаковке
    row.append("count_in_good")  # в товаре
    row.append("weight")
    row.append("type")
    row.append("expiration_date")  # срок годности
    row.extend(['', ''])
    row.append("description")
    row.extend(['', '', '', '', '', ''])
    row.append("produced country")


def get_cleaning_products_row():
    """ Генерирует строку для записи в  Бытовая химия > Моющие и чистящие средства """
    row = []
    get_basic_row(row)
    row.extend(["", ""])
    row.append("brand")
    row.append("title")
    row.append("sizes_string")
    row.append("expiration_date")  # срок годности
    row.append("type")
    row.append("")
    row.append("weight")
    row.extend(["", ""])
    row.append("produced country")
    row.extend(['', '', '', '', '', '', '', ''])
    row.append("description")
    row.append("")
    row.append("structure")


def get_air_freshener_row():
    """ Генерирует строку для записи в  Бытовая химия > Освежители воздуха """
    row = []
    get_basic_row(row)
    row.extend(["", ""])
    row.append("brand")
    row.append("title")
    row.append("smell")  # аромат
    row.append("volume")
    row.append("count_in_pack")  # в упаковке
    row.append("expiration_date")  # срок годности
    row.append("type")
    row.append("")
    row.append("")
    row.append("produced country")
    row.extend(['', '', '', '', '', '', ''])
    row.append("description")


def get_bags_row():
    """ Генерирует строку для записи в  Галантерея и украшения > Сумка """
    row = []
    get_basic_row(row)
    row.extend(["", ""])
    row.append("brand")
    row.append("title")
    row.append("")
    row.append("color")
    row.append("color_name")
    row.append("type")
    row.append("title")
    row.append("gender")
    row.extend(['', '', '', '', ''])
    row.append("description")
    row.append("")
    row.append("")
    row.append("produced country")


def get_food_accessories_row():
    """ Генерирует строку для записи в  Дом и сад > Аксессуары для приготовления пищи """
    row = []
    get_basic_row(row)
    row.extend(["", ""])
    row.append("brand")
    row.append("title")
    row.append("color")
    row.append("color_name")
    row.append("type")
    row.extend(['', '', ''])
    row.append("description")
    row.extend(['', ''])
    row.append("material")
    row.extend(['', ''])
    row.append("produced country")


def get_inventory_for_home_row():
    """ Генерирует строку для записи в Дом и сад > Инвентарь для дома  и Дом и сад > Инвентарь для уборки"""
    row = []
    get_basic_row(row)
    row.extend(["", ""])
    row.append("brand")
    row.append("title")
    row.append("color")
    row.append("color_name")
    row.append("type")
    row.extend(['', '', ''])
    row.append("description")
    row.extend(['', ''])
    row.append("material")
    row.extend(['', ''])
    row.append("produced country")


def get_inventory_for_cleaning_row():
    """ Генерирует строку для записи в Дом и сад > Инвентарь для уборки """
    row = []
    get_basic_row(row)
    row.extend(["", ""])
    row.append("brand")
    row.append("title")
    row.append("color")
    row.append("color_name")
    row.append("type")
    row.extend(['', '', ''])
    row.append("description")
    row.extend(['', ''])
    row.append("material")
    row.extend(['', ''])
    row.append("produced country")


def get_disposable_tableware_row():
    """ Генерирует строку для записи в Дом и сад > Одноразовая посуда """
    row = []
    get_basic_row(row)
    row.extend(["", ""])
    row.append("brand")
    row.append("title")
    row.append("color")
    row.append("color_name")
    row.append("type")
    row.extend(['', '', ''])
    row.append("description")
    row.extend(['', ''])
    row.append("material")
    row.extend(['', ''])
    row.append("produced country")


def get_dishes_row():
    """ Генерирует строку для записи в Дом и сад > Столовая посуда """
    row = []
    get_basic_row(row)
    row.extend(["", ""])
    row.append("brand")
    row.append("title")
    row.append("color")
    row.append("color_name")
    row.append("")
    row.append("type")
    row.extend(['', '', ''])
    row.append("description")
    row.extend(['', ''])
    row.append("material")
    row.extend(['', '', '', '', '', '', '', '', ''])
    row.append("produced country")


def get_things_storage_row():
    """ Генерирует строку для записи в Дом и сад > Хранение вещей """
    row = []
    get_basic_row(row)
    row.extend(["", ""])
    row.append("brand")
    row.append("title")
    row.extend(["", "", "", ""])
    row.append("type")
    row.extend(['', '', ''])
    row.append("description")
    row.extend(['', ''])
    row.append("material")
    row.extend(['', '', ''])
    row.append("produced country")


def get_paper_row():
    """ Генерирует строку для записи в Канцелярские товары > Бумага """
    row = []
    get_basic_row(row)
    row.extend(["", ""])
    row.append("brand")
    row.append("title")
    row.append("")
    row.append("sheets count")
    row.append("color")
    row.append("color_name")
    row.append("type")
    row.append("")
    row.append("produced country")
    row.extend(['', '', '', ''])
    row.append("description")


def get_demonstration_boards_row():
    """ Генерирует строку для записи в Канцелярские товары > Демонстрационные доски """
    row = []
    get_basic_row(row)
    row.extend(["", ""])
    row.append("brand")
    row.append("title")
    row.append("color")
    row.append("color_name")
    row.append("type")
    row.append("produced country")
    row.extend(['', '', '', ''])
    row.append("description")


def get_child_bags_row():
    """ Генерирует строку для записи в Канцелярские товары > Детские рюкзаки, ранцы, сумки """
    row = []
    get_basic_row(row)
    row.extend(["", ""])
    row.append("brand")
    row.append("title")
    row.append("color")
    row.append("color_name")
    row.append("type")
    row.append("")
    row.append("produced country")
    row.extend(['', '', '', '', '', '', ''])
    row.append("description")


def get_glue_row():
    """ Генерирует строку для записи в Канцелярские товары > Краска, клей """
    row = []
    get_basic_row(row)
    row.extend(["", ""])
    row.append("brand")
    row.append("title")
    row.append("color")
    row.append("color_name")
    row.append("volume")
    row.append("")
    row.append("expiration_date")  # срок годности
    row.append("type")
    row.append("")
    row.append("produced country")
    row.extend(['', '', ''])
    row.append("description")


def get_pencil_box_row():
    """ Генерирует строку для записи в Канцелярские товары > Пенал """
    row = []
    get_basic_row(row)
    row.extend(["", ""])
    row.append("brand")
    row.append("title")
    row.append("color")
    row.append("color_name")
    row.append("")
    row.append("")
    row.append("type")
    row.append("pencil_box_type")  # вид пенала ,
    row.append("number of branches")  # кол-во отделений
    row.append("filling")  # наполнение пенала
    row.append("")
    row.append("produced country")
    row.extend(['', '', '', '', ''])
    row.append("description")


def get_seal_and_stamp_row():
    """ Генерирует строку для записи в Канцелярские товары > Печати и штампы """
    row = []
    get_basic_row(row)
    row.extend(["", ""])
    row.append("brand")
    row.append("title")
    row.append("color")
    row.append("color_name")
    row.append("type")
    row.append("")
    row.append("produced country")
    row.append("material")
    row.extend(['', '', ''])
    row.append("description")


def get_writing_materials_row():
    """ Генерирует строку для записи в Канцелярские товары > Письменные принадлежности """
    row = []
    get_basic_row(row)
    row.extend(["", ""])
    row.append("brand")
    row.append("title")
    row.append("color")
    row.append("color_name")
    row.append("")
    row.append("type")
    row.append("")
    row.append("produced country")
    row.extend(['', '', '', '', ''])
    row.append("description")


def get_personal_hygiene_row():
    """ Генерирует строку для записи в Красота и Гигиена > Товары личной гигиены """
    row = []
    get_basic_row(row)
    row.extend(["", ""])
    row.append("title")
    row.append("brand")
    row.append("")
    row.append("")
    row.append("color")
    row.append("color_name")
    row.append("type")
    row.append("expiration_date")  # срок годности
    row.extend(['', ''])
    row.append("description")
    row.extend(['', '', ''])
    row.append("produced country")


def get_clothes_row():
    """ Генерирует строку для записи в Одежда > Одежда """
    row = []
    get_basic_row(row)
    row.extend(["", ""])
    row.append("brand")
    row.append("title")
    row.append("color")
    row.append("russian_size")
    row.append("our_size")
    row.append("color_name")
    row.append("type")
    row.append("gender")
    row.extend(['', '', '', '', '', ''])
    row.append("produced country")
    row.append('')
    row.append("description")


def get_juse_drinks_row():
    """ Генерирует строку для записи в Красота и Гигиена > Товары личной гигиены """
    row = []
    get_basic_row(row)
    row.extend(["", ""])
    row.append("brand")
    row.append("title")
    row.append("count_in_pack")
    row.append("weight")
    row.append("type")
    row.append("minimal_temp")
    row.append("maximum_temp")
    row.append("expiration_date")  # срок годности
    row.append("storage_conditions")  # условия хранения как на упаковке
    row.append("structure")
    row.extend(['', ''])
    row.append("description")
    row.extend(['', '', '', '', '', '', '', '', '', ''])
    row.append("produced country")


def get_bread_row():
    """ Генерирует строку для записи в  Продукты питания > Хлеб и кондитерские изделия """
    row = []
    get_basic_row(row)
    row.extend(["", ""])
    row.append("brand")
    row.append("title")
    row.append("weight")
    row.append("count_in_pack")
    row.append("type")
    row.append("minimal_temp")
    row.append("maximum_temp")
    row.append("expiration_date")  # срок годности
    row.append("storage_conditions")  # условия хранения как на упаковке
    row.append("structure")
    row.extend(['', '', '', ''])
    row.append("description")
    row.extend(['', '', '', '', '', '', '', '', '', '', '', ''])
    row.append("produced country")


def get_lamp_row():
    """ Генерирует строку для записи в Строительство и ремонт > Лампочка """
    row = []
    get_basic_row(row)
    row.extend(["", ""])
    row.append("brand")
    row.append("title")
    row.append("")
    row.append("count_in_pack")
    row.append("cokol_lamp")  # тип цоколя
    row.append("power")  # мощность, ВТ
    row.append("type")
    row.extend(['', '', ''])
    row.append("produced country")
    row.extend(['', '', '', '', '', ''])
    row.append("description")


def get_big_lamp_row():
    """ Генерирует строку для записи в Строительство и ремонт > Светильник """
    row = []
    get_basic_row(row)
    row.extend(["", ""])
    row.append("brand")
    row.append("title")
    row.extend(["", "", ""])
    row.append("type")
    row.extend(['', '', '', '', '', '', '', ''])
    row.append("description")
    row.extend(['', ''])
    row.append("produced country")


def get_fire_fighting_row():
    """ Генерирует строку для записи в Строительство и ремонт > Средства защиты и пожаротушения """
    row = []
    get_basic_row(row)
    row.extend(["", ""])
    row.append("brand")
    row.append("title")
    row.extend(["", ""])
    row.append("type")
    row.extend(['', ''])
    row.append("produced country")
    row.extend(['', '', '', ''])
    row.append("description")


def get_children_creativity_row():
    """ Генерирует строку для записи в Хобби и творчество > Детское творчество и развитие """
    row = []
    get_basic_row(row)
    row.extend(["", ""])
    row.append("brand")
    row.append("title")
    row.extend(["", ""])
    row.append("type")
    row.extend([''])
    row.append("produced country")
    row.extend(['', '', '', '', '', ''])
    row.append("description")


def get_set_for_creativity_row():
    """ Генерирует строку для записи в  Хобби и творчество > Набор для рукоделия, творчества """
    row = []
    get_basic_row(row)
    row.extend(["", ""])
    row.append("brand")
    row.append("title")
    row.extend(["", ""])
    row.append("type")
    row.extend([''])
    row.append("produced country")
    row.extend(['', '', '', '', '', ''])
    row.append("description")

def get_bag_row():
    """ Генерирует строку для записи в  Электроника > Рюкзаки, чехлы, сумки """
    row = []
    get_basic_row(row)
    row.extend(["", ""])
    row.append("brand")
    row.append("title")
    row.append("color")
    row.append("type")
    row.extend([''])
    row.append("produced country")
    row.extend(['', ''])
    row.append("description")