import openpyxl
import json


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


def get_main_image(images):
    space_index = images.find(" ")
    if space_index != -1:
        return images[:space_index]
    else:
        if len(images) > 5:
            return images
        else:
            return ""


def get_additional_images(images):
    space_index = images.find(" ")
    if space_index != -1:
        return images[space_index + 1:]
    else:
        return ""


def get_basic_row(row, input_row, commercial_type):
    row.append("")
    row.append(input_row[0])  # артикул
    row.append(input_row[8])  # название товара
    row.append(input_row[5])
    row.append(input_row[6])
    row.append("Не облагается")
    row.append("")
    row.append(commercial_type)
    row.append("")
    row.append(input_row[1])
    row.append(input_row[3])
    row.append(input_row[4])
    row.append(input_row[2])
    row.append(get_main_image(input_row[10]))
    row.append(get_additional_images(input_row[10]))


def get_count_in_pack(input_row):
    input_count = input_row[12]
    if input_count == "" or not input_count:
        return "1"
    else:
        return input_count


def get_color_name(input_row):
    return input_row[13]


def get_brand(input_row):
    return input_row[11]


def get_title(input_row):
    return input_row[8]


def get_type(commercial_type):
    categories_match_file = open("types.json", 'r', encoding='utf-8')
    categories_match_dict = json.load(categories_match_file)
    return categories_match_dict.get(commercial_type, "")


def get_description(input_row):
    return input_row[9]


def get_stationery_row(input_row, commercial_type):
    """ Генерирует строку для записи в Канцелярские товары > Канцелярия """
    row = []
    get_basic_row(row, input_row, commercial_type)
    row.extend(["", ""])
    row.append(get_brand(input_row))
    row.append(get_title(input_row))
    row.append(get_count_in_pack(input_row))
    row.append("")
    row.append(get_color_name(input_row))
    row.append(get_type(commercial_type))
    row.append("")
    row.extend(['', '', '', ''])
    row.append(get_description(input_row))
    row.extend(['', ''])
    row.append("")
    return row


def get_sheets_count(input_row):
    sheets_count = input_row[14]
    if sheets_count == "" or not sheets_count:
        return '100'
    else:
        return sheets_count


def get_density(input_row):
    density = input_row[15]
    if density == "" or not density:
        return '60'
    else:
        return density


def get_paper_products_row(input_row, commercial_type):
    """ Генерирует строку для записи в  Канцелярские товары > Бумажная продукция """
    row = []
    get_basic_row(row, input_row, commercial_type)
    row.extend(["", ""])
    row.append(get_brand(input_row))
    row.append(get_title(input_row))
    row.append("")
    row.append(get_sheets_count(input_row))
    row.append("")
    row.append(get_color_name(input_row))
    row.append(get_type(commercial_type))
    row.append("")
    row.append("")
    row.extend(['', '', '', ''])
    row.append(get_description(input_row))
    row.append("")
    row.append("")
    row.append("")
    row.append(get_density(input_row))  # плотность бумаги
    return row


def get_folders_files_row(input_row, commercial_type):
    """ Генерирует строку для записи в  Канцелярские товары > Папки и файлы """
    row = []
    get_basic_row(row, input_row, commercial_type)
    row.extend(["", ""])
    row.append(get_brand(input_row))
    row.append(get_title(input_row))
    row.append("")
    row.append(get_color_name(input_row))
    row.append(get_type(commercial_type))
    row.append("")
    row.append("")
    row.append("")
    row.extend(['', '', '', ''])
    row.append(get_description(input_row))


def get_medical_devices_row(input_row, commercial_type):
    """ Генерирует строку для записи в  Аптека > Медицинские изделия """
    row = []
    get_basic_row(row, input_row, commercial_type)
    row.extend(["", ""])
    row.append(get_brand(input_row))
    row.append(get_title(input_row))
    row.append("")
    row.append("")
    row.append(get_type(commercial_type))
    row.append("")
    row.append("")
    row.extend(['', '', '', '', '', '', '', '', ''])
    row.append(get_description(input_row))


def get_size(input_row):
    size = input_row[16]
    if not size:
        return ""
    else:
        return size


def get_expiration_date():
    return '730'


def get_medical_supplies_row(input_row, commercial_type):
    """ Генерирует строку для записи в  Аптека > Медицинские расходные материалы """
    row = []
    get_basic_row(row, input_row, commercial_type)
    row.extend(["", ""])
    row.append(get_brand(input_row))
    row.append(get_title(input_row))
    row.append("")
    row.append("")
    row.append("")  # в упаковке
    row.append("")  # в товаре
    row.append("")
    row.append(get_type(commercial_type))
    row.append(get_expiration_date())  # срок годности
    row.extend(['', ''])
    row.append(get_description(input_row))


def get_cleaning_products_row(input_row, commercial_type):
    """ Генерирует строку для записи в  Бытовая химия > Моющие и чистящие средства """
    row = []
    get_basic_row(row, input_row, commercial_type)
    row.extend(["", ""])
    row.append(get_brand(input_row))
    row.append(get_title(input_row))
    row.append("")
    row.append(get_expiration_date())  # срок годности
    row.append(get_type(commercial_type))
    row.append("")
    row.append("")
    row.extend(["", ""])
    row.append("")
    row.extend(['', '', '', '', '', '', '', ''])
    row.append(get_description(input_row))


def get_air_freshener_row(input_row, commercial_type):
    """ Генерирует строку для записи в  Бытовая химия > Освежители воздуха """
    row = []
    get_basic_row(row, input_row, commercial_type)
    row.extend(["", ""])
    row.append(get_brand(input_row))
    row.append(get_title(input_row))
    row.append("")
    row.append("")
    row.append("")  # в упаковке
    row.append(get_expiration_date())  # срок годности
    row.append(get_type(commercial_type))
    row.append("")
    row.append("")
    row.append("")
    row.extend(['', '', '', '', '', '', ''])
    row.append(get_description(input_row))


def get_bags_row(input_row, commercial_type):
    """ Генерирует строку для записи в  Галантерея и украшения > Сумка """
    row = []
    get_basic_row(row, input_row, commercial_type)
    row.extend(["", ""])
    row.append(get_brand(input_row))
    row.append(get_title(input_row))
    row.append("")
    row.append("")
    row.append(get_color_name(input_row))
    row.append(get_type(commercial_type))
    row.append(get_title(input_row))
    row.append("")
    row.extend(['', '', '', '', ''])
    row.append(get_description(input_row))


def get_food_accessories_row(input_row, commercial_type):
    """ Генерирует строку для записи в  Дом и сад > Аксессуары для приготовления пищи """
    row = []
    get_basic_row(row, input_row, commercial_type)
    row.extend(["", ""])
    row.append(get_brand(input_row))
    row.append(get_title(input_row))
    row.append("")
    row.append(get_color_name(input_row))
    row.append(get_type(commercial_type))
    row.extend(['', '', ''])
    row.append(get_description(input_row))


def get_inventory_for_home_row(input_row, commercial_type):
    """ Генерирует строку для записи в Дом и сад > Инвентарь для дома """
    row = []
    get_basic_row(row, input_row, commercial_type)
    row.extend(["", ""])
    row.append(get_brand(input_row))
    row.append(get_title(input_row))
    row.append("")
    row.append(get_color_name(input_row))
    row.append(get_type(commercial_type))
    row.extend(['', '', ''])
    row.append(get_description(input_row))


def get_inventory_for_cleaning_row(input_row, commercial_type):
    """ Генерирует строку для записи в Дом и сад > Инвентарь для уборки """
    row = []
    get_basic_row(row, input_row, commercial_type)
    row.extend(["", ""])
    row.append(get_brand(input_row))
    row.append(get_title(input_row))
    row.append("")
    row.append(get_color_name(input_row))
    row.append(get_type(commercial_type))
    row.extend(['', '', ''])
    row.append(get_description(input_row))


def get_disposable_tableware_row(input_row, commercial_type):
    """ Генерирует строку для записи в Дом и сад > Одноразовая посуда """
    row = []
    get_basic_row(row, input_row, commercial_type)
    row.extend(["", ""])
    row.append(get_brand(input_row))
    row.append(get_title(input_row))
    row.append("")
    row.append(get_color_name(input_row))
    row.append(get_type(commercial_type))
    row.extend(['', '', ''])
    row.append(get_description(input_row))


def get_dishes_row(input_row, commercial_type):
    """ Генерирует строку для записи в Дом и сад > Столовая посуда """
    row = []
    get_basic_row(row, input_row, commercial_type)
    row.extend(["", ""])
    row.append(get_brand(input_row))
    row.append(get_title(input_row))
    row.append("")
    row.append(get_color_name(input_row))
    row.append("")
    row.append(get_type(commercial_type))
    row.extend(['', '', ''])
    row.append(get_description(input_row))


def get_things_storage_row(input_row, commercial_type):
    """ Генерирует строку для записи в Дом и сад > Хранение вещей """
    row = []
    get_basic_row(row, input_row, commercial_type)
    row.extend(["", ""])
    row.append(get_brand(input_row))
    row.append(get_title(input_row))
    row.extend(["", "", "", ""])
    row.append(get_type(commercial_type))
    row.extend(['', '', ''])
    row.append(get_description(input_row))


def get_paper_row(input_row, commercial_type):
    """ Генерирует строку для записи в Канцелярские товары > Бумага """
    row = []
    get_basic_row(row, input_row, commercial_type)
    row.extend(["", ""])
    row.append(get_brand(input_row))
    row.append(get_title(input_row))
    row.append("")
    row.append(get_sheets_count(input_row))
    row.append("")
    row.append(get_color_name(input_row))
    row.append(get_type(commercial_type))
    row.append("")
    row.append("")
    row.extend(['', '', '', ''])
    row.append(get_description(input_row))


def get_demonstration_boards_row(input_row, commercial_type):
    """ Генерирует строку для записи в Канцелярские товары > Демонстрационные доски """
    row = []
    get_basic_row(row, input_row, commercial_type)
    row.extend(["", ""])
    row.append(get_brand(input_row))
    row.append(get_title(input_row))
    row.append("")
    row.append(get_color_name(input_row))
    row.append(get_type(commercial_type))
    row.append("")
    row.extend(['', '', '', ''])
    row.append(get_description(input_row))


def get_child_bags_row(input_row, commercial_type):
    """ Генерирует строку для записи в Канцелярские товары > Детские рюкзаки, ранцы, сумки """
    row = []
    get_basic_row(row, input_row, commercial_type)
    row.extend(["", ""])
    row.append(get_brand(input_row))
    row.append(get_title(input_row))
    row.append("")
    row.append(get_color_name(input_row))
    row.append(get_type(commercial_type))
    row.append("")
    row.append("")
    row.extend(['', '', '', '', '', '', ''])
    row.append(get_description(input_row))


def get_glue_row(input_row, commercial_type):
    """ Генерирует строку для записи в Канцелярские товары > Краска, клей """
    row = []
    get_basic_row(row, input_row, commercial_type)
    row.extend(["", ""])
    row.append(get_brand(input_row))
    row.append(get_title(input_row))
    row.append("")
    row.append(get_color_name(input_row))
    row.append("")
    row.append("")
    row.append(get_expiration_date())  # срок годности
    row.append(get_type(commercial_type))
    row.append("")
    row.append("")
    row.extend(['', '', ''])
    row.append(get_description(input_row))


def get_pencil_box_row(input_row, commercial_type):
    """ Генерирует строку для записи в Канцелярские товары > Пенал """
    row = []
    get_basic_row(row, input_row, commercial_type)
    row.extend(["", ""])
    row.append(get_brand(input_row))
    row.append(get_title(input_row))
    row.append("")
    row.append(get_color_name(input_row))
    row.append("")
    row.append("")
    row.append(get_type(commercial_type))
    row.append("Каркасный пенал")  # вид пенала ,
    row.append("1")  # кол-во отделений
    row.append("Нет")  # наполнение пенала
    row.append("")
    row.append("")
    row.extend(['', '', '', '', ''])
    row.append(get_description(input_row))


def get_seal_and_stamp_row(input_row, commercial_type):
    """ Генерирует строку для записи в Канцелярские товары > Печати и штампы """
    row = []
    get_basic_row(row, input_row, commercial_type)
    row.extend(["", ""])
    row.append(get_brand(input_row))
    row.append(get_title(input_row))
    row.append("")
    row.append(get_color_name(input_row))
    row.append(get_type(commercial_type))
    row.append("")
    row.append("")
    row.append('')
    row.extend(['', '', ''])
    row.append(get_description(input_row))


def get_writing_materials_row(input_row, commercial_type):
    """ Генерирует строку для записи в Канцелярские товары > Письменные принадлежности """
    row = []
    get_basic_row(row, input_row, commercial_type)
    row.extend(["", ""])
    row.append(get_brand(input_row))
    row.append(get_title(input_row))
    row.append("")
    row.append(get_color_name(input_row))
    row.append("")
    row.append(get_type(commercial_type))
    row.append("")
    row.append("")
    row.extend(['', '', '', '', ''])
    row.append(get_description(input_row))


def get_personal_hygiene_row(input_row, commercial_type):
    """ Генерирует строку для записи в Красота и Гигиена > Товары личной гигиены """
    row = []
    get_basic_row(row, input_row, commercial_type)
    row.extend(["", ""])
    row.append(get_title(input_row))
    row.append(get_brand(input_row))
    row.append("")
    row.append("")
    row.append("")
    row.append(get_color_name(input_row))
    row.append(get_type(commercial_type))
    row.append(get_expiration_date())  # срок годности
    row.extend(['', ''])
    row.append(get_description(input_row))
    row.extend(['', '', ''])
    row.append("")


def get_clothes_row(input_row, commercial_type):
    """ Генерирует строку для записи в Одежда > Одежда """
    row = []
    get_basic_row(row, input_row, commercial_type)
    row.extend(["", ""])
    row.append(get_brand(input_row))
    row.append(get_title(input_row))
    row.append("")
    row.append("универсальный")
    row.append(get_size(input_row))
    row.append(get_color_name(input_row))
    row.append(get_type(commercial_type))
    row.append("Мужской")
    row.extend(['', '', '', '', '', ''])
    row.append("")
    row.append('')
    row.append(get_description(input_row))


def get_juse_drinks_row(input_row, commercial_type):
    """ Генерирует строку для записи в Продукты питания > Соки, воды, напитки """
    row = []
    get_basic_row(row, input_row, commercial_type)
    row.extend(["", ""])
    row.append(get_brand(input_row))
    row.append(get_title(input_row))
    row.append("count_in_pack")
    row.append("weight")
    row.append(get_type(commercial_type))
    row.append("minimal_temp")
    row.append("maximum_temp")
    row.append(get_expiration_date())  # срок годности
    row.append("storage_conditions")  # условия хранения как на упаковке
    row.append("structure")
    row.extend(['', ''])
    row.append(get_description(input_row))
    row.extend(['', '', '', '', '', '', '', '', '', ''])
    row.append("")


def get_bread_row(input_row, commercial_type):
    """ Генерирует строку для записи в  Продукты питания > Хлеб и кондитерские изделия """
    row = []
    get_basic_row(row, input_row, commercial_type)
    row.extend(["", ""])
    row.append(get_brand(input_row))
    row.append(get_title(input_row))
    row.append("weight")
    row.append("count_in_pack")
    row.append(get_type(commercial_type))
    row.append("minimal_temp")
    row.append("maximum_temp")
    row.append(get_expiration_date())  # срок годности
    row.append("storage_conditions")  # условия хранения как на упаковке
    row.append("structure")
    row.extend(['', '', '', ''])
    row.append(get_description(input_row))
    row.extend(['', '', '', '', '', '', '', '', '', '', '', ''])
    row.append("")


def get_lamp_row(input_row, commercial_type):
    """ Генерирует строку для записи в Строительство и ремонт > Лампочка """
    row = []
    get_basic_row(row, input_row, commercial_type)
    row.extend(["", ""])
    row.append(get_brand(input_row))
    row.append(get_title(input_row))
    row.append("")
    row.append("count_in_pack")
    row.append("cokol_lamp")  # тип цоколя
    row.append("power")  # мощность, ВТ
    row.append(get_type(commercial_type))
    row.extend(['', '', ''])
    row.append("")
    row.extend(['', '', '', '', '', ''])
    row.append(get_description(input_row))


def get_big_lamp_row(input_row, commercial_type):
    """ Генерирует строку для записи в Строительство и ремонт > Светильник """
    row = []
    get_basic_row(row, input_row, commercial_type)
    row.extend(["", ""])
    row.append(get_brand(input_row))
    row.append(get_title(input_row))
    row.extend(["", "", ""])
    row.append(get_type(commercial_type))
    row.extend(['', '', '', '', '', '', '', ''])
    row.append(get_description(input_row))
    row.extend(['', ''])
    row.append("")


def get_fire_fighting_row(input_row, commercial_type):
    """ Генерирует строку для записи в Строительство и ремонт > Средства защиты и пожаротушения """
    row = []
    get_basic_row(row, input_row, commercial_type)
    row.extend(["", ""])
    row.append(get_brand(input_row))
    row.append(get_title(input_row))
    row.extend(["", ""])
    row.append(get_type(commercial_type))
    row.extend(['', ''])
    row.append("")
    row.extend(['', '', '', ''])
    row.append(get_description(input_row))


def get_children_creativity_row(input_row, commercial_type):
    """ Генерирует строку для записи в Хобби и творчество > Детское творчество и развитие """
    row = []
    get_basic_row(row, input_row, commercial_type)
    row.extend(["", ""])
    row.append(get_brand(input_row))
    row.append(get_title(input_row))
    row.extend(["", ""])
    row.append(get_type(commercial_type))
    row.extend([''])
    row.append("")
    row.extend(['', '', '', '', '', ''])
    row.append(get_description(input_row))


def get_set_for_creativity_row(input_row, commercial_type):
    """ Генерирует строку для записи в  Хобби и творчество > Набор для рукоделия, творчества """
    row = []
    get_basic_row(row, input_row, commercial_type)
    row.extend(["", ""])
    row.append(get_brand(input_row))
    row.append(get_title(input_row))
    row.extend(["", ""])
    row.append(get_type(commercial_type))
    row.extend([''])
    row.append("")
    row.extend(['', '', '', '', '', ''])
    row.append(get_description(input_row))


def get_bag_row(input_row, commercial_type):
    """ Генерирует строку для записи в  Электроника > Рюкзаки, чехлы, сумки """
    row = []
    get_basic_row(row, input_row, commercial_type)
    row.extend(["", ""])
    row.append(get_brand(input_row))
    row.append(get_title(input_row))
    row.append("")
    row.append(get_type(commercial_type))
    row.extend([''])
    row.append("")
    row.extend(['', ''])
    row.append(get_description(input_row))


def get_first_dict_value(dict):
    if type(dict) == str:
        return 1
    for i in dict:
        if get_first_dict_value(dict[i]) == 1:
            return dict[i]


def get_ozon_category(initial_category: str):
    categories_match_file = open("categories.json", 'r', encoding='utf-8')
    categories_match_dict = json.load(categories_match_file)
    categories_match_file.close()
    initial_category = initial_category.split("/")
    our_category = categories_match_dict
    for category in initial_category:
        if category in our_category:
            our_category = our_category[category]
        else:
            return False
    if type(our_category) == str:
        return our_category
    elif type(our_category) == dict:
        return get_first_dict_value(our_category)


def write_row(filename, row):
    wb = openpyxl.load_workbook(filename)
    ws = wb.active
    ws.append(row)
    wb.save(filename)


def main():
    wb = openpyxl.load_workbook("input.xlsx")
    # Получаем активный лист
    sheet = wb.active
    # Читаем и выводим каждую строку
    chancellery_categories = [i.strip() for i in open("chancellery.txt", 'r', encoding='utf-8').readlines()]
    paper_products_categories = [i.strip() for i in
                                 open("paper_products_categories.txt", 'r', encoding='utf-8').readlines()]
    for input_row in sheet.iter_rows(values_only=True):
        print(input_row)
        ozon_category = get_ozon_category(input_row[7])
        if ozon_category in chancellery_categories:
            row = get_stationery_row(input_row, ozon_category)
            write_row("chancellery_output.xlsx", row)
        if ozon_category in paper_products_categories:
            row = get_paper_products_row(input_row, ozon_category)
            write_row("paper_products_output.xlsx", row)


if __name__ == "__main__":
    main()
