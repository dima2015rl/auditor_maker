import re

import xlrd
import xlwt
import subprocess
import json
import flet as ft
import sys


def lessons_shift(groups_dict):
    for key in list(groups_dict.keys()):
        while len(groups_dict[key]) > 4:
            groups_dict[key].popitem()
        group_keys = list(groups_dict[key].keys())
        for number in group_keys:
            if number > 4:
                index = group_keys.index(number)
                prev_index = index - 1
                len_from_index = len(group_keys[index:])
                if len_from_index <= 4 and prev_index < 0 \
                        or len_from_index < 4 < len_from_index + group_keys[prev_index] and prev_index >= 0:
                    for p, value in enumerate(list(groups_dict[key])):
                        groups_dict[key][value][0] = f"{value} пара\n" + groups_dict[key][value][0] if groups_dict[key][value][0] != "" \
                            else f"{value} пара"
                        groups_dict[key][p + 1] = groups_dict[key].pop(value)
                        group_keys[index] = p + 1
                    break
                else:
                    groups_dict[key][number][0] = f"{number} пара\n" + groups_dict[key][number][0] if \
                        groups_dict[key][number][0] != "" \
                        else f"{number} пара"
                    groups_dict[key][group_keys[prev_index] + 1] = groups_dict[key].pop(number)
                    group_keys[index] = group_keys[prev_index] + 1
    return groups_dict


def isint(num):
    try:
        int(num)
        return True
    except ValueError:
        return False


def app(page: ft.Page):
    def read(e, file):
        try:
            rb = xlrd.open_workbook(file, formatting_info=True)
            sheet = rb.sheet_by_index(0)
            year_group_dict = {}
            data_lst = []
            groupe_lst = []
            for row_index in range(1, sheet.nrows):
                groupe = sheet.cell_value(row_index, 0)
                number = int(sheet.cell_value(row_index, 1)) if isint(sheet.cell_value(row_index, 1)) is True else 1
                time = sheet.cell_value(row_index, 2)
                quest = sheet.cell_value(row_index, 3)
                teacher = sheet.cell_value(row_index, 4)
                audit = sheet.cell_value(row_index, 5)
                if isinstance(audit, float):
                    audit = str(int(audit))
                full_name = teacher.split(' ')
                teacher_shorted_name = full_name[0] + ' ' + full_name[1][0] + '.' \
                                       + full_name[2][0] + '.' if len(full_name) >= 3 else ""

                practic_name = ''
                if quest.upper().find('ПРАКТИКА') != -1:
                    practic_name = quest[0:quest.find(' ')]
                    if practic_name.upper() == "ПРОИЗВОДСТВЕННАЯ":
                        practic_name = "ПП"
                    elif practic_name.upper() == "УЧЕБНАЯ":
                        practic_name = "УП"

                practic_name = practic_name if quest.upper().find('НЕТ ЗАНЯТИЙ') == -1 else "нет занятий"
                text = f'{groupe}***{number}***{time}***{quest}***{teacher_shorted_name}***{audit}'
                data_lst.append(text)
                groupe_lst.append(groupe)
                group_l = groupe.split('-')
                year=''
                for ch in group_l:
                    if isint(ch):
                        if int(ch) > 10:
                            year = int(ch)

                if year not in year_group_dict:
                    year_group_dict[year] = {}
                if groupe not in year_group_dict[year]:
                    year_group_dict[year][groupe] = {}
                if number in year_group_dict[year][groupe]:
                    original_list = year_group_dict[year][groupe][number]
                    original_list[1] = f"{original_list[1]}\n{teacher_shorted_name}"
                    if original_list[2] != audit:
                        original_list[2] = original_list[2] + "/" + audit
                else:
                    year_group_dict[year][groupe][number] = [practic_name, teacher_shorted_name, audit]
            return groupe_lst, data_lst, year_group_dict
        except Exception as ex:
            print(f'Error in function READ() - {ex}')

    def write(e, data, group, out, row_h, row_h_s, col_w, col_w_s, f1s, f2s, table, table2, border, ft1, ft2, ft1_b,
              ft2_b,
              year_groups_dict):
        try:
            if data and group:
                wb = xlwt.Workbook()
                ws = wb.add_sheet(table, cell_overwrite_ok=True)
                ws2 = wb.add_sheet(table2, cell_overwrite_ok=True)

                # Выставление ширины столбцов
                ws.col(0).width = 256 * int(col_w_s)
                for i in range(1, 10):
                    ws.col(i).width = 256 * int(col_w)
                ws2.col(0).width = 256 * int(col_w_s)
                for i in range(1, 10):
                    ws2.col(i).width = 256 * int(col_w)

                # Создание центрирования
                alignment = xlwt.Alignment()
                alignment.wrap = 1
                alignment.horz = xlwt.Alignment.HORZ_CENTER
                alignment.vert = xlwt.Alignment.VERT_CENTER

                # Создание границ
                borders = xlwt.Borders()
                borders.left = xlwt.Borders.MEDIUM
                borders.right = xlwt.Borders.MEDIUM
                borders.top = xlwt.Borders.MEDIUM
                borders.bottom = xlwt.Borders.MEDIUM

                # Создание шрифтов
                font0 = xlwt.Font()
                font0.name = ft1
                font0.height = int(f1s) * 20
                font0.bold = ft1_b

                font1 = xlwt.Font()
                font1.name = ft2
                font1.height = int(f2s) * 20
                font1.bold = ft2_b

                # Создание стилей
                style0 = xlwt.XFStyle()
                style0.alignment = alignment
                style0.font = font0
                style0.font.bold = True

                style1 = xlwt.XFStyle()
                style1.alignment = alignment
                style1.font = font1

                # Добавление границ
                if border:
                    style0.borders = borders
                    style1.borders = borders
                # шабка таблицы границ
                ws.write(1, 0, 'Группа', style0)
                ws2.write(1, 0, 'Группа', style0)
                ws.write(1, 5, 'Группа', style0)
                ws2.write(1, 5, 'Группа', style0)

                for i in range(1, 5):
                    ws.write(1, i, f"{i} Пара", style0)
                    ws.write(1, i + 5, f"{i} Пара", style0)
                    ws2.write(1, i, f"{i} Пара", style0)
                    ws2.write(1, i + 5, f"{i} Пара", style0)
                    ws.write(0, i, f"", style0)
                    ws.write(0, i + 5, f"", style0)
                    ws2.write(0, i, f"", style0)
                    ws2.write(0, i + 5, f"", style0)
                for i in range(0, 2):
                    ws.row(i).height_mismatch = True
                    ws.row(i).height = 256 * int(row_h_s)
                    ws2.row(i).height_mismatch = True
                    ws2.row(i).height = 256 * int(row_h_s)

                ctr = 0
                current_year = int(max(list(year_groups_dict.keys())))
                print(current_year)
                years= list(year_groups_dict.keys())
                years.sort()
                years.reverse()
                print(year_groups_dict)
                for year in years:
                    year_groups_dict[year] = lessons_shift(year_groups_dict[year])
                    if year == current_year:
                        for i, group in enumerate(year_groups_dict[year]):
                            print(group)
                            ws.write(i + 2, 0, group, style0)
                            for j in range(1, 5):
                                text = ""
                                if j in year_groups_dict[year][group]:
                                    desk = list(filter(lambda a: a != "", year_groups_dict[year][group][j]))
                                    text = "\n".join(desk)
                                ws.write(i + 2, j, text, style1)
                    elif year == current_year - 1:
                        for i, group in enumerate(year_groups_dict[year]):
                            ws.write(i + 2, 5, group, style0)
                            for j in range(1, 5):
                                text = ""
                                if j in year_groups_dict[year][group]:
                                    desk = list(filter(lambda a: a != "", year_groups_dict[year][group][j]))
                                    text = "\n".join(desk)
                                ws.write(i + 2, j + 5, text, style1)
                    elif year == current_year - 2:
                        for i, group in enumerate(year_groups_dict[year]):
                            ws2.write(i + 2, 0, group, style0)
                            for j in range(1, 5):
                                text = ""
                                if j in year_groups_dict[year][group]:
                                    desk = list(filter(lambda a: a != "", year_groups_dict[year][group][j]))
                                    text = "\n".join(desk)
                                ws2.write(i + 2, j, text, style1)
                    elif year == current_year - 3:
                        for i, group in enumerate(year_groups_dict[year]):
                            ws2.write(i + 2, 5, group, style0)
                            for j in range(1, 5):
                                text = ""
                                if j in year_groups_dict[year][group]:
                                    desk = list(filter(lambda a: a != "", year_groups_dict[year][group][j]))
                                    text = "\n".join(desk)
                                ws2.write(i + 2, j + 5, text, style1)
                            ctr += 1
                    else:
                        for i, group in enumerate(year_groups_dict[year]):
                            ws2.write(i + ctr + 2, 5, group, style0)
                            for j in range(1, 5):
                                text = ""
                                if j in year_groups_dict[year][group]:
                                    desk = list(filter(lambda a: a != "", year_groups_dict[year][group][j]))
                                    text = "\n".join(desk)
                                ws2.write(i + ctr + 2, j + 5, text, style1)
                wb.save(out)
            else:
                print('Error in function WRITE() - Значения не переданы')
        except Exception as ex:
            print(f"`Error1` - {ex}")

    def start(e):
        try:
            try:
                CONFIG_PATH = 'config.json'
                with open(CONFIG_PATH, 'r') as file:
                    data = json.load(file)
            except Exception as ex:
                print(f"Error1 - {ex}")

            if not data:
                FILE = 'file.xls'
                OUT_FILE = 'output.xls'
                ROW_HEIGHT = 14
                ROW_HEIGHT_SPECIAL = 6
                COLUMN_WIDTH = 50
                COLUMN_WIDTH_SPECIAL = 30
                BORDER = True
                TABLE = 'Расписание'
                FONT1_SIZE = 24
                FONT2_SIZE = 16
                FONT1 = 'Times New Roman'
                FONT2 = 'Times New Roman'
                FONT1_BOLD = True
                FONT2_BOLD = False
                START_WHERE_END = False
            else:
                FILE = str(data["FILE"])
                OUT_FILE = str(data['OUT_FILE'])
                ROW_HEIGHT = int(data['ROW_HEIGHT'])
                ROW_HEIGHT_SPECIAL = int(data['ROW_HEIGHT_SPECIAL'])
                COLUMN_WIDTH = int(data["COLUMN_WIDTH"])
                COLUMN_WIDTH_SPECIAL = int(data["COLUMN_WIDTH_SPECIAL"])
                BORDER = bool(data["BORDER"])
                TABLE = str(data["TABLE"])
                TABLE2 = str(data["TABLE2"])
                FONT1_SIZE = int(data["FONT1_SIZE"])
                FONT2_SIZE = int(data["FONT2_SIZE"])
                FONT1 = str(data['FONT1'])
                FONT2 = str(data['FONT2'])
                FONT1_BOLD = bool(data['FONT1_BOLD'])
                FONT2_BOLD = bool(data['FONT2_BOLD'])
                START_WHERE_END = bool(data['START_WHERE_END'])

            FILE = file_in_button.value if file_in_button.value != '' else FILE
            OUT_FILE = file_out_button.value if file_out_button.value != '' else OUT_FILE

            ROW_HEIGHT = row_button.value if row_button.value != '' else ROW_HEIGHT
            ROW_HEIGHT_SPECIAL = row_special_button.value if row_special_button.value != '' else ROW_HEIGHT_SPECIAL

            COLUMN_WIDTH = column_button.value if column_button.value != '' else COLUMN_WIDTH
            COLUMN_WIDTH_SPECIAL = column_special_button.value if column_special_button.value != '' else COLUMN_WIDTH_SPECIAL

            FONT1 = font1_button.value if font1_button.value != '' else FONT1
            FONT2 = font2_button.value if font2_button.value != '' else FONT2

            FONT1_SIZE = font1_size_button.value if font1_size_button.value != '' else FONT1_SIZE
            FONT2_SIZE = font2_size_button.value if font2_size_button.value != '' else FONT2_SIZE

            TABLE = table_button.value if table_button.value != '' else TABLE
            TABLE2 = table_button2.value if table_button2.value != '' else TABLE2

            FONT1_BOLD = font1_bold_button.value
            FONT2_BOLD = font2_bold_button.value

            BORDER = border_button.value
            START_WHERE_END = start_excel_button.value

            groups, data, year_groups_dict = read(e, file=FILE)
            # Запись данных
            write(e, data=data, group=groups, out=OUT_FILE, row_h=ROW_HEIGHT, row_h_s=ROW_HEIGHT_SPECIAL,
                  col_w=COLUMN_WIDTH, col_w_s=COLUMN_WIDTH_SPECIAL, table=TABLE, table2=TABLE2, border=BORDER,
                  ft1=FONT1, f1s=FONT1_SIZE, ft1_b=FONT1_BOLD, ft2=FONT2, f2s=FONT2_SIZE, ft2_b=FONT2_BOLD,
                  year_groups_dict=year_groups_dict)

            if START_WHERE_END:
                subprocess.Popen(['start', 'excel', OUT_FILE], shell=True)
        except Exception as ex:
            print(f"Error - {ex}")

    page.window_height = 640
    page.window_width = 720
    page.title = 'Автоматизация расписания'
    page.window_resizable = False

    file_in_button = ft.TextField(label='Входной файл', width=620, height=60)
    file_out_button = ft.TextField(label='Выходной файл', width=620, height=60)

    font1_button = ft.TextField(label='Шрифт заголовков', width=300, height=60)
    font2_button = ft.TextField(label='Шрифт основной', width=300, height=60)

    font1_size_button = ft.TextField(label='Размер шрифта заголовков', width=300, height=60)
    font2_size_button = ft.TextField(label='Размер основного шрифта', width=300, height=60)

    font1_bold_button = ft.Checkbox(label="Жирный шрифт заголовков", value=False)
    font2_bold_button = ft.Checkbox(label="Жирный основной шрифт", value=False)

    row_button = ft.TextField(label='Высота строки', width=300, height=60)
    row_special_button = ft.TextField(label='Высота строки наименований', width=300, height=60)

    column_button = ft.TextField(label='Ширина колонны', width=300, height=60)
    column_special_button = ft.TextField(label='Ширина колонны заголовков', width=300, height=60)

    border_button = ft.Checkbox(label="Границы у ячеек", value=False)
    start_excel_button = ft.Checkbox(label="Запустить Excel по окончанию", value=False)

    table_button = ft.TextField(label='Название листа1', width=300, height=60)
    table_button2 = ft.TextField(label='Название листа2', width=350, height=60)

    page.add(
        ft.Row([file_in_button], alignment=ft.MainAxisAlignment.CENTER),
        ft.Row([file_out_button], alignment=ft.MainAxisAlignment.CENTER),

        ft.Row([font1_button, ft.Container(width=0), row_button], alignment=ft.MainAxisAlignment.CENTER),
        ft.Row([font2_button, ft.Container(width=0), row_special_button], alignment=ft.MainAxisAlignment.CENTER),

        ft.Row([font1_size_button, ft.Container(width=0), column_button], alignment=ft.MainAxisAlignment.CENTER),
        ft.Row([font2_size_button, ft.Container(width=0), column_special_button],
               alignment=ft.MainAxisAlignment.CENTER),

        ft.Row([ft.Container(width=20), font1_bold_button,
                ft.Container(width=85), border_button], alignment=ft.MainAxisAlignment.START),
        ft.Row([ft.Container(width=20), font2_bold_button,
                ft.Container(width=95), start_excel_button], alignment=ft.MainAxisAlignment.START),
        ft.Row([ft.ElevatedButton("Начать", on_click=start, width=300, height=60), ft.Container(width=0),
                table_button], alignment=ft.MainAxisAlignment.CENTER)
    )


if __name__ == '__main__':
    try:
        ft.app(target=app)
    except Exception as ex:
        print(f"Error - {ex}")
    finally:
        sys.exit()
