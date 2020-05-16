import random
import re

import yaml
import os
from datetime import datetime

from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.shared import Pt
from openpyxl import Workbook
from openpyxl.styles import PatternFill

AUDS = {
    'В-301': 0,
    'В-302': 0,
    'В-304': 0,
    'В-201': 0,
    'В-202': 0,
    'В-214': 0,
    'К-207': 0,
    'Г-102': 0,
    'Г-105': 0,
    'Г-110': 0,
    'Г-110б': 0,
    'Г-206': 0,
    'Г-212': 0,
    'Г-219': 0,
    'Г-314': 0,
    'И-305': 0,
    'И-307': 0,
    'И-307a': 0,
    'И-309': 0,
    'И-311': 0,
    'И-315': 0,
    'И-317': 0,
    'И-319': 0,
    'K301': 0,
    'K302': 0,
    'K303': 0,
    'K305': 0,
    'K313': 0,
    'K315а': 0,
    'В-02': 1,
    'Ж-216': 1,
    'И-202': 1,
    'K311': 2,
    'K315': 2,
    'K317': 2,
    'Ж-319': 2,
    'Ж-321': 2,
    'Ж-323': 2,
    'Ж-амф.': 3,
    'Г-амф.': 3,
}


def process_file(document: Document):
    data = []
    table = document.tables[0]
    for row in table.rows[1:]:
        cells = list(row.cells)
        if cells[2].text:
            dt = datetime.strptime(cells[2].text.strip() + " " + cells[3].text.strip(), "%d.%m.%Y %H.%M")

            data.append({
                'title': cells[0].text,
                'type': cells[1].text,
                'date': dt,
            })
        else:
            print(f"Остутсвует дата для {cells[0].text} в файле {document}")
    return data


def fill_data_file():
    """
    Формирует data.yaml из docx документов в папке ./data/2020
    :return:
    """
    data = {}
    for file in os.listdir("./data"):
        if file.endswith(".docx"):
            title = file.replace('.docx', '')
            data[title] = process_file(Document(os.path.join("./data", file)))

    with open("data.yaml", "w", encoding="utf8") as f:
        yaml.dump(data, f, allow_unicode=True)


def set_cell_value(cell, value, highlight=False):
    if len(cell.paragraphs) == 0:
        cell.add_paragraph()
    par = cell.paragraphs[0]

    if len(par.runs) == 0:
        par.add_run()

    par.runs[0].font.size = Pt(12)
    if highlight:
        par.runs[0].font.highlight_color = highlight
    par.runs[0].text = value


def calculate_schedule():
    """
    Рассчет расписания по data.yaml
    :return:
    """
    with open("data.yaml", 'r', encoding="utf8") as f:
        data = yaml.load(f)

    random.seed(1)

    # словарик под расписание аудтиорий
    auds_schedule = {
        aud: {}
        for aud, v in AUDS.items()
    }

    # тусуем аудитории чтобы усреднить результат
    all_aud_ids = list(AUDS.keys())
    random.shuffle(all_aud_ids)

    # расстояние между занятиями
    gap = 5

    # тусуем занятия для проектирования
    all_items_to_schedule = []
    for key, items in data.items():
        all_items_to_schedule.extend(items)
    random.shuffle(all_items_to_schedule)

    # проектируем
    for item in all_items_to_schedule:
        found_aud = False

        # расставляем занятия по приоритетами
        for priority in range(max(AUDS.values()) + 1):
            # формируем список аудиторий с указаным приоритетам
            aud_ids = [k for k in all_aud_ids if AUDS[k] == priority]
            # ищем свбодные аудитории
            for aud in aud_ids:
                dates = auds_schedule[aud]
                date_items = dates.setdefault(item['date'].date(), [])
                # если нашли, то фиксируем
                if not date_items:
                    date_items.append(item)
                    item['aud'] = aud
                    found_aud = True
                    break

            # если не нашли, пытаемся запихнуть занятия второй сменной
            if not found_aud:
                for aud in aud_ids:
                    dates = auds_schedule[aud]
                    date_items = dates.setdefault(item['date'].date(), [])

                    if all(abs((item['date'] - i['date']).total_seconds()) / 3600 >= gap for i in date_items):
                        date_items.append(item)
                        item['aud'] = aud
                        found_aud = True
                        break
            # если нашли занятия то нет смысла пытатся влезть в аудитории с более низким приоритетом
            if found_aud:
                break

        if not found_aud:
            print(f"Не удалось найти аудитории для {item['title']}")

    return data


def create_docx_documents(data):
    """
    Формирование docx документов по составленому расписанию
    :param data:
    :return:
    """
    for key, items in data.items():
        doc = Document("template.docx")
        table = doc.tables[0]

        previous_title = None
        previous_row_index = 1
        row_index = 0

        for row_index, item in enumerate(items, 1):
            row = table.add_row()

            cells = list(row.cells)

            title = item['title'].strip()
            if title != previous_title:
                if previous_title:
                    cell_start = table.rows[previous_row_index].cells[0]
                    cell_end = table.rows[row_index - 1].cells[0]
                    cell_start.merge(cell_end)
                    set_cell_value(cell_start, previous_title)

                    previous_row_index = row_index

                previous_title = title

            for index, value in enumerate([
                item['type'].strip(),
                "{:%d.%m.%Y}".format(item['date']),
                "{:%H.%M}".format(item['date']),
                item['aud'],
            ], 1):
                set_cell_value(cells[index], value, {
                    1: WD_COLOR_INDEX.BRIGHT_GREEN,
                    2: WD_COLOR_INDEX.YELLOW,
                    3: WD_COLOR_INDEX.RED,
                }.get(AUDS.get(item['aud'])))

            row_index += 1

        if previous_row_index != row_index:
            cell_start = table.rows[previous_row_index].cells[0]
            cell_end = table.rows[row_index - 1].cells[0]
            cell_start.merge(cell_end)
            set_cell_value(cell_start, previous_title)

        doc.save(os.path.join("output", f"{key}.docx"))


def generate_auds_schedule_document(data):
    """
    Формирование сетки аудиторий в формате дата/аудитория
    :param data:
    :return:
    """
    result = {}

    for key, items in data.items():
        for item in items:
            date_item = result.setdefault(item['date'].date(), {})
            all_items = date_item.setdefault(item['aud'], [])
            all_items.append(item)

    wb = Workbook()
    ws = wb.active

    auds = sorted(AUDS.keys(), key=lambda x: x.replace("K", "К").replace("B", "В"))

    for date_index, date in enumerate(sorted(result.keys())):
        ws.cell(date_index + 2, 1, "{:%d.%m.%Y}".format(date))
        for aud_index, aud in enumerate(auds):
            ws.cell(1, aud_index + 2, f"{aud}")
            items = result.get(date, {}).get(aud, [])
            if items:
                title = "\n\n".join(f"{'{:%H.%M}'.format(i['date'])}: {i['title']}" for i in sorted(items, key=lambda x: x['date']))
                cell = ws.cell(date_index + 2, aud_index + 2, f"{title}")
                cell.fill = PatternFill(fill_type='solid', start_color="FFDDDDDD", end_color="FFDDDDDD")

    wb.save("auds.xlsx")


def main():
    # fill_data_file()
    # data = calculate_schedule()
    # with open("processed.yaml", "w", encoding='utf8') as f:
    #     yaml.dump(data, f, allow_unicode=True)

    with open("processed.yaml", encoding='utf8') as f:
        data = yaml.load(f)

    # create_docx_documents(data)
    generate_auds_schedule_document(data)


if __name__ == '__main__':
    main()
