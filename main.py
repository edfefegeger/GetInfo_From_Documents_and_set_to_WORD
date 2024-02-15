import fitz
import easyocr
from docx import Document
import re
import torch
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

def process_pdf(pdf_path, keywords, word_path):
    reader = easyocr.Reader(['en', 'ru'], gpu=True)

    # Поиск ключевых слов и даты
    found_keywords = []
    found_date = None  # Здесь будем хранить найденную дату

    start_page = 1  # Начальная страница текущего документа

    with fitz.open(pdf_path) as pdf:
        for page_num in range(len(pdf)):
            page = pdf.load_page(page_num)
            text = page.get_text()

            print(f"Обрабатывается страница {page_num + 1}...")

            # Если на странице есть текст или изображения, обрабатываем ее
            images = page.get_images(full=True)
            if text or images:

                # Поиск текста в изображениях
                for img_index, img in enumerate(images):
                    xref = img[0]
                    base_image = pdf.extract_image(xref)
                    image_bytes = base_image["image"]
                    result = reader.readtext(image_bytes)
                    for detection in result:
                        img_text = detection[1]
                        text += " " + img_text  # Добавляем текст изображения к тексту страницы

                # Поиск ключевых слов
                for keyword in keywords:
                    if keyword in text:
                        print(f"Ключевое слово '{keyword}' найдено")
                        found_keywords.append(keyword)
                if not found_date:  # Проверяем, была ли найдена дата ранее
                    date = find_dates(text)
                    if date:
                        print("Дата найдена:", date)
                        found_date = date 

                if len(pdf) == 1:  # Если документ содержит только одну страницу
                    print("Документ содержит только одну страницу. Завершение обработки.")
                    end_page = 1  # Конечная страница текущего документа
                    update_word_table(word_path, keywords, found_keywords, found_date, start_page, end_page)
                    found_keywords = []
                    found_date = None
                else:
                    if "End" in text:  
                        print(f"Найдена пометка 'End' на странице {page_num + 1}. Завершение документа.")
                        end_page = page_num + 1  # Конечная страница текущего документа
                        update_word_table(word_path, keywords, found_keywords, found_date, start_page, end_page)
                        found_keywords = []
                        found_date = None
                        start_page = page_num + 2  # Начальная страница следующего документа      
                
        # Записываем информацию в файл Word после окончания обработки документа

    return found_keywords, found_date




def update_word_table(word_path, keywords, found_keywords, found_date, start_page, end_page):
    doc = Document(word_path)
    table = doc.tables[0]

    # Находим индекс столбца "№ з/п"
    for cell in table.rows[0].cells:
        if cell.text.strip() == "№ з/п":
            num_index = cell._element.getparent().index(cell._element) - 2
            break  

    # Находим индекс столбца "Наименование документа"
    for cell in table.rows[0].cells:
        if cell.text.strip() == "Наименование документа":
            column_index = cell._element.getparent().index(cell._element)
            break
    for cell in table.rows[0].cells:
        if cell.text.strip() == "Номера листов":
            list_index = cell._element.getparent().index(cell._element)
            break   
    for cell in table.rows[0].cells:
        if cell.text.strip() == "исходящие":
            incoming_index = cell._element.getparent().index(cell._element)
            break  

    # Добавляем новую строку в таблицу
    new_row_index = len(table.rows)
    new_row = table.add_row()

    # Список уже добавленных ключей
    added_keywords = []

    # Если найдены ключевые слова, добавляем их и дату в таблицу
    cell = table.cell(new_row_index, column_index)
    if found_keywords:
        for found_keyword in found_keywords:
            # Если ключ уже добавлен, пропускаем его
            if found_keyword in added_keywords:
                continue
            
            key_description = keywords.get(found_keyword)
            
            if key_description is None:
                print(f"Описание для ключа '{found_keyword}' не найдено.")
                continue
            
            cell.text += f"{key_description}"
            if found_date:
                cell.text += f", от {found_date}"
           
            # Добавляем ключ в список уже добавленных
            added_keywords.append(found_keyword)

            run = cell.paragraphs[0].runs[0]
            font = run.font
            num_char = len(cell.text)
            print(f"Добавлено {num_char} символов для страницы:")

            if num_char >= 30:
                font_size = 9  # Устанавливаем желаемый размер шрифта
                font.size = Pt(font_size)
            if num_char >= 65:
                font_size = 7.5 # Устанавливаем желаемый размер шрифта
                font.size = Pt(font_size)

            font.spacing = Pt(-0.5)  # Устанавливаем пробелы между словами
            font.kerning = True  # Включаем уменьшение расстояния между буквами

            # Сжатие текста в ячейке
            cell.paragraphs[0].paragraph_format.keep_together = True
    else:
        cell.text += ""  # Добавляем пустую строку, если ключевые слова не найдены

    # Добавляем диапазон страниц в столбец "Номера листов"
    if start_page == end_page:
        pages_range = f"{start_page}"" "
    else:
        pages_range = f"{start_page}-{end_page}"

    list_cell = table.cell(new_row_index, list_index)
    list_cell.text = pages_range
    list_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER  # Выравнивание по центру

    list_num = table.cell(new_row_index, num_index + 1)
    list_num.text = str(new_row_index - 1)
    list_num.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER  # Выравнивание по центру
    
    # Устанавливаем выравнивание по центру для ячейки с диапазоном страниц
    for paragraph in list_cell.paragraphs:
        for run in paragraph.runs:
            run.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Находим номер, удовлетворяющий заданным условиям
    first_matching_number = find_first_matching_number(word_path)
    if first_matching_number:
        incoming_cell = table.cell(new_row_index, incoming_index)
        incoming_cell.text = first_matching_number

    doc.save(word_path)


def process_image(image_path, keywords, word_path):
    reader = easyocr.Reader(['en', 'ru'], gpu=True)

    # Поиск ключевых слов и даты
    found_keywords = []
    found_date = None  # Здесь будем хранить найденную дату

    # Распознаем текст на изображении
    result = reader.readtext(image_path)
    for detection in result:
        text = detection[1]
        for keyword in keywords:
            if keyword in text:
                print(f"Ключевое слово '{keyword}' найдено")
                found_keywords.append(keyword)

        # Поиск даты в тексте
        if not found_date:  # Проверяем, была ли найдена дата ранее
            date = find_dates(text)
            
            if date:
                print("Дата найдена:", date)
                found_date = date
            else:
                print("Дата не найдена")

    update_word_table(word_path, keywords, found_keywords, found_date)

    return found_keywords, found_date



def find_first_matching_number(word_path):
    doc = Document(word_path)
    for paragraph in doc.paragraphs:
        text = paragraph.text
        match = re.search(r'№\d{1,5}дск', text)  # Регулярное выражение для поиска номера
        if match:
            print("Номер найден после №:", match.group())
            return match.group()
            
    return None



def find_dates(text):
    # Шаблон для поиска даты в формате DD.MM.YYYY
    date_pattern = r'\b\d{2}[,.]?\d{2}[,.]?\d{4}\b'

    # Находим первое совпадение с шаблоном
    match = re.search(date_pattern, text)
    if match:
        print("Дата найдена:", match.group())
        return match.group()
    else:
        return None

def read_keys(keys_path):
    keys = {}
    doc = Document(keys_path)
    table = doc.tables[0]  # Предполагаем, что таблица находится на первой странице документа
    for row in table.rows[1:]:  # Пропускаем первую строку, так как это заголовок
        key = row.cells[0].text.strip()  # Берем текст из второй ячейки в строке (столбец "Значение ключа")
        key_description = row.cells[1].text.strip()  # Берем текст из третьей ячейки в строке (столбец "Описание ключа")
        
        keys[key] = key_description
        print(f"Добавлен Ключ: '{key}' С описанием: '{key_description}'")
    return keys

if __name__ == "__main__":
    file_path = input("Введите путь к файлу (PDF, JPEG): ")
    word_path = "result.docx"
    keys_path = "keys.docx"
    keywords = read_keys(keys_path)
    found_keywords, found_date = process_pdf(file_path, keywords, word_path)
    try:
        update_word_table(word_path, keywords, found_keywords, found_date)  # Передаем словарь с описаниями ключей в функцию
    except Exception as e:
        print("Конец")
