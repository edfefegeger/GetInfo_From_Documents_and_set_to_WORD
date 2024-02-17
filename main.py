import fitz
import easyocr
from docx import Document
import re
import torch
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

def process_pdf(pdf_path, keywords, word_path, is_two):
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
                        print(f"Ключевое слово '{keyword}' найдено с описание ")
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
                        update_word_table(word_path, keywords, found_keywords, found_date, start_page, end_page, is_two)
                        found_keywords = []
                        found_date = None
                        start_page = page_num + 2  # Начальная страница следующего документа 
  

        # Записываем информацию в файл Word после окончания обработки документа

    return found_keywords, found_date

def update_word_table(word_path, keywords, found_keywords, found_date, start_page, end_page, is_two):
    doc = Document(word_path)
    table = doc.tables[0]
    count = 0

    # Найдем индексы столбцов в таблице
    for cell in table.rows[0].cells:
        if cell.text.strip() == "Наименование документа":
            column_index = cell._element.getparent().index(cell._element)
        elif cell.text.strip() == "Номера листов":
            list_index = cell._element.getparent().index(cell._element)
        elif cell.text.strip() == "исходящие":
            incoming_index = cell._element.getparent().index(cell._element)
        elif cell.text.strip() == "№ з/п":
            num_index = cell._element.getparent().index(cell._element) - 2

    # Добавляем новую строку в таблицу
    new_row_index = len(table.rows)
    new_row = table.add_row()

    # Если найдены ключевые слова, добавляем их и дату в таблицу
    cell = table.cell(new_row_index, column_index)
    if found_keywords:
        for found_keyword in found_keywords:
            key_description = keywords.get(found_keyword)
            
            if key_description is None:
                print(f"Описание для ключа '{found_keyword}' не найдено.")
                continue
            key_texts = key_description['description']  # Список строк, содержащих обе части описания
            if found_date:
                key_texts = [text + f", от {found_date}" for text in key_texts]  # Добавляем дату ко всем строкам описания

            # Добавляем каждую часть описания в новую строку таблицы
            for text_part in key_texts:
                if count != 0:
                    if is_two == True: 
                        
                        new_row = table.add_row()
                    
                new_row.cells[column_index].text = text_part
                count += 1

            # Применяем форматирование к каждой части описания
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    apply_format(run, key_description['format'])

    # Добавляем диапазон страниц в столбец "Номера листов"
    if start_page == end_page:
        pages_range = f"{start_page}"" "
    else:
        pages_range = f"{start_page}-{end_page}"

    list_cell = table.cell(new_row_index, list_index)
    list_cell.text = pages_range
    list_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER  # Выравнивание по центру

    # Добавляем номер заказа в соответствующую ячейку
    list_num = table.cell(new_row_index, num_index + 1)
    list_num.text = str(new_row_index - 1)
    list_num.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER  # Выравнивание по центру
    
    # Если есть информация об исходящих, добавляем ее
    first_matching_number = find_first_matching_number(word_path)
    if first_matching_number:
        incoming_cell = table.cell(new_row_index, incoming_index)
        incoming_cell.text = first_matching_number
        incoming_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER  # Выравнивание по центру

    doc.save(word_path)


def apply_format(run, format_dict):
    if format_dict['bold'] is not None:
        run.bold = format_dict['bold']
    if format_dict['italic'] is not None:
        run.italic = format_dict['italic']
    if format_dict['underline'] is not None:
        run.underline = format_dict['underline']
    if format_dict['font_color'] is not None:
        run.font.color.rgb = format_dict['font_color']
    if format_dict['font_size'] is not None:
        run.font.size = format_dict['font_size']
    if format_dict['font_name'] is not None:
        run.font.name = format_dict['font_name']
    if format_dict['highlight_color'] is not None:
        run.font.highlight_color = format_dict['highlight_color']
    if format_dict['superscript'] is not None:
        run.font.superscript = format_dict['superscript']
    if format_dict['subscript'] is not None:
        run.font.subscript = format_dict['subscript']
    if format_dict['strike'] is not None:
        run.font.strike = format_dict['strike']
    if format_dict['double_strike'] is not None:
        run.font.double_strike = format_dict['double_strike']
    if format_dict['all_caps'] is not None:
        run.font.all_caps = format_dict['all_caps']
    if format_dict['small_caps'] is not None:
        run.font.small_caps = format_dict['small_caps']
    if format_dict['shadow'] is not None:
        run.font.shadow = format_dict['shadow']
    if format_dict['outline'] is not None:
        run.font.outline = format_dict['outline']
    if format_dict['emboss'] is not None:
        run.font.emboss = format_dict['emboss']
    if format_dict['imprint'] is not None:
        run.font.imprint = format_dict['imprint']





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
    is_two = False  # Initialize is_two here
    doc = Document(keys_path)
    key = ''  # Переменная для хранения текущего ключа
    description = ''  # Переменная для хранения описания текущего ключа
    description2 = ''  # Переменная для хранения второй половины описания текущего ключа
    cell_format = None  # Переменная для хранения форматирования текущего ключа
    
    for row in doc.tables[0].rows[1:]:  # Пропускаем первую строку, так как это заголовок
        cell_0_text = row.cells[1].text.strip()
        cell_0_number = row.cells[0].text.strip()

        if cell_0_number:  # Если ячейка не пустая, это начало нового ключа
            # Если есть предыдущий ключ, сохраняем его в словарь
            if key:
                keys[key] = {'description': [description, description2], 'format': cell_format}  # Сохраняем информацию о форматировании в keys
                print(f"Добавлен Ключ: {key} с описанием: {description}")
            key = cell_0_text  # Обновляем текущий ключ
            description = row.cells[2].text.strip()  # Берем текст из второй ячейки в строке (столбец "Описание ключа")
            description2 = ''  # Сбрасываем вторую половину описания для нового ключа
            cell_format = {}  # Сбрасываем форматирование для нового ключа
            for paragraph in row.cells[1].paragraphs:
                for run in paragraph.runs:
                    cell_format['bold'] = run.bold
                    cell_format['italic'] = run.italic
                    cell_format['underline'] = run.underline
                    cell_format['font_color'] = run.font.color.rgb
                    cell_format['font_size'] = run.font.size
                    cell_format['font_name'] = run.font.name
                    cell_format['highlight_color'] = run.font.highlight_color
                    cell_format['superscript'] = run.font.superscript
                    cell_format['subscript'] = run.font.subscript
                    cell_format['strike'] = run.font.strike
                    cell_format['double_strike'] = run.font.double_strike
                    cell_format['all_caps'] = run.font.all_caps
                    cell_format['small_caps'] = run.font.small_caps
                    cell_format['shadow'] = run.font.shadow
                    cell_format['outline'] = run.font.outline
                    cell_format['emboss'] = run.font.emboss
                    cell_format['imprint'] = run.font.imprint
        else:
            # Если ячейка пустая, это вторая половина описания ключа
            is_two = True
            description2 += " " + row.cells[2].text.strip()  # Добавляем новую строку к текущей второй половине описания
            key += " " + row.cells[1].text.strip()  # Добавляем новую строку к текущему ключу

    # Сохраняем информацию о последнем ключе
    if key:
        keys[key] = {'description': [description, description2], 'format': cell_format}  # Сохраняем информацию о форматировании в keys
        print(f"Добавлен Ключ: {key} с описанием: {description}")

    return keys, is_two




if __name__ == "__main__":
    file_path = input("Введите путь к файлу (PDF, JPEG): ")
    word_path = "result.docx"
    keys_path = "keys.docx"
    keywords, is_two = read_keys(keys_path)
    found_keywords, found_date = process_pdf(file_path, keywords, word_path, is_two)
    try:
        update_word_table(word_path, keywords, found_keywords, found_date, is_two)  # Передаем словарь с описаниями ключей в функцию
    except Exception as e:
        print("Конец")
