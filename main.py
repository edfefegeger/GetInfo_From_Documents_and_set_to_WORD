import fitz
import easyocr
from docx import Document
import re
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
                        print(f"Ключевое слово '{keyword}' ")
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
    # Находим индекс столбцов
    column_indices = {}
    for cell in table.rows[0].cells:
        if cell.text.strip() == "Наименование документа":
            column_indices['name'] = cell._element.getparent().index(cell._element)
        elif cell.text.strip() == "Номера листов":
            column_indices['pages'] = cell._element.getparent().index(cell._element)
        elif cell.text.strip() == "исходящие":
            column_indices['outgoing'] = cell._element.getparent().index(cell._element)
        elif cell.text.strip() == "№ з/п":
            column_indices['number'] = cell._element.getparent().index(cell._element) - 2

    # Добавляем новую строку в таблицу для каждого найденного ключа
    for found_keyword in found_keywords:
        key_description = keywords.get(found_keyword)
        if key_description is None:
            print(f"Описание для ключа '{found_keyword}' не найдено.")
            continue
        key_text = key_description['description']
        if found_date:
            key_text += f", от {found_date}"
        # Разделяем описание ключа на несколько строк, если необходимо
        description_cells = key_text.split('\n')
        max_cells = max(len(description_cells), 1)  # Максимальное количество ячеек, которое нужно добавить
        for i in range(max_cells):
            new_row_index = len(table.rows)
            new_row = table.add_row()
            # Записываем описание ключа в соответствующую ячейку
            if i < len(description_cells):
                cell = table.cell(new_row_index, column_indices['name'])
                cell.text = description_cells[i]
                # Применяем форматирование из словаря keywords к ячейке
                apply_format(cell, key_description['format'])
            # Добавляем диапазон страниц в ячейку "Номера листов"
            if i == 0:
                if start_page == end_page:
                    pages_range = f"{start_page}"
                else:
                    pages_range = f"{start_page}-{end_page}"
                list_cell = table.cell(new_row_index, column_indices['pages'])
                list_cell.text = pages_range
                list_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER  # Выравнивание по центру
            # Добавляем номер заказа в соответствующую ячейку
            list_num = table.cell(new_row_index, column_indices['number'] + 1)
            list_num.text = str(new_row_index - 1)
            list_num.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER  # Выравнивание по центру

            # Добавляем исходящий номер в соответствующую ячейку, если он есть
            if i == 0:
                first_matching_number = find_first_matching_number(word_path)
                if first_matching_number:
                    incoming_cell = table.cell(new_row_index, column_indices['outgoing'])
                    incoming_cell.text = first_matching_number
                    incoming_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER  # Выравнивание по центру

    doc.save(word_path)

def apply_format(cell, format_dict):
    """
    Apply formatting from format_dict to the given cell.
    """
    run = cell.paragraphs[0].runs[0]
    font = run.font
    # Apply formatting properties from format_dict
    font.bold = format_dict.get('bold', False)
    font.italic = format_dict.get('italic', False)
    font.underline = format_dict.get('underline', None)
    font.color.rgb = format_dict.get('font_color', None)
    font.size = format_dict.get('font_size', None)
    font.name = format_dict.get('font_name', None)
    font.highlight_color = format_dict.get('highlight_color', None)
    font.superscript = format_dict.get('superscript', None)
    font.subscript = format_dict.get('subscript', None)
    font.strike = format_dict.get('strike', None)
    font.double_strike = format_dict.get('double_strike', None)
    font.all_caps = format_dict.get('all_caps', None)
    font.small_caps = format_dict.get('small_caps', None)
    font.shadow = format_dict.get('shadow', None)
    font.outline = format_dict.get('outline', None)
    font.emboss = format_dict.get('emboss', None)
    font.imprint = format_dict.get('imprint', None)





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
    key = None  # Переменная для хранения текущего ключа
    key_description = None  # Переменная для хранения описания текущего ключа 
    for row in table.rows[1:]:  # Пропускаем первую строку, так как это заголовок
        cell_1_text = row.cells[0].text.strip()  # Берем текст из первой ячейки в строке (столбец "Значение ключа")
        cell_2_text = row.cells[1].text.strip()  # Берем текст из второй ячейки в строке (столбец "Описание ключа") 
        # Если первая ячейка не пустая, это новый ключ
        if cell_1_text:
            if key:  # Если уже существует текущий ключ, сохраняем его
                # Добавляем информацию о форматировании ячейки в словарь keys
                keys[key] = {'description': key_description, 'format': cell_format} 
            # Инициализируем новый ключ
            key = cell_1_text
            key_description = cell_2_text
            cell_format = {}  # Инициализируем форматирование для нового ключа 
        else:  # Если первая ячейка пустая, это продолжение значения ключа
            # Добавляем описание в предыдущий ключ
            key_description += "\n" + cell_2_text  # Продолжаем описание на новой строке 
            # Обновляем форматирование для текущего ключа
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
    # Добавляем последний ключ в словарь после завершения цикла
    if key:
        keys[key] = {'description': key_description, 'format': cell_format}
    
    # Выводим все ключи и их описания в консоль
    for key, value in keys.items():
        print(f"Ключ: '{key}'")
        print(f"Описание: '{value['description']}'")
    
    return keys




if __name__ == "__main__":
    file_path = input("Введите путь к файлу для обработки: ")
    word_path = "result.docx"
    keys_path = "keys.docx"
    keywords = read_keys(keys_path)
    found_keywords, found_date = process_pdf(file_path, keywords, word_path)
    try:
        update_word_table(word_path, keywords, found_keywords, found_date)  # Передаем словарь с описаниями ключей в функцию
    except Exception as e:
        print("Конец")
