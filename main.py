import fitz
import easyocr
from docx import Document
import re
import torch
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from fuzzywuzzy import fuzz

def clear_word_table(word_path):
    doc = Document(word_path)
    table = doc.tables[0]
    for row in table.rows[2:]:  
        table._element.remove(row._element)  # Удаляем строку
    doc.save(word_path)

def process_pdf(pdf_path, keywords, word_path, threshold, languages, text_q, count):
    count = 1

    reader = easyocr.Reader(['en', languages], gpu=True)

    # Поиск ключевых слов и даты
    found_keywords = []
    found_date = None  # Здесь будем хранить найденную дату
    found_outgoing_num = None
    total_text = ""
    total_found_keywords = []
    start_page = 1  # Начальная страница текущего документа

    with fitz.open(pdf_path) as pdf:
        for page_num in range(len(pdf)):
            page = pdf.load_page(page_num)
            text = page.get_text()

            print(f"Обрабатывается страница {page_num + 1}...")

            if text_q == 'y':
                print("Распознанный текст на странице:")
                print(text)  # Выводим распознанный текст страницы

            # Если на странице есть текст, обрабатываем ее
            if text:
                # Поиск ключевых слов
                for keyword in keywords:
                    if keyword in found_keywords:
                        continue  # Пропускаем ключевое слово, если оно уже было найдено
                    key_words = keyword.split()  # Разбиваем ключевое слово на отдельные слова
                    found_count = sum(word in text for word in key_words)  # Подсчитываем количество найденных слов
                    # Вычисляем процент распознавания для ключа
                    recognition_percentage = (found_count / len(key_words)) * 100
                    if recognition_percentage >= threshold:
                        print(f"Ключевое слово '{keyword}' найдено с процентом распознавания {recognition_percentage}%")
                        total_found_keywords.append(keyword)
                    else: 
                        print(f"Ключевое слово '{keyword}' не добавлено с процентом распознавания {recognition_percentage}%")

                # Поиск даты, если она еще не была найдена
                if not found_date:
                    date = find_dates(text)
                    if date:
                        print("Дата найдена:", date)
                        found_date = date 

                # Поиск исходящего номера, если он еще не был найден
                if not found_outgoing_num:
                    outgoing_num = find_first_matching_number(text)
                    if outgoing_num:
                        print(f"Найден исходящий номер {outgoing_num}")
                        found_outgoing_num = outgoing_num

            # Если на странице есть изображения, ищем текст в них
            images = page.get_images(full=True)
            if images:
                for img_index, img in enumerate(images):
                    xref = img[0]
                    base_image = pdf.extract_image(xref)
                    image_bytes = base_image["image"]
                    result = reader.readtext(image_bytes)
                    for detection in result:
                        img_text = detection[1]
                        total_text += " " + img_text  # Добавляем текст изображения к тексту страницы
                    if text_q == 'y':
                        print("Распознанный текст на странице:")
                        print(text)  # Выводим распознанный текст страницы
                        # print("Распознанный текст на странице:")
                        # print(text)  # Выводим распознанный текст страницы

                # Поиск ключевых слов после добавления текста изображения
                for keyword in keywords:
                    if keyword in found_keywords:
                        continue  # Пропускаем ключевое слово, если оно уже было найдено
                    key_words = keyword.split()  # Разбиваем ключевое слово на отдельные слова
                    found_count = sum(word in text for word in key_words)  # Подсчитываем количество найденных слов
                    # Вычисляем процент распознавания для ключа
                    recognition_percentage = (found_count / len(key_words)) * 100
                    if recognition_percentage >= threshold:
                        print(f"Ключевое слово '{keyword}' с процентом распознавания {recognition_percentage}%")
                        total_found_keywords.append(keyword)
                    else: 
                        print(f"Ключевое слово '{keyword}' не добавлено с процентом распознавания {recognition_percentage}%")

                # Поиск исходящего номера после добавления текста изображения
                if not found_outgoing_num:
                    outgoing_num = find_first_matching_number(text)
                    if outgoing_num:
                        print(f"Найден исходящий номер  {outgoing_num}")
                        found_outgoing_num = outgoing_num

                # Поиск даты после добавления текста изображения
                if not found_date:
                    date = find_dates(text)
                    if date:
                        print("Дата найдена:", date)
                        found_date = date 

            if len(pdf) == 1:  # Если документ содержит только одну страницу
                print("Документ содержит только одну страницу. Завершение обработки.")
                end_page = 1  # Конечная страница текущего документа
                update_word_table(word_path, keywords, found_keywords, found_date, start_page, end_page, found_outgoing_num, count)
                found_keywords = []
                found_date = None
                found_outgoing_num = None
            else:
                if "End" in text:
                    # Вычисляем процент ключевых слов на всех страницах
                    total_keyword_count = sum(sum(word in total_text for word in keyword.split()) for keyword in keywords)
                    if total_text.strip():  # Проверяем, содержит ли total_text хотя бы одно слово
                        total_keywords_percentage = (total_keyword_count / len(total_text.split())) * 100
                    else:
                        total_keywords_percentage = 0  # Устанавливаем процент ключевых слов равным нулю, если total_text пуст
                    print(f"Суммарный процент ключевых слов до страницы 'End': {total_keywords_percentage}%")

                    print(f"Найдена пометка 'End' на странице {page_num + 1}. Завершение документа.")
                    end_page = page_num + 1  # Конечная страница текущего документа
                    update_word_table(word_path, keywords, total_found_keywords, found_date, start_page, end_page, found_outgoing_num, count)
                    count += 1
                    total_found_keywords = []
                    found_date = None
                    found_outgoing_num = None
                    start_page = page_num + 2  # Начальная страница следующего документа      

    # Записываем информацию в файл Word после окончания обработки документа
    return total_found_keywords, found_date



def update_word_table(word_path, keywords, found_keywords, found_date, start_page, end_page, found_outgoing_num, count):
    doc = Document(word_path)
    table = doc.tables[0]
    outgoing_index = None  # Добавляем индекс для столбца с исходящим номером

    font_name = 'Times New Roman'
    font_size = Pt(12)
    # Находим индекс столбца "Наименование документа"
    for cell in table.rows[0].cells:
        if cell.text.strip() == "Наименование документа":
            column_index = cell._element.getparent().index(cell._element)
            break
    for cell in table.rows[0].cells:
        if cell.text.strip() == "Номера листов":
            list_index = cell._element.getparent().index(cell._element)
            break   
    for cell in table.rows[1].cells:
        if cell.text.strip() == "исходящий":
            outgoing_index = cell._element.getparent().index(cell._element) - 1
            break 
    for cell in table.rows[0].cells:
        if cell.text.strip() == "№ з/п":
            num_index = cell._element.getparent().index(cell._element) - 2
            break 

    # Добавляем новую строку в таблицу
    new_row_index = len(table.rows)
    new_row = table.add_row()
    added_keywords = []
    # Если найдены ключевые слова, добавляем их и дату в таблицу
    cell = table.cell(new_row_index, column_index)
    if found_keywords:
        found_keyword = found_keywords[0]  # Берем только первое найденное ключевое слово
        key_description = keywords.get(found_keyword)

        if key_description is None:
            print(f"Описание для ключа '{found_keyword}' не найдено.")
            return
        
        key_text = key_description['description']
        key_text2 = key_description['description2'] # Получаем description2, если он есть, или пустую строку

        



        if key_text2 != "":
            if found_date:
                key_text2 += f", от {found_date}"
            is_two_str = True
            new_row = table.add_row()
            column_cell = new_row.cells[column_index]  # Получаем ячейку в нужном столбце
            # Используем column_cell вместо column_index
            run2 = column_cell.paragraphs[0].add_run(key_text2)
        # Применяем форматирование к тексту
            key_format2 = key_description['format']
            # column_index.text = key_text2  # Обновляем текст ячейки с key_text2
            if key_format2['bold'] is not None:
                run2.bold = key_format2['bold']
            if key_format2['italic'] is not None:
                run2.italic = key_format2['italic']
            if key_format2['underline'] is not None:
                run2.underline = key_format2['underline']
            if key_format2['font_color'] is not None:
                run2.font.color.rgb = key_format2['font_color']
            if key_format2['font_size'] is not None:
                run2.font.size = key_format2['font_size']
            if key_format2['font_name'] is not None:
                run2.font.name = key_format2['font_name']
            if key_format2['highlight_color'] is not None:
                run2.font.highlight_color = key_format2['highlight_color']
            if key_format2['superscript'] is not None:
                run2.font.superscript = key_format2['superscript']
            if key_format2['subscript'] is not None:
                run2.font.subscript = key_format2['subscript']
            if key_format2['strike'] is not None:
                run2.font.strike = key_format2['strike']
            if key_format2['double_strike'] is not None:
                run2.font.double_strike = key_format2['double_strike']
            if key_format2['all_caps'] is not None:
                run2.font.all_caps = key_format2['all_caps']
            if key_format2['small_caps'] is not None:
                run2.font.small_caps = key_format2['small_caps']
            if key_format2['shadow'] is not None:
                run2.font.shadow = key_format2['shadow']
            if key_format2['outline'] is not None:
                run2.font.outline = key_format2['outline']
            if key_format2['emboss'] is not None:
                run2.font.emboss = key_format2['emboss']
            if key_format2['imprint'] is not None:
                run2.font.imprint = key_format2['imprint']
        else:
            if found_date:
                key_text += f", от {found_date}"
        cell_paragraphs = cell.paragraphs
        if not cell_paragraphs:  # Если в ячейке нет абзацев, создаем новый
            new_paragraph = cell.add_paragraph()
        else:
            new_paragraph = cell_paragraphs[-1]  # Или берем последний абзац, если он уже существует
        # Добавляем текст с форматированием
        run = new_paragraph.add_run(key_text)
        # Применяем форматирование к тексту
        key_format = key_description['format']
        if key_format['bold'] is not None:
            run.bold = key_format['bold']
        if key_format['italic'] is not None:
            run.italic = key_format['italic']
        if key_format['underline'] is not None:
            run.underline = key_format['underline']
        if key_format['font_color'] is not None:
            run.font.color.rgb = key_format['font_color']
        if key_format['font_size'] is not None:
            run.font.size = key_format['font_size']
        if key_format['font_name'] is not None:
            run.font.name = key_format['font_name']
        if key_format['highlight_color'] is not None:
            run.font.highlight_color = key_format['highlight_color']
        if key_format['superscript'] is not None:
            run.font.superscript = key_format['superscript']
        if key_format['subscript'] is not None:
            run.font.subscript = key_format['subscript']
        if key_format['strike'] is not None:
            run.font.strike = key_format['strike']
        if key_format['double_strike'] is not None:
            run.font.double_strike = key_format['double_strike']
        if key_format['all_caps'] is not None:
            run.font.all_caps = key_format['all_caps']
        if key_format['small_caps'] is not None:
            run.font.small_caps = key_format['small_caps']
        if key_format['shadow'] is not None:
            run.font.shadow = key_format['shadow']
        if key_format['outline'] is not None:
            run.font.outline = key_format['outline']
        if key_format['emboss'] is not None:
            run.font.emboss = key_format['emboss']
        if key_format['imprint'] is not None:
            run.font.imprint = key_format['imprint']
    else:
        null_row = table.add_row()
        column_cell = null_row.cells[column_index]  # Получаем ячейку в нужном столбце
        # Используем column_cell вместо column_index
        null_row.text = ""  # Обновляем текст ячейки с key_text

    # Добавляем диапазон страниц в столбец "Номера листов"
    if start_page == end_page:
        pages_range = f"{start_page}"" "
    else:
        pages_range = f"{start_page}-{end_page}"

    list_cell = table.cell(new_row_index, list_index)
    list_cell.text = pages_range
    list_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER  # Выравнивание по центру
    # Применяем параметры шрифта к тексту ячейки
    for paragraph in list_cell.paragraphs:
        for run in paragraph.runs:
            run.font.name = font_name
            run.font.size = font_size
    
    # if is_two_str == False:
        # Добавляем номер заказа в соответствующую ячейку
    if outgoing_index is not None:  
        outgoing_cell = table.cell(new_row_index, outgoing_index)
        if found_outgoing_num:
            outgoing_cell.text = found_outgoing_num
        outgoing_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER  # Выравнивание по центру
        # Применяем параметры шрифта к тексту ячейки
        for paragraph in outgoing_cell.paragraphs:
            for run in paragraph.runs:
                run.font.name = font_name
                run.font.size = font_size
    else:
        print("Столбец 'исходящие' не найден в таблице.")

    list_num = table.cell(new_row_index, num_index + 1)
    list_num.text = str(count),
    list_num.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER  # Выравнивание по центру
    # Применяем параметры шрифта к тексту ячейки
    for paragraph in list_num.paragraphs:
        for run in paragraph.runs:
            run.font.name = font_name
            run.font.size = font_size

    doc.save(word_path)



def find_first_matching_number(text):
    pattern = r'№(\d{1,5})(дск)'  # Включаем номер и "дск" в группы, чтобы их можно было извлечь
    match = re.search(pattern, text)
    if match:
        return match.group(1) + match.group(2)  # Возвращаем значение номера и "дск"
    else:           
        return None



def find_dates(text):
    # Шаблон для поиска даты в формате DD.MM.YYYY
    date_pattern = r'\b\d{2}[,.]?\d{2}[,.]?\d{4}\b'
    # Находим первое совпадение с шаблоном
    match = re.search(date_pattern, text)
    if match:
        return match.group()
    else:
        return None

def read_keys(keys_path):
    keys = {}
    doc = Document(keys_path)
    key = ''  # Переменная для хранения текущего ключа
    description = ''  # Переменная для хранения описания текущего ключа
    description2 = ''  # Переменная для хранения второго описания текущего ключа
    cell_format = None  # Переменная для хранения форматирования текущего ключа
    
    for row in doc.tables[0].rows[1:]:  # Пропускаем первую строку, так как это заголовок
        cell_0_text = row.cells[1].text.strip()
        cell_0_number = row.cells[0].text.strip()

        if cell_0_number:  # Если ячейка не пустая, это начало нового ключа
            # Если есть предыдущий ключ, сохраняем его в словарь
            if key:
                keys[key] = {'description': description, 'description2': description2, 'format': cell_format}  # Сохраняем информацию о форматировании в keys
                print(f"Добавлен Ключ: {key}. с описанием: {description} {description2}")
            key = cell_0_text  # Обновляем текущий ключ
            description = row.cells[2].text.strip()  # Берем текст из второй ячейки в строке (столбец "Описание ключа")
            description2 = ''  # Обнуляем description2 для нового ключа
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
            # Если ячейка пустая, это продолжение описания или ключа
                description2 += row.cells[2].text.strip()  # Добавляем новую строку к текущему описанию 2
                key += " " + row.cells[1].text.strip()  # Добавляем новую строку к текущему ключу
    
    # Сохраняем информацию о последнем ключе
    if key:
        keys[key] = {'description': description, 'description2': description2, 'format': cell_format}  # Сохраняем информацию о форматировании в keys
        print(f"Добавлен Ключ: {key}. с описанием: {description} {description2}")

    return keys




if __name__ == "__main__":
    file_path = input("Введите путь к файлу (PDF): ")
    threshold = int(input("Введите минимальное пороговое значение для распознавания текста в %: "))
    languages = input("Введите язык для использования (ru или uk или be) ")
    text_q = input("Выводить распознаный текст? (y = да, n = нет) ")

    word_path = "result.docx"
    keys_path = "keys.docx"
    keywords = read_keys(keys_path)
    clear_word_table(word_path)  # Очищаем таблицу перед обработкой нового файла PDF
    print("Таблица 'Result.docx' очищена")
    count = 1
    found_keywords, found_date = process_pdf(file_path, keywords, word_path, threshold, languages, text_q, count)
    try:
        update_word_table(word_path, keywords, found_keywords, found_date, count)  # Передаем словарь с описаниями ключей в функцию
    except Exception as e:
        print("Конец")

    input("Press Enter to exit...")
