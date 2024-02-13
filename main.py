import fitz
import easyocr
from docx import Document
import re

def process_pdf(pdf_path, keywords):
    reader = easyocr.Reader(['en', 'ru'])

    # Поиск ключевых слов и даты
    found_keywords = []
    found_date = None  # Здесь будем хранить найденную дату

    with fitz.open(pdf_path) as pdf:
        for page_num in range(len(pdf)):
            page = pdf.load_page(page_num)
            text = page.get_text()

            if text:  # Если на странице есть текст
                # Поиск ключевых слов
                for keyword in keywords:
                    if keyword in text:
                        print(f"Ключевое слово '{keyword}' найдено")
                        found_keywords.append(keyword)

                # Поиск даты
                if not found_date:  # Проверяем, была ли найдена дата ранее
                    date = find_dates(text)
                    if date:
                        print("Дата найдена:", date)
                        found_date = date

            else:  # Если на странице нет текста, попытаемся найти его в виде изображения
                images = page.get_images(full=True)
                for img_index, img in enumerate(images):
                    xref = img[0]
                    base_image = pdf.extract_image(xref)
                    image_bytes = base_image["image"]

                    # Распознаем текст на изображении
                    result = reader.readtext(image_bytes)
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

    return found_keywords, found_date


def process_image(image_path, keywords):
    reader = easyocr.Reader(['en', 'ru'])

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

    return found_keywords, found_date

def find_dates(text):
    # Шаблон для поиска даты в формате DD.MM.YYYY
    date_pattern = r'\b\d{2}[,.]?\d{2}[,.]?\d{4}\b'

    # Находим первое совпадение с шаблоном
    match = re.search(date_pattern, text)
    if match:
        print("Дата найдена:", match.group())
        return match.group()
    else:
        print("Дата не найдена")
        return None

def update_word_table(file_path, keywords, word_path):
    if file_path.endswith('.pdf'):
        found_keywords, found_date = process_pdf(file_path, keywords)
    elif file_path.endswith(('.jpg', '.jpeg')):
        found_keywords, found_date = process_image(file_path, keywords)
    else:
        print("Неподдерживаемый формат файла.")
        return

    doc = Document(word_path)
    table = doc.tables[0]

    # Добавляем новую строку в таблицу
    new_row_index = len(table.rows)
    new_row = table.add_row()

    # Находим индекс столбца "Наименование документа"
    for cell in table.rows[0].cells:
        if cell.text.strip() == "Наименование документа":
            column_index = cell._element.getparent().index(cell._element)
            break

    if found_keywords:
        found_keyword = found_keywords[0]
        table.cell(new_row_index, column_index).text = found_keyword

        # Добавляем дату в конец найденного ключевого слова
        if found_date:
            table.cell(new_row_index, column_index).text += f", от {found_date}"

    doc.save(word_path)

if __name__ == "__main__":
    file_path = input("Введите путь к файлу (PDF, JPEG): ")
    word_path = "test.docx"
    keywords = input("Введите ключевые слова через пробел: ").split()
    update_word_table(file_path, keywords, word_path)
