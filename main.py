import fitz
import easyocr
from docx import Document
import re

def process_pdf(pdf_path, keywords):
    reader = easyocr.Reader(['en', 'ru'])

    # Поиск ключевых слов и даты
    found_keywords = []
    with fitz.open(pdf_path) as pdf:
        for page in pdf:
            # Получаем изображения на странице
            images = page.get_images(full=True)
            for img_index, img in enumerate(images):
                xref = img[0]
                base_image = pdf.extract_image(xref)
                image_bytes = base_image["image"]

                # Распознаем текст на изображении
                result = reader.readtext(image_bytes)
                for detection in result:
                    text = detection[1]
                    found = False
                    for keyword in keywords:
                        if keyword in text:
                            print(f"Ключевое слово '{keyword}' найдено")
                            found_keywords.append(keyword)
                            found = True
                            break
                    if not found:
                        print(f"Ключевое слово не найдено в тексте: {text}")

                    # Поиск даты в тексте
                    date = find_dates(text)
                    if date:
                        print("Дата найдена:", date)
                        found_keywords.append(date)

    return found_keywords

def find_dates(text):
    # Шаблон для поиска даты в формате DD.MM.YYYY
    date_pattern = r'\b\d{2}\.\d{2}\.\d{4}\b'

    # Находим первое совпадение с шаблоном
    match = re.search(date_pattern, text)
    if match:
        return match.group()
    else:
        return None

def update_word_table(pdf_path, keywords, word_path):
    found_keywords = process_pdf(pdf_path, keywords)

    doc = Document(word_path)
    table = doc.tables[0]  

    for cell in table.rows[0].cells:
        if cell.text.strip() == "Наименование документа":
            column_index = cell._element.getparent().index(cell._element)

    new_row_index = 1 
    new_row = table.add_row()

    if found_keywords:
        table.cell(new_row_index + 1, column_index).text = found_keywords[0]  

    doc.save(word_path)

if __name__ == "__main__":
    pdf_path = input("Введите путь к PDF-файлу: ")
    word_path = "test.docx"
    keywords = input("Введите ключевые слова через пробел: ").split()
    update_word_table(pdf_path, keywords, word_path)
