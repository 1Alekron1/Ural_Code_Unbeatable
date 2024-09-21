from docx import Document
import re

from docx.shared import RGBColor


def check_structural_elements(doc):
    required_sections = [
        "Титульный лист", "Список исполнителей", "Реферат",
        "Содержание", "Термины и определения",
        "Перечень сокращений и обозначений", "Введение",
        "Основная часть отчета о НИР", "Заключение"
    ]

    found_sections = []
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip().lower()
        for section in required_sections:
            if section.lower() in text:
                found_sections.append(section)

    missing_sections = set(required_sections) - set(found_sections)
    return missing_sections, found_sections


def extract_references_list(doc):
    # Находим раздел "Литература" или "Список литературы" и извлекаем все ссылки
    is_reference_section = False
    references = []

    for paragraph in doc.paragraphs:
        text = paragraph.text.strip().lower()
        if "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ".lower() in text:
            is_reference_section = True
        elif is_reference_section and text:  # добавляем ссылки, пока не кончился раздел
            references.append(paragraph)


    return references


def check_references_in_text(doc, references):
    # Поиск всех сносок в квадратных скобках
    references_in_text = []
    reference_pattern = re.compile(r'\[(\d+)\]')

    for paragraph in doc.paragraphs:
        matches = reference_pattern.findall(paragraph.text)
        if matches:
            references_in_text.extend([int(m) for m in matches])

    # Проверяем индексацию
    is_indexed_correctly = references_in_text == sorted(references_in_text)

    # Проверка на использование всех ссылок из списка литературы
    used_references = set(references_in_text)
    all_references = set(range(1, len(references) + 1))
    unused_references = all_references - used_references

    return references_in_text, is_indexed_correctly, unused_references


def highlight_unused_references(doc, references, unused_references):
    for i, paragraph in enumerate(references):
        ref_num = i + 1
        if ref_num in unused_references:
            # Выделяем красным неиспользованные ссылки
            for run in paragraph.runs:
                run.font.color.rgb = RGBColor(255, 0, 0)  # Устанавливаем красный цвет текста


def main(doc_path):
    doc = Document(doc_path)

    # Проверка структурных элементов
    missing_sections, found_sections = check_structural_elements(doc)
    print(f"Найдены разделы: {found_sections}")
    print(f"Отсутствуют разделы: {missing_sections}")

    # Извлечение списка литературы
    reference_paragraphs = extract_references_list(doc)
    references = [p.text for p in reference_paragraphs]
    print(f"Найденные ссылки на литературу: {references}")

    # Проверка ссылок в тексте
    references_in_text, is_indexed_correctly, unused_references = check_references_in_text(doc, references)
    print(f"Найденные ссылки в тексте: {references_in_text}")
    if is_indexed_correctly:
        print("Индексация ссылок правильная.")
    else:
        print("Ошибка в индексации ссылок.")

    if unused_references:
        print(f"Неиспользованные ссылки из списка литературы: {unused_references}")
        # Выделяем красным цветом неиспользованные ссылки
        highlight_unused_references(doc, reference_paragraphs, unused_references)

    # Сохранение документа с примечаниями
    doc.save("output_with_notes.docx")


# Использование:
doc_path = "Бордунов Александр Максимович Отчет по работе — копия.docx"
main(doc_path)
