import re

from docx import Document
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
    # Ищем последнее вхождение раздела "Литература" или "Список литературы"
    references = []
    last_reference_section_index = -1

    # Ищем последний индекс вхождения раздела "Литература"
    for i, paragraph in enumerate(doc.paragraphs):
        text = paragraph.text.strip().lower()
        if "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ".lower() in text or "список литературы" in text:
            last_reference_section_index = i

    # Если раздел найден, начинаем сбор ссылок после него
    if last_reference_section_index != -1:
        for paragraph in doc.paragraphs[last_reference_section_index + 2:]:
            text = paragraph.text.strip()
            if text:  # Добавляем ссылки, пока не встретится пустая строка
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


def highlight_incorrectly_indexed_references(doc, references_in_text):
    # Поиск всех сносок и их индексации
    sorted_references = sorted(references_in_text)

    for paragraph in doc.paragraphs:
        text = paragraph.text
        matches = re.finditer(r'\[(\d+)\]', text)
        for match in matches:
            ref_number = int(match.group(1))
            # Если текущая ссылка не на месте (не соответствует правильному порядку)
            if ref_number != sorted_references.pop(0):
                # Подсвечиваем эту ссылку
                start_idx = match.start(1)
                end_idx = match.end(1)
                for run in paragraph.runs:
                    if run.text in match.group(0):  # Подсвечиваем весь текст ссылки
                        run.font.color.rgb = RGBColor(255, 0, 0)  # Красный цвет для неправильной ссылки


def extract_conclusion_text(doc):
    # Ищем раздел "Заключение"
    conclusion_start_index = -1
    conclusion_text = []

    for i, paragraph in enumerate(doc.paragraphs):
        text = paragraph.text.strip().lower()
        if "заключение" == text:
            conclusion_start_index = i
            break

    # Если раздел "Заключение" найден, извлекаем текст до следующего раздела
    if conclusion_start_index != -1:
        for paragraph in doc.paragraphs[conclusion_start_index + 1:]:
            text = paragraph.text.strip()
            # Проверяем, не начался ли следующий раздел (например, другой заголовок)
            if text and text.isupper():
                break  # Следующий раздел найден
            conclusion_text.append(paragraph.text)

    return "\n".join(conclusion_text)


def extract_section_text(doc, section_name):
    section_start_index = -1
    section_text = []

    # Поиск начала указанного раздела
    for i, paragraph in enumerate(doc.paragraphs):
        text = paragraph.text.strip().lower()
        if section_name.lower() == text:
            section_start_index = i
            break

    # Если раздел найден, извлекаем текст до следующего заголовка
    if section_start_index != -1:
        for paragraph in doc.paragraphs[section_start_index + 1:]:
            text = paragraph.text.strip()
            # Проверяем, не начался ли следующий раздел (например, заглавными буквами)
            if text and text.isupper():
                break  # Следующий раздел найден
            section_text.append(paragraph.text)

    return "\n".join(section_text)  # Собираем текст раздела


def main(doc_path):
    result = {"Введение": "", "Заключение": "", "Найдены разделы": [], "Отсутствуют разделы": [],
              "Найденные ссылки на литературу": [], "Найденные ссылки в тексте": [], "Ошибка": "",
              "Неиспользованные ссылки из списка литературы": []}
    doc = Document(doc_path)
    conclusion = extract_section_text(doc, "Заключение")
    entry = extract_section_text(doc, "Введение")
    result["Введение"] = entry
    result["Заключение"] = conclusion

    # Проверка структурных элементов
    missing_sections, found_sections = check_structural_elements(doc)
    result["Найдены разделы"] = found_sections
    result["Отсутствуют разделы"] = missing_sections

    # Извлечение списка литературы
    reference_paragraphs = extract_references_list(doc)
    references = [p.text for p in reference_paragraphs]
    result["Найденные ссылки на литературу"] = references

    # Проверка ссылок в тексте
    references_in_text, is_indexed_correctly, unused_references = check_references_in_text(doc, references)
    result["Найденные ссылки в тексте"] = references_in_text
    if is_indexed_correctly:
        result["Ошибка"] = "Индексация ссылок правильная."
    else:
        result["Ошибка"] = "Ошибка в индексации ссылок."
        # Подсвечиваем ссылки, которые идут в неправильном порядке

    if unused_references:
        result["Неиспользованные ссылки из списка литературы"] = unused_references
        highlight_unused_references(doc, reference_paragraphs, unused_references)

    # Сохранение документа с примечаниями
    doc.save("output_with_notes.docx")
    for i in result:
        print(f"{i}:\n{result[i]}")


# Использование:
doc_path = "Бордунов Александр Максимович Отчет по работе.docx"
main(doc_path)
