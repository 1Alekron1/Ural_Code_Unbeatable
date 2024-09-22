import asyncio
from io import BytesIO

from aiogram import Router, F, types
from aiogram.types import Message
from aiogram.fsm.context import FSMContext
from aiogram.types import FSInputFile
from docx import Document

from doc_file_an import check


router = Router()

def save_word_document(filename, content):
    document = Document()
    document.add_paragraph(content)
    document.save(filename)

@router.message(F.document)
async def handle_document(message: Message, state: FSMContext):
    document = message.document

    # Показываем пользователю, что файл обрабатывается
    await message.answer("Файл получен! Начинается обработка...")

    # Загружаем файл
    file_info = await message.bot.get_file(document.file_id)
    downloaded_file = await message.bot.download_file(file_info.file_path)

    result = check(downloaded_file)


    processed_file_path = "Исходник с правками.docx"

    formatted_message = ()


    # Отправляем обработанный файл обратно
    processed_file = FSInputFile(processed_file_path)
    await message.answer_document(processed_file)
    formatted_message = (
        "**Результаты проверки:**\n\n"
        "**Введение**: {Введение}\n"
        "**Заключение**: {Заключение}\n"
        "**Пересечение Введения и Заключения по смыслу**: {test}\n"
        "**Найдены разделы**: {Найдены_разделы}\n"
        "**Отсутствуют разделы**: {Отсутствуют_разделы}\n"
        "**Найденные ссылки в тексте**: {Найденные_ссылки_в_тексте}\n"
        "**Ошибка**: {Ошибка}\n"
        "**Неиспользованные ссылки из списка литературы**: {Неиспользованные_ссылки}"
    ).format(
        Введение=result['Введение'] or 'Не найдено',
        Заключение=result['Заключение'] or 'Не найдено',
        test=result["Пересечение Введения и Заключения по смыслу"] or 'Не найдено',
        Найдены_разделы=', '.join(result['Найдены разделы']) or 'Отсутствуют',
        Отсутствуют_разделы=', '.join(result['Отсутствуют разделы']) or 'Все разделы найдены',
        Найденные_ссылки_на_литературу=', '.join(result['Найденные ссылки на литературу']) or 'Не найдены',
        Найденные_ссылки_в_тексте=', '.join(map(str, result['Найденные ссылки в тексте'])) or 'Не найдены',
        Ошибка=result['Ошибка'] or 'Ошибок нет',
        Неиспользованные_ссылки=', '.join(
            map(str, result['Неиспользованные ссылки из списка литературы'])) or 'Нет неиспользованных ссылок'
    )

    text_file = open("Review.txt", "w")

    text_file.write(formatted_message)

    text_file.close()

    second_file = "Review.txt"
    processed_file_2 = FSInputFile(second_file)
    await message.answer_document(processed_file_2)
