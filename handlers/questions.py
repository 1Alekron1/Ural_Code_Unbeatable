import asyncio
from aiogram import Router, F, types
from aiogram.types import Message
from aiogram.fsm.context import FSMContext
from aiogram.types import FSInputFile

router = Router()

@router.message(F.document)
async def handle_document(message: Message, state: FSMContext):
    document = message.document

    # Показываем пользователю, что файл обрабатывается
    await message.answer("Файл получен! Начинается обработка...")

    # Загружаем файл
    file_info = await message.bot.get_file(document.file_id)
    downloaded_file = await message.bot.download_file(file_info.file_path)

    # Имитируем обработку файла (например, задержка на 5 секунд)
    await asyncio.sleep(5)

    # Здесь добавьте логику для обработки файла
    # Например, можно записать файл в локальную директорию и затем обработать его
    processed_file_path = "processed_file.txt"
    with open(processed_file_path, 'wb') as f:
        f.write(downloaded_file.read())  # Имитация сохранения файла

    # Показываем сообщение, что файл обработан
    await message.answer("Файл успешно обработан!")

    # Отправляем обработанный файл обратно
    processed_file = FSInputFile(processed_file_path)
    await message.answer_document(processed_file)
