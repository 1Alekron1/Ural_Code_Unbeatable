import asyncio
import os

from aiogram import Bot, Dispatcher
from dotenv import load_dotenv

from handlers import questions

load_dotenv()


async def main():
    bot = Bot(token=os.getenv('TOKEN'))
    dp = Dispatcher()

    dp.include_routers(questions.router)

    await bot.delete_webhook(drop_pending_updates=False)
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
