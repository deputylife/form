import os
from aiogram import Bot, Dispatcher
from aiogram.types import ParseMode
from aiogram.utils.executor import start_polling

API_TOKEN = os.getenv("TELEGRAM_API_TOKEN")
bot = Bot(token=API_TOKEN)
dp = Dispatcher(bot)

async def on_startup(dp):
    print("Bot is starting up...")

async def on_shutdown(dp):
    print("Bot is shutting down...")

if __name__ == '__main__':
    from aiogram import executor
    executor.start_polling(dp, on_startup=on_startup, on_shutdown=on_shutdown)
