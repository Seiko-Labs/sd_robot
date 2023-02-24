import time

from aiogram import Bot, Dispatcher, executor, types
from asyncio import sleep

from web import kill_browser, ExplorerOtbasy

API_TOKEN = "5923118349:AAERjGNhdPAc-ncR9iu6Kyyj7iHLBbtOgPs"

bot = Bot(token=API_TOKEN)
dp = Dispatcher(bot)
check = None


@dp.message_handler(commands=['start'])
async def start_reporting(message: types.Message):
    global check
    check = True
    while check:
        all_users = []
        try:
            kill_browser()
            explorer = ExplorerOtbasy()
            all_users = explorer.bpm_parse()
        except ValueError as e:
            if str(e) == 'Заявки не найдены':
                ...
        if len(all_users) > 0:
            for user in all_users:
                kill_browser()
                explorer = ExplorerOtbasy()
                try:
                    explorer.bpm_edit(user)
                    await message.reply("Заявки найдены")
                    break
                except ValueError:
                    continue
        else:
            await message.reply("Заявки не найдены")
        await sleep(300)


@dp.message_handler(commands=['stop'])
async def stop_reporting(message: types.Message):
    global check

if __name__ == '__main__':
    executor.start_polling(dp, skip_updates=True)
