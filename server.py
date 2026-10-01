import os
import asyncio
from aiohttp import web
from aiogram import Bot, Dispatcher, types
from aiogram.filters import CommandStart
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton

TOKEN = os.getenv("BOT_TOKEN")

if not TOKEN:
    raise RuntimeError("BOT_TOKEN is not set")

bot = Bot(token=TOKEN)
dp = Dispatcher()

TEXTS = {
    "kk": {
        "welcome": "👋 Қош келдіңіз!\n\nМен сізге әдемі презентациялар жасауға көмектесемін.",
        "create": "📝 Презентация жасау",
        "help": "📚 Көмек",
        "settings": "⚙️ Баптаулар",
        "stats": "📊 Статистика",
        "help_text": (
            "📚 Көмек\n\n"
            "1. 📝 «Презентация жасау» батырмасын басыңыз.\n"
            "2. Презентация тақырыбын жазыңыз.\n"
            "3. Слайд санын таңдаңыз.\n"
            "4. Дизайнды таңдаңыз немесе автоматты дизайнды қолданыңыз.\n"
            "5. 🤖 Бот презентацияны дайындайды.\n"
            "6. 📎 Дайын .pptx файлын аласыз."
        ),
    },
    "ru": {
        "welcome": "👋 Добро пожаловать!\n\nЯ помогу вам создавать красивые презентации.",
        "create": "📝 Создать презентацию",
        "help": "📚 Помощь",
        "settings": "⚙️ Настройки",
        "stats": "📊 Статистика",
        "help_text": (
            "📚 Помощь\n\n"
            "1. 📝 Нажмите «Создать презентацию».\n"
            "2. Напишите тему презентации.\n"
            "3. Выберите количество слайдов.\n"
            "4. Выберите дизайн или автоматический дизайн.\n"
            "5. 🤖 Бот подготовит презентацию.\n"
            "6. 📎 Вы получите готовый файл .pptx."
        ),
    },
    "uz": {
        "welcome": "👋 Xush kelibsiz!\n\nMen sizga chiroyli taqdimotlar yaratishda yordam beraman.",
        "create": "📝 Taqdimot yaratish",
        "help": "📚 Yordam",
        "settings": "⚙️ Sozlamalar",
        "stats": "📊 Statistika",
        "help_text": (
            "📚 Yordam\n\n"
            "1. 📝 «Taqdimot yaratish» tugmasini bosing.\n"
            "2. Taqdimot mavzusini yozing.\n"
            "3. Slaydlar sonini tanlang.\n"
            "4. Dizaynni tanlang yoki avtomatik dizayndan foydalaning.\n"
            "5. 🤖 Bot taqdimotni tayyorlaydi.\n"
            "6. 📎 Tayyor .pptx faylini olasiz."
        ),
    },
}

user_lang = {}


def get_lang(user_id):
    return user_lang.get(user_id, "ru")


def keyboard(lang):
    t = TEXTS[lang]

    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text=t["create"])],
            [KeyboardButton(text=t["help"])],
            [
                KeyboardButton(text=t["settings"]),
                KeyboardButton(text=t["stats"]),
            ],
        ],
        resize_keyboard=True,
    )


@dp.message(CommandStart())
async def start(message: types.Message):
    lang = get_lang(message.from_user.id)

    await message.answer(
        TEXTS[lang]["welcome"],
        reply_markup=keyboard(lang),
    )


@dp.message()
async def messages(message: types.Message):
    lang = get_lang(message.from_user.id)
    t = TEXTS[lang]

    if message.text == t["help"]:
        await message.answer(t["help_text"])

    elif message.text == t["create"]:
        await message.answer(
            "📝 Жақында бұл жерде тақырыпты жазып, дайын презентация аласыз."
            if lang == "kk"
            else
            "📝 Скоро здесь вы сможете написать тему и получить готовую презентацию."
            if lang == "ru"
            else
            "📝 Tez orada bu yerda mavzuni yozib, tayyor taqdimot olishingiz mumkin."
        )

    elif message.text == t["settings"]:
        await message.answer("⚙️ Настройки / Баптаулар / Sozlamalar")

    elif message.text == t["stats"]:
        await message.answer("📊 Статистика пока пуста.")

    else:
        await message.answer(
            "Түсіндім. Презентация жасау үшін 📝 батырмасын басыңыз."
            if lang == "kk"
            else
            "Пожалуйста, выберите действие из меню."
            if lang == "ru"
            else
            "Iltimos, menyudan kerakli bo‘limni tanlang."
        )


async def health(request):
    return web.Response(text="Presentation Bot is running!")


async def start_web():
    app = web.Application()
    app.router.add_get("/", health)

    runner = web.AppRunner(app)
    await runner.setup()

    port = int(os.getenv("PORT", "10000"))

    site = web.TCPSite(runner, "0.0.0.0", port)
    await site.start()


async def main():
    await start_web()
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
