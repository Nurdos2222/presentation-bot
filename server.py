import os
import asyncio
from pathlib import Path

from aiohttp import web
from aiogram import Bot, Dispatcher, F
from aiogram.filters import CommandStart
from aiogram.types import (
    Message,
    ReplyKeyboardMarkup,
    KeyboardButton,
    FSInputFile,
)

from openai import AsyncOpenAI
from pptx import Presentation
from pptx.util import Inches, Pt


BOT_TOKEN = os.getenv("BOT_TOKEN")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

if not BOT_TOKEN:
    raise RuntimeError("BOT_TOKEN is not set")

if not OPENAI_API_KEY:
    raise RuntimeError("OPENAI_API_KEY is not set")


bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()
ai = AsyncOpenAI(api_key=OPENAI_API_KEY)

PORT = int(os.getenv("PORT", "10000"))

user_language = {}
user_state = {}


TEXT = {
    "kk": {
        "welcome": "👋 Қош келдіңіз!\n\nМен сізге әдемі презентациялар жасауға көмектесемін.",
        "create": "📝 Презентация жасау",
        "help": "📚 Көмек",
        "settings": "⚙️ Баптаулар",
        "stats": "📊 Статистика",
        "language": "🌐 Тіл",
        "back": "⬅️ Артқа",
        "auto": "✨ Автоматты дизайн",
        "design": "🎨 Дизайнды таңдау",
        "slides": "📊 Слайд саны",
        "ask_topic": "📝 Презентацияның тақырыбын жазыңыз.",
        "choose_slides": "📊 Слайд санын таңдаңыз:",
        "choose_design": "🎨 Дизайнды таңдаңыз:",
        "working": "⏳ Презентация дайындалып жатыр...",
        "done": "✅ Презентация дайын!",
        "error": "❌ Презентация жасау кезінде қате болды. Қайтадан көріңіз.",
        "help_text": (
            "📚 Көмек\n\n"
            "1. 📝 «Презентация жасау» батырмасын басыңыз.\n"
            "2. Тақырыпты жазыңыз.\n"
            "3. Слайд санын таңдаңыз.\n"
            "4. Дизайнды таңдаңыз немесе автоматты дизайнды қолданыңыз.\n"
            "5. 🤖 Бот презентацияны жасайды.\n"
            "6. 📎 Дайын PPTX файлын аласыз."
        ),
        "stats_text": "📊 Статистика\n\nСіз жасаған презентациялар: {count}",
    },

    "ru": {
        "welcome": "👋 Добро пожаловать!\n\nЯ помогу вам создавать красивые презентации.",
        "create": "📝 Создать презентацию",
        "help": "📚 Помощь",
        "settings": "⚙️ Настройки",
        "stats": "📊 Статистика",
        "language": "🌐 Язык",
        "back": "⬅️ Назад",
        "auto": "✨ Автоматический дизайн",
        "design": "🎨 Выбрать дизайн",
        "slides": "📊 Количество слайдов",
        "ask_topic": "📝 Напишите тему презентации.",
        "choose_slides": "📊 Выберите количество слайдов:",
        "choose_design": "🎨 Выберите дизайн:",
        "working": "⏳ Презентация создаётся...",
        "done": "✅ Презентация готова!",
        "error": "❌ Произошла ошибка при создании презентации. Попробуйте ещё раз.",
        "help_text": (
            "📚 Помощь\n\n"
            "1. 📝 Нажмите «Создать презентацию».\n"
            "2. Напишите тему.\n"
            "3. Выберите количество слайдов.\n"
            "4. Выберите дизайн или автоматический дизайн.\n"
            "5. 🤖 Бот создаст презентацию.\n"
            "6. 📎 Вы получите готовый PPTX-файл."
        ),
        "stats_text": "📊 Статистика\n\nСоздано презентаций: {count}",
    },

    "uz": {
        "welcome": "👋 Xush kelibsiz!\n\nMen sizga chiroyli taqdimotlar yaratishda yordam beraman.",
        "create": "📝 Taqdimot yaratish",
        "help": "📚 Yordam",
        "settings": "⚙️ Sozlamalar",
        "stats": "📊 Statistika",
        "language": "🌐 Til",
        "back": "⬅️ Orqaga",
        "auto": "✨ Avtomatik dizayn",
        "design": "🎨 Dizayn tanlash",
        "slides": "📊 Slaydlar soni",
        "ask_topic": "📝 Taqdimot mavzusini yozing.",
        "choose_slides": "📊 Slaydlar sonini tanlang:",
        "choose_design": "🎨 Dizaynni tanlang:",
        "working": "⏳ Taqdimot tayyorlanmoqda...",
        "done": "✅ Taqdimot tayyor!",
        "error": "❌ Taqdimot yaratishda xatolik yuz berdi. Qaytadan urinib ko‘ring.",
        "help_text": (
            "📚 Yordam\n\n"
            "1. 📝 «Taqdimot yaratish» tugmasini bosing.\n"
            "2. Mavzuni yozing.\n"
            "3. Slaydlar sonini tanlang.\n"
            "4. Dizaynni yoki avtomatik dizaynni tanlang.\n"
            "5. 🤖 Bot taqdimotni yaratadi.\n"
            "6. 📎 Tayyor PPTX faylini olasiz."
        ),
        "stats_text": "📊 Statistika\n\nYaratilgan taqdimotlar: {count}",
    },
}


def lang_of(user_id):
    return user_language.get(user_id, "ru")


def main_keyboard(lang):
    t = TEXT[lang]

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


def settings_keyboard(lang):
    t = TEXT[lang]

    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text=t["language"])],
            [KeyboardButton(text=t["back"])],
        ],
        resize_keyboard=True,
    )


def language_keyboard():
    return ReplyKeyboardMarkup(
        keyboard=[
            [
                KeyboardButton(text="🇰🇿 Қазақша"),
                KeyboardButton(text="🇷🇺 Русский"),
            ],
            [KeyboardButton(text="🇺🇿 O‘zbekcha")],
        ],
        resize_keyboard=True,
    )


def slides_keyboard(lang):
    t = TEXT[lang]

    return ReplyKeyboardMarkup(
        keyboard=[
            [
                KeyboardButton(text="5"),
                KeyboardButton(text="10"),
                KeyboardButton(text="15"),
            ],
            [KeyboardButton(text=t["back"])],
        ],
        resize_keyboard=True,
    )


def design_keyboard(lang):
    t = TEXT[lang]

    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text=t["auto"])],
            [
                KeyboardButton(text="✨ Modern"),
                KeyboardButton(text="🤍 Minimal"),
            ],
            [
                KeyboardButton(text="🌑 Dark"),
                KeyboardButton(text="🎓 Academic"),
            ],
            [KeyboardButton(text=t["back"])],
        ],
        resize_keyboard=True,
    )


def clean_filename(text):
    chars = '<>:"/\\|?*'
    for char in chars:
        text = text.replace(char, "")
    return text[:60].strip() or "presentation"


async def generate_content(topic, slides, language):
    language_name = {
        "kk": "Kazakh",
        "ru": "Russian",
        "uz": "Uzbek",
    }[language]

    prompt = f"""
Create a presentation in {language_name}.

Topic: {topic}
Number of slides: {slides}

Return ONLY valid JSON in this format:

{{
  "title": "Presentation title",
  "slides": [
    {{
      "title": "Slide title",
      "bullets": [
        "Bullet 1",
        "Bullet 2",
        "Bullet 3"
      ]
    }}
  ]
}}

Rules:
- Create exactly {slides} slides.
- Keep the information educational and useful.
- Each slide should have 3-5 short bullet points.
- Do not use markdown.
- Do not add explanations outside JSON.
"""

    response = await ai.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {
                "role": "system",
                "content": "You create structured educational presentations.",
            },
            {
                "role": "user",
                "content": prompt,
            },
        ],
        temperature=0.7,
    )

    import json

    return json.loads(response.choices[0].message.content)


def add_slide(prs, title, bullets, design):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    width = prs.slide_width
    height = prs.slide_height

    # Background
    if design == "🌑 Dark":
        background = slide.background
        background.fill.solid()
        background.fill.fore_color.rgb = __import__(
            "pptx"
        ).dml.color.RGBColor(25, 25, 30)

    # Title
    title_box = slide.shapes.add_textbox(
        Inches(0.7),
        Inches(0.5),
        Inches(12),
        Inches(1),
    )

    title_frame = title_box.text_frame
    title_frame.text = title

    for paragraph in title_frame.paragraphs:
        paragraph.font.size = Pt(28)
        paragraph.font.bold = True

    # Content
    body = slide.shapes.add_textbox(
        Inches(0.9),
        Inches(1.7),
        Inches(11.5),
        Inches(5),
    )

    frame = body.text_frame
    frame.word_wrap = True

    for i, bullet in enumerate(bullets):
        paragraph = frame.paragraphs[0] if i == 0 else frame.add_paragraph()
        paragraph.text = "• " + bullet
        paragraph.font.size = Pt(20)
        paragraph.space_after = Pt(12)


def create_presentation(data, design, filename):
    prs = Presentation()

    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # Title slide
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    title = slide.shapes.add_textbox(
        Inches(1),
        Inches(2.3),
        Inches(11.3),
        Inches(1.5),
    )

    frame = title.text_frame
    frame.text = data["title"]

    for p in frame.paragraphs:
        p.font.size = Pt(38)
        p.font.bold = True

    # Content slides
    for item in data["slides"]:
        add_slide(
            prs,
            item["title"],
            item["bullets"],
            design,
        )

    prs.save(filename)


@dp.message(CommandStart())
async def start(message: Message):
    user_id = message.from_user.id
    lang = lang_of(user_id)

    await message.answer(
        TEXT[lang]["welcome"],
        reply_markup=main_keyboard(lang),
    )


@dp.message(F.text == "🇰🇿 Қазақша")
async def kazakh(message: Message):
    user_language[message.from_user.id] = "kk"

    await message.answer(
        TEXT["kk"]["welcome"],
        reply_markup=main_keyboard("kk"),
    )


@dp.message(F.text == "🇷🇺 Русский")
async def russian(message: Message):
    user_language[message.from_user.id] = "ru"

    await message.answer(
        TEXT["ru"]["welcome"],
        reply_markup=main_keyboard("ru"),
    )


@dp.message(F.text == "🇺🇿 O‘zbekcha")
async def uzbek(message: Message):
    user_language[message.from_user.id] = "uz"

    await message.answer(
        TEXT["uz"]["welcome"],
        reply_markup=main_keyboard("uz"),
    )


@dp.message()
async def all_messages(message: Message):
    user_id = message.from_user.id
    lang = lang_of(user_id)
    t = TEXT[lang]
    text = message.text or ""

    state = user_state.get(user_id, {})

    if text == t["help"]:
        await message.answer(t["help_text"])
        return

    if text == t["settings"]:
        await message.answer(
            "⚙️",
            reply_markup=settings_keyboard(lang),
        )
        return

    if text == t["language"]:
        await message.answer(
            "🌐",
            reply_markup=language_keyboard(),
        )
        return

    if text == t["stats"]:
        count = state.get("count", 0)

        await message.answer(
            t["stats_text"].format(count=count)
        )
        return

    if text == t["back"]:
        user_state.pop(user_id, None)

        await message.answer(
            t["welcome"],
            reply_markup=main_keyboard(lang),
        )
        return

    if text == t["create"]:
        user_state[user_id] = {
            "step": "topic",
            "count": state.get("count", 0),
        }

        await message.answer(
            t["ask_topic"],
            reply_markup=ReplyKeyboardMarkup(
                keyboard=[
                    [KeyboardButton(text=t["back"])]
                ],
                resize_keyboard=True,
            ),
        )
        return

    if state.get("step") == "topic":
        state["topic"] = text
        state["step"] = "slides"

        await message.answer(
            t["choose_slides"],
            reply_markup=slides_keyboard(lang),
        )
        return

    if state.get("step") == "slides" and text in {"5", "10", "15"}:
        state["slides"] = int(text)
        state["step"] = "design"

        await message.answer(
            t["choose_design"],
            reply_markup=design_keyboard(lang),
        )
        return

    if state.get("step") == "design":
        state["design"] = text

        await message.answer(
            t["working"]
        )

        try:
            data = await generate_content(
                state["topic"],
                state["slides"],
                lang,
            )

            filename = (
                Path("/tmp")
                / f"{clean_filename(data['title'])}.pptx"
            )

            create_presentation(
                data,
                state["design"],
                str(filename),
            )

            await message.answer_document(
                FSInputFile(str(filename)),
                caption=t["done"],
            )

            state["count"] = state.get("count", 0) + 1
            user_state[user_id] = state

            try:
                filename.unlink()
            except Exception:
                pass

        except Exception:
            await message.answer(t["error"])

        return

    await message.answer(
        t["welcome"],
        reply_markup=main_keyboard(lang),
    )


async def health(request):
    return web.Response(text="Presentation Bot is running!")


async def run_web():
    app = web.Application()
    app.router.add_get("/", health)

    runner = web.AppRunner(app)
    await runner.setup()

    site = web.TCPSite(
        runner,
        "0.0.0.0",
        PORT,
    )

    await site.start()


async def main():
    await run_web()
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
