import os
import asyncio
from aiohttp import web
from aiogram import Bot, Dispatcher, F
from aiogram.types import Message, FSInputFile, InlineKeyboardMarkup, InlineKeyboardButton
from aiogram.filters import CommandStart
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE


BOT_TOKEN = os.getenv("BOT_TOKEN")

if not BOT_TOKEN:
    raise RuntimeError("BOT_TOKEN is not set")


bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()


# =========================
# USER DATA
# =========================

users = {}

DEFAULT_LANG = "ru"
DEFAULT_SLIDES = 5
DEFAULT_DESIGN = "modern"


# =========================
# TEXTS
# =========================

TEXT = {
    "ru": {
        "welcome": "👋 Добро пожаловать!\n\nЯ помогу создать презентацию PowerPoint.\n\nВыберите действие:",
        "create": "📝 Создать презентацию",
        "help": "📚 Помощь",
        "settings": "⚙️ Настройки",
        "stats": "📊 Статистика",
        "language": "🌐 Язык",
        "slides": "📑 Количество слайдов",
        "design": "🎨 Дизайн",
        "back": "⬅️ Назад",
        "choose_lang": "🌐 Выберите язык:",
        "choose_slides": "📑 Выберите количество слайдов:",
        "choose_design": "🎨 Выберите дизайн:",
        "ask_topic": "📝 Напишите тему презентации:",
        "creating": "⏳ Создаю презентацию...",
        "ready": "✅ Презентация готова!",
        "help_text": (
            "📚 Как пользоваться ботом:\n\n"
            "1️⃣ Нажмите «📝 Создать презентацию».\n"
            "2️⃣ Выберите количество слайдов.\n"
            "3️⃣ Выберите дизайн.\n"
            "4️⃣ Напишите тему.\n"
            "5️⃣ Подождите несколько секунд.\n"
            "6️⃣ Получите готовый PowerPoint-файл 📥"
        ),
        "stats_text": "📊 Ваша статистика:\n\n🎬 Создано презентаций: {count}",
        "error": "❌ Произошла ошибка. Попробуйте ещё раз.",
        "modern": "✨ Modern",
        "minimal": "🤍 Minimal",
        "dark": "🌑 Dark",
        "academic": "🎓 Academic",
    },

    "kk": {
        "welcome": "👋 Қош келдіңіз!\n\nМен PowerPoint презентациясын жасауға көмектесемін.\n\nӘрекетті таңдаңыз:",
        "create": "📝 Презентация жасау",
        "help": "📚 Көмек",
        "settings": "⚙️ Баптаулар",
        "stats": "📊 Статистика",
        "language": "🌐 Тіл",
        "slides": "📑 Слайд саны",
        "design": "🎨 Дизайн",
        "back": "⬅️ Артқа",
        "choose_lang": "🌐 Тілді таңдаңыз:",
        "choose_slides": "📑 Слайд санын таңдаңыз:",
        "choose_design": "🎨 Дизайнды таңдаңыз:",
        "ask_topic": "📝 Презентация тақырыбын жазыңыз:",
        "creating": "⏳ Презентация жасалуда...",
        "ready": "✅ Презентация дайын!",
        "help_text": (
            "📚 Ботты пайдалану:\n\n"
            "1️⃣ «📝 Презентация жасау» батырмасын басыңыз.\n"
            "2️⃣ Слайд санын таңдаңыз.\n"
            "3️⃣ Дизайнды таңдаңыз.\n"
            "4️⃣ Тақырыпты жазыңыз.\n"
            "5️⃣ Бірнеше секунд күтіңіз.\n"
            "6️⃣ Дайын PowerPoint файлын алыңыз 📥"
        ),
        "stats_text": "📊 Статистикаңыз:\n\n🎬 Жасалған презентациялар: {count}",
        "error": "❌ Қате пайда болды. Қайтадан көріңіз.",
        "modern": "✨ Modern",
        "minimal": "🤍 Minimal",
        "dark": "🌑 Dark",
        "academic": "🎓 Academic",
    },

    "uz": {
        "welcome": "👋 Xush kelibsiz!\n\nMen PowerPoint taqdimotini yaratishga yordam beraman.\n\nAmalni tanlang:",
        "create": "📝 Taqdimot yaratish",
        "help": "📚 Yordam",
        "settings": "⚙️ Sozlamalar",
        "stats": "📊 Statistika",
        "language": "🌐 Til",
        "slides": "📑 Slaydlar soni",
        "design": "🎨 Dizayn",
        "back": "⬅️ Orqaga",
        "choose_lang": "🌐 Tilni tanlang:",
        "choose_slides": "📑 Slaydlar sonini tanlang:",
        "choose_design": "🎨 Dizaynni tanlang:",
        "ask_topic": "📝 Taqdimot mavzusini yozing:",
        "creating": "⏳ Taqdimot yaratilmoqda...",
        "ready": "✅ Taqdimot tayyor!",
        "help_text": (
            "📚 Botdan foydalanish:\n\n"
            "1️⃣ «📝 Taqdimot yaratish» tugmasini bosing.\n"
            "2️⃣ Slaydlar sonini tanlang.\n"
            "3️⃣ Dizaynni tanlang.\n"
            "4️⃣ Mavzuni yozing.\n"
            "5️⃣ Bir necha soniya kuting.\n"
            "6️⃣ Tayyor PowerPoint faylini oling 📥"
        ),
        "stats_text": "📊 Statistikangiz:\n\n🎬 Yaratilgan taqdimotlar: {count}",
        "error": "❌ Xatolik yuz berdi. Qayta urinib ko‘ring.",
        "modern": "✨ Modern",
        "minimal": "🤍 Minimal",
        "dark": "🌑 Dark",
        "academic": "🎓 Academic",
    }
}


# =========================
# KEYBOARDS
# =========================

def main_keyboard(lang):
    t = TEXT[lang]

    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text=t["create"], callback_data="create")],
            [
                InlineKeyboardButton(text=t["help"], callback_data="help"),
                InlineKeyboardButton(text=t["settings"], callback_data="settings")
            ],
            [InlineKeyboardButton(text=t["stats"], callback_data="stats")]
        ]
    )


def settings_keyboard(lang):
    t = TEXT[lang]

    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text=t["language"], callback_data="language")],
            [InlineKeyboardButton(text=t["slides"], callback_data="slides")],
            [InlineKeyboardButton(text=t["design"], callback_data="design")],
            [InlineKeyboardButton(text=t["back"], callback_data="back")]
        ]
    )


def language_keyboard(lang):
    t = TEXT[lang]

    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(text="🇰🇿 Қазақша", callback_data="lang_kk"),
                InlineKeyboardButton(text="🇷🇺 Русский", callback_data="lang_ru")
            ],
            [
                InlineKeyboardButton(text="🇺🇿 O'zbekcha", callback_data="lang_uz")
            ],
            [InlineKeyboardButton(text=t["back"], callback_data="settings")]
        ]
    )


def slides_keyboard(lang):
    t = TEXT[lang]

    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(text="5️⃣", callback_data="slides_5"),
                InlineKeyboardButton(text="🔟", callback_data="slides_10"),
                InlineKeyboardButton(text="1️⃣5️⃣", callback_data="slides_15")
            ],
            [InlineKeyboardButton(text=t["back"], callback_data="settings")]
        ]
    )


def design_keyboard(lang):
    t = TEXT[lang]

    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text=t["modern"], callback_data="design_modern")],
            [InlineKeyboardButton(text=t["minimal"], callback_data="design_minimal")],
            [InlineKeyboardButton(text=t["dark"], callback_data="design_dark")],
            [InlineKeyboardButton(text=t["academic"], callback_data="design_academic")],
            [InlineKeyboardButton(text=t["back"], callback_data="settings")]
        ]
    )


# =========================
# SLIDE CONTENT
# =========================

def generate_content(topic, count, lang):
    if lang == "kk":
        titles = [
            topic,
            f"{topic}: негізгі түсініктер",
            f"{topic}: маңызды ерекшеліктер",
            f"{topic}: негізгі бағыттар",
            f"{topic}: артықшылықтары",
            f"{topic}: қазіргі маңызы",
            f"{topic}: қызықты фактілер",
            f"{topic}: мысалдар",
            f"{topic}: қорытынды",
            "Назарларыңызға рақмет!"
        ]

        texts = [
            f"{topic} тақырыбына арналған презентация.",
            f"{topic} туралы негізгі түсініктер мен ақпарат.",
            f"Тақырыптың маңызды ерекшеліктері мен негізгі сипаттамалары.",
            f"{topic} бойынша негізгі бағыттар мен қолданылу салалары.",
            f"{topic} тақырыбының негізгі артықшылықтары.",
            f"Қазіргі уақытта {topic} тақырыбының маңызы.",
            f"{topic} туралы қызықты фактілер мен мәліметтер.",
            f"{topic} бойынша мысалдар мен практикалық ақпарат.",
            f"{topic} бойынша негізгі қорытындылар.",
            "Презентация аяқталды."
        ]

    elif lang == "uz":
        titles = [
            topic,
            f"{topic}: asosiy tushunchalar",
            f"{topic}: muhim xususiyatlar",
            f"{topic}: asosiy yo‘nalishlar",
            f"{topic}: afzalliklari",
            f"{topic}: zamonaviy ahamiyati",
            f"{topic}: qiziqarli faktlar",
            f"{topic}: misollar",
            f"{topic}: xulosa",
            "E’tiboringiz uchun rahmat!"
        ]

        texts = [
            f"{topic} mavzusiga bag‘ishlangan taqdimot.",
            f"{topic} haqida asosiy tushunchalar va ma’lumotlar.",
            f"Mavzuning muhim xususiyatlari va asosiy jihatlari.",
            f"{topic} bo‘yicha asosiy yo‘nalishlar va qo‘llanish sohalari.",
            f"{topic} mavzusining asosiy afzalliklari.",
            f"Bugungi kunda {topic} mavzusining ahamiyati.",
            f"{topic} haqida qiziqarli faktlar va ma’lumotlar.",
            f"{topic} bo‘yicha misollar va amaliy ma’lumotlar.",
            f"{topic} bo‘yicha asosiy xulosalar.",
            "Taqdimot yakunlandi."
        ]

    else:
        titles = [
            topic,
            f"{topic}: основные понятия",
            f"{topic}: важные особенности",
            f"{topic}: основные направления",
            f"{topic}: преимущества",
            f"{topic}: современное значение",
            f"{topic}: интересные факты",
            f"{topic}: примеры",
            f"{topic}: вывод",
            "Спасибо за внимание!"
        ]

        texts = [
            f"Презентация на тему «{topic}».",
            f"Основные понятия и информация о теме «{topic}».",
            f"Важные особенности и ключевые характеристики темы.",
            f"Основные направления и области применения.",
            f"Главные преимущества и особенности.",
            f"Современное значение данной темы.",
            f"Интересные факты и полезная информация.",
            f"Практические примеры по теме.",
            f"Основные выводы по теме «{topic}».",
            "Презентация завершена."
        ]

    result = []

    for i in range(count):
        result.append((titles[i % len(titles)], texts[i % len(texts)]))

    return result


# =========================
# PPTX DESIGN
# =========================

def add_text(slide, text, x, y, w, h, size, bold=False, color=None):
    box = slide.shapes.add_textbox(
        Inches(x),
        Inches(y),
        Inches(w),
        Inches(h)
    )

    tf = box.text_frame
    tf.clear()

    p = tf.paragraphs[0]
    p.text = text
    p.alignment = PP_ALIGN.LEFT

    run = p.runs[0]
    run.font.size = Pt(size)
    run.font.bold = bold

    if color:
        run.font.color.rgb = RGBColor(*color)

    return box


def add_decoration(slide, design):
    if design == "modern":
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(0),
            Inches(0),
            Inches(0.18),
            Inches(7.5)
        )
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(60, 90, 200)
        shape.line.fill.background()

    elif design == "dark":
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(0),
            Inches(0),
            Inches(13.33),
            Inches(0.18)
        )
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(120, 90, 220)
        shape.line.fill.background()

    elif design == "academic":
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(0),
            Inches(0),
            Inches(13.33),
            Inches(0.12)
        )
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(40, 70, 120)
        shape.line.fill.background()


def create_presentation(topic, count, design, lang, filename):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    slides = generate_content(topic, count, lang)

    if design == "dark":
        bg = (25, 25, 35)
        title_color = (255, 255, 255)
        text_color = (220, 220, 230)

    elif design == "academic":
        bg = (245, 247, 250)
        title_color = (30, 55, 100)
        text_color = (60, 60, 70)

    elif design == "minimal":
        bg = (250, 250, 250)
        title_color = (35, 35, 35)
        text_color = (80, 80, 80)

    else:
        bg = (245, 248, 255)
        title_color = (35, 55, 100)
        text_color = (70, 75, 90)

    for index, (title, body) in enumerate(slides):

        slide = prs.slides.add_slide(prs.slide_layouts[6])

        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(*bg)

        add_decoration(slide, design)

        if index == 0:
            add_text(
                slide,
                title,
                1.0,
                2.25,
                11.3,
                1.4,
                38,
                True,
                title_color
            )

            add_text(
                slide,
                body,
                1.05,
                3.75,
                10.8,
                1.0,
                20,
                False,
                text_color
            )

        else:
            add_text(
                slide,
                title,
                0.9,
                0.75,
                11.5,
                1.0,
                30,
                True,
                title_color
            )

            # Decorative line
            line = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                Inches(0.9),
                Inches(1.75),
                Inches(2.0),
                Inches(0.08)
            )
            line.fill.solid()
            line.fill.fore_color.rgb = RGBColor(80, 110, 200)
            line.line.fill.background()

            add_text(
                slide,
                "• " + body,
                1.0,
                2.25,
                11.0,
                2.5,
                22,
                False,
                text_color
            )

            add_text(
                slide,
                str(index + 1),
                11.9,
                6.65,
                0.5,
                0.4,
                12,
                False,
                text_color
            )

    prs.save(filename)


# =========================
# USER STATE
# =========================

def get_user(user_id):
    if user_id not in users:
        users[user_id] = {
            "lang": DEFAULT_LANG,
            "slides": DEFAULT_SLIDES,
            "design": DEFAULT_DESIGN,
            "count": 0,
            "state": None
        }

    return users[user_id]


# =========================
# START
# =========================

@dp.message(CommandStart())
async def start(message: Message):
    user = get_user(message.from_user.id)
    lang = user["lang"]

    await message.answer(
        TEXT[lang]["welcome"],
        reply_markup=main_keyboard(lang)
    )


# =========================
# CREATE
# =========================

@dp.callback_query(F.data == "create")
async def create_start(callback):
    user = get_user(callback.from_user.id)
    lang = user["lang"]

    user["state"] = "topic"

    await callback.message.edit_text(
        TEXT[lang]["ask_topic"]
    )

    await callback.answer()


@dp.message()
async def messages(message: Message):
    user = get_user(message.from_user.id)

    if user.get("state") != "topic":
        return

    topic = message.text.strip()

    if not topic:
        return

    user["state"] = None
    lang = user["lang"]

    status = await message.answer(
        TEXT[lang]["creating"]
    )

    filename = f"presentation_{message.from_user.id}.pptx"

    try:
        await asyncio.to_thread(
            create_presentation,
            topic,
            user["slides"],
            user["design"],
            lang,
            filename
        )

        user["count"] += 1

        await status.delete()

        await message.answer_document(
            FSInputFile(filename),
            caption=TEXT[lang]["ready"]
        )

        try:
            os.remove(filename)
        except OSError:
            pass

    except Exception as e:
        print("PPTX ERROR:", e)

        await status.edit_text(
            TEXT[lang]["error"]
        )


# =========================
# HELP
# =========================

@dp.callback_query(F.data == "help")
async def help_callback(callback):
    user = get_user(callback.from_user.id)
    lang = user["lang"]

    await callback.message.edit_text(
        TEXT[lang]["help_text"],
        reply_markup=InlineKeyboardMarkup(
            inline_keyboard=[
                [
                    InlineKeyboardButton(
                        text=TEXT[lang]["back"],
                        callback_data="back"
                    )
                ]
            ]
        )
    )

    await callback.answer()


# =========================
# SETTINGS
# =========================

@dp.callback_query(F.data == "settings")
async def settings(callback):
    user = get_user(callback.from_user.id)
    lang = user["lang"]

    await callback.message.edit_text(
        "⚙️",
        reply_markup=settings_keyboard(lang)
    )

    await callback.answer()


# =========================
# LANGUAGE
# =========================

@dp.callback_query(F.data == "language")
async def language(callback):
    user = get_user(callback.from_user.id)
    lang = user["lang"]

    await callback.message.edit_text(
        TEXT[lang]["choose_lang"],
        reply_markup=language_keyboard(lang)
    )

    await callback.answer()


@dp.callback_query(F.data.startswith("lang_"))
async def set_language(callback):
    user = get_user(callback.from_user.id)

    lang = callback.data.replace("lang_", "")

    if lang not in TEXT:
        lang = "ru"

    user["lang"] = lang

    await callback.message.edit_text(
        TEXT[lang]["welcome"],
        reply_markup=main_keyboard(lang)
    )

    await callback.answer()


# =========================
# SLIDES
# =========================

@dp.callback_query(F.data == "slides")
async def slides(callback):
    user = get_user(callback.from_user.id)
    lang = user["lang"]

    await callback.message.edit_text(
        TEXT[lang]["choose_slides"],
        reply_markup=slides_keyboard(lang)
    )

    await callback.answer()


@dp.callback_query(F.data.startswith("slides_"))
async def set_slides(callback):
    user = get_user(callback.from_user.id)

    number = int(callback.data.replace("slides_", ""))

    user["slides"] = number

    lang = user["lang"]

    await callback.message.edit_text(
        TEXT[lang]["choose_slides"],
        reply_markup=slides_keyboard(lang)
    )

    await callback.answer("✅")


# =========================
# DESIGN
# =========================

@dp.callback_query(F.data == "design")
async def design(callback):
    user = get_user(callback.from_user.id)
    lang = user["lang"]

    await callback.message.edit_text(
        TEXT[lang]["choose_design"],
        reply_markup=design_keyboard(lang)
    )

    await callback.answer()


@dp.callback_query(F.data.startswith("design_"))
async def set_design(callback):
    user = get_user(callback.from_user.id)

    design_name = callback.data.replace("design_", "")

    user["design"] = design_name

    lang = user["lang"]

    await callback.message.edit_text(
        TEXT[lang]["choose_design"],
        reply_markup=design_keyboard(lang)
    )

    await callback.answer("✅")


# =========================
