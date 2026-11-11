"""Сборка presentation/psihologia_shkoly_adler_jung.pptx.

Слайд 1280x720 px (1 px = 9525 EMU), тёплая палитра как в index.html.
Шрифты: Georgia (заголовки, цитаты), Segoe UI (тело).
"""
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from copy import deepcopy

PX = 9525  # EMU на пиксель при 96 dpi
W, H = 1280, 720
PADX, PADY = 80, 64

BG          = RGBColor(0xFD, 0xFC, 0xF8)
INK         = RGBColor(0x1C, 0x19, 0x17)
INK_SOFT    = RGBColor(0x3F, 0x3A, 0x36)
MUTED       = RGBColor(0x6E, 0x67, 0x63)
LINE        = RGBColor(0xE5, 0xE1, 0xD8)
LINE_STRONG = RGBColor(0xD6, 0xD0, 0xC2)
ADLER       = RGBColor(0xB4, 0x53, 0x09)
ADLER_DARK  = RGBColor(0x9A, 0x4A, 0x06)
ADLER_TINT  = RGBColor(0xFD, 0xF0, 0xDF)
JUNG        = RGBColor(0x4C, 0x35, 0x75)
JUNG_TINT   = RGBColor(0xED, 0xE7, 0xF5)
L1_BG       = RGBColor(0xFA, 0xF7, 0xEF)
L2_BG       = RGBColor(0xEF, 0xE7, 0xD5)
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)

SERIF = "Georgia"
SANS  = "Segoe UI"


def px(v):
    return Emu(int(round(v * PX)))


def new_deck():
    prs = Presentation()
    prs.slide_width = px(W)
    prs.slide_height = px(H)
    return prs


def add_slide(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])  # пустой макет
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = BG
    return s


def rect(slide, x, y, w, h, fill=None, line=None, lw=1.0, shape=MSO_SHAPE.RECTANGLE):
    sh = slide.shapes.add_shape(shape, px(x), px(y), px(w), px(h))
    if fill is None:
        sh.fill.background()
    else:
        sh.fill.solid()
        sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(lw)
    sh.shadow.inherit = False
    return sh


def hline(slide, x, y, w, color=LINE, lw=1.0):
    """Горизонтальная линия толщиной lw pt."""
    return rect(slide, x, y, w, max(1.0, lw * 96 / 72), fill=color)


def vline(slide, x, y, h, color=LINE, lw=1.0):
    return rect(slide, x, y, max(1.0, lw * 96 / 72), h, fill=color)


def tb(slide, x, y, w, h):
    box = slide.shapes.add_textbox(px(x), px(y), px(w), px(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.TOP
    # отключаем автоусадку — размеры задаём сами
    bodyPr = tf._txBody.find(qn('a:bodyPr'))
    for tag in ('a:normAutofit', 'a:spAutoFit'):
        el = bodyPr.find(qn(tag))
        if el is not None:
            bodyPr.remove(el)
    return tf


def _spc(run, px_amount):
    """Межбуквенный интервал: px -> 1/100 pt."""
    run.font._rPr.set('spc', str(int(round(px_amount * 72 / 96 * 100))))


def para(tf, text, size, color, font=SANS, bold=False, italic=False,
         align=PP_ALIGN.LEFT, line=1.25, before=0, after=0, caps=False,
         spacing=None, first=False):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align
    p.line_spacing = line
    p.space_before = Pt(before)
    p.space_after = Pt(after)
    run = p.add_run()
    run.text = text.upper() if caps else text
    run.font.size = Pt(size * 72 / 96)      # px -> pt
    run.font.name = font
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    if spacing:
        _spc(run, spacing)
    return p


def rich(tf, parts, size, color, font=SANS, align=PP_ALIGN.LEFT, line=1.25,
         before=0, after=0, first=False):
    """Абзац из нескольких прогонов: parts = [(текст, bold, italic, цвет|None)]."""
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align
    p.line_spacing = line
    p.space_before = Pt(before)
    p.space_after = Pt(after)
    for text, bold, italic, col in parts:
        run = p.add_run()
        run.text = text
        run.font.size = Pt(size * 72 / 96)
        run.font.name = font
        run.font.bold = bold
        run.font.italic = italic
        run.font.color.rgb = col or color
    return p


# ─────────────────────────── слайд 1 · титул ───────────────────────────
def s01(prs):
    s = add_slide(prs)
    t = tb(s, PADX, PADY, 700, 24)
    para(t, "Психология школы · Тема №8", 12, MUTED, bold=True,
         caps=True, spacing=3.0, first=True)
    t = tb(s, 1036, PADY - 14, 160, 60)
    para(t, "08", 52, MUTED, font=SERIF, italic=True,
         align=PP_ALIGN.RIGHT, line=1.0, first=True)

    t = tb(s, PADX, 200, 640, 210)
    para(t, "Психология", 94, INK, font=SERIF, line=0.95, first=True)
    para(t, "школы", 94, MUTED, font=SERIF, italic=True, line=0.95)
    t = tb(s, PADX, 430, 640, 90)
    para(t, "в подходах Альфреда Адлера", 22, INK_SOFT, font=SERIF, italic=True,
         line=1.4, first=True)
    para(t, "и Карла Густава Юнга", 22, INK_SOFT, font=SERIF, italic=True, line=1.4)

    vline(s, 747, 218, 250)
    for i, (letter, col, who, what) in enumerate([
        ("А", ADLER, "Альфред Адлер", "индивидуальная психология"),
        ("Ю", JUNG, "Карл Густав Юнг", "аналитическая психология"),
    ]):
        y = 222 + i * 122
        t = tb(s, 781, y, 90, 76)
        para(t, letter, 70, col, font=SERIF, italic=True, line=0.85, first=True)
        t = tb(s, 866, y + 12, 330, 60)
        para(t, who, 17, INK, bold=True, line=1.2, first=True)
        para(t, what, 13, MUTED, line=1.3, before=3)

    hline(s, PADX, 636, 1120)
    t = tb(s, PADX, 646, 700, 26)
    para(t, "Сабадаш Павел Андреевич · группа ИС-24", 14, INK, bold=True,
         first=True)
    t = tb(s, 700, 646, 500, 26)
    para(t, "РКСИ · Ростов-на-Дону · 2026", 14, MUTED,
         align=PP_ALIGN.RIGHT, first=True)


# ─────────────────────────── слайд 2 · контекст ───────────────────────────
def s02(prs):
    s = add_slide(prs)
    head(s, "Контекст", "Почему школа стала объектом психологии",
         "Начало XX века: массовая школа, детство и бессознательное", None)
    left = tb(s, PADX, 200, 611, 400)
    for i, txt in enumerate([
        "К началу XX века в Европе складывается массовая обязательная школа. "
        "Впервые огромное число детей одновременно оказывается в одном "
        "пространстве, подчиняется общим правилам и сравнивается по единой мерке.",
        "Школа становится первым институтом после семьи, где ребёнок "
        "систематически переживает успех и неудачу, сталкивается с авторитетом "
        "и оценкой. Именно здесь формируются его представления о себе и о своём "
        "месте среди других.",
        "Психоанализ выходит за пределы клиники. Интерес к детству и "
        "бессознательным мотивам переносится на педагогику. Адлер и Юнг — оба "
        "из круга Фрейда — предлагают свои ответы на вопрос: что происходит с "
        "личностью ребёнка в школе.",
    ]):
        para(left, txt, 16, INK_SOFT, line=1.65, after=14, first=(i == 0))

    vline(s, 747, 200, 360)
    items = [("1900-е", ADLER, "Массовая школа в Европе"),
             ("1911", ADLER, "А. Адлер создаёт индивидуальную психологию"),
             ("1913", JUNG, "К. Г. Юнг: аналитическая психология"),
             ("1919–20", ADLER, "Венские медико-педагогические консультации")]
    for i, (year, col, desc) in enumerate(items):
        y = 200 + i * 92
        t = tb(s, 779, y, 400, 40)
        para(t, year, 30, col, font=SERIF, italic=True, line=1.0, first=True)
        t = tb(s, 779, y + 40, 420, 44)
        para(t, desc, 13, MUTED, line=1.45, first=True)


# ─────────────────── слайд 3 · Адлер: неполноценность ───────────────────
def s03(prs):
    s = add_slide(prs)
    head(s, "Глава 1 · Адлер", "Чувство неполноценности и компенсация",
         "Как недостаток становится источником развития", ADLER)

    rect(s, PADX, 200, 1120, 78, fill=ADLER_TINT)
    rect(s, PADX, 200, 4, 78, fill=ADLER)
    t = tb(s, 108, 214, 1076, 54)
    rich(t, [("Чувство неполноценности", True, False, ADLER_DARK),
             (" — нормальное переживание каждого ребёнка: он мал, слаб и "
              "зависим от взрослых. Это не болезнь, а стимул к развитию.",
              False, False, INK)], 16, INK, line=1.5, first=True)

    for x, accent, lab, title, bullets in [
        (PADX, ADLER, "Конструктивный путь", "Рост и сотрудничество",
         ["преодоление трудности, рост усилий",
          "поиск другой области успеха",
          "уверенность и вклад в общее дело"]),
        (662, JUNG, "Деструктивный путь", "Защита и уход",
         ["агрессия, демонстративность, обман",
          "лень как защита самооценки",
          "уход в болезнь, отказ от деятельности"]),
    ]:
        hline(s, x, 306, 538, color=accent, lw=1.5)
        t = tb(s, x, 318, 538, 20)
        para(t, lab, 11, accent, bold=True, caps=True, spacing=2.0, first=True)
        t = tb(s, x, 344, 538, 34)
        para(t, title, 24, INK, font=SERIF, line=1.15, first=True)
        t = tb(s, x, 388, 538, 150)
        for i, b in enumerate(bullets):
            para(t, "—  " + b, 15, INK_SOFT, line=1.5, before=7,
                 first=(i == 0))

    foot(s, [("Для учителя: сравнение и унижение толкают ребёнка к ", False),
             ("деструктивной компенсации", True), (".", False)])


# ─────────────────── слайд 4 · Адлер: социальный интерес ───────────────────
def s04(prs):
    s = add_slide(prs)
    head(s, "Глава 1 · Адлер", "Социальный интерес и школьная практика",
         "Gemeinschaftsgefühl — чувство общности", ADLER)

    t = tb(s, PADX, 200, 488, 220)
    para(t, "«Критерий психического здоровья — связь с людьми, а не "
            "превосходство»", 44, INK, font=SERIF, italic=True, line=1.1,
         first=True)
    t = tb(s, PADX, 396, 488, 26)
    para(t, "Альфред Адлер", 19, MUTED, first=True)
    t = tb(s, PADX, 442, 488, 180)
    rich(t, [("Социальный интерес — ", False, False, None),
             ("врождённая склонность быть частью сообщества", False, True, INK),
             (", сотрудничать и вносить вклад в общее дело. Его нужно "
              "воспитывать, и школа — вторая среда после семьи, где это "
              "происходит.", False, False, None)],
         16, INK_SOFT, line=1.65, first=True)

    vline(s, 624, 200, 400)
    t = tb(s, 656, 200, 544, 22)
    para(t, "Венская практика · 1919–1932", 11, ADLER, bold=True, caps=True,
         spacing=2.4, first=True)
    rows = [("1919–20", "Медико-педагогические консультации.",
             " Учитель, врач и психолог работают вместе с классом."),
            ("1920-е", "Экспериментальная школа.",
             " Без наказаний и соперничества ради отметки: общие дела и "
             "детское самоуправление."),
            ("Сегодня", "Школьная психологическая служба.",
             " «Трудное» поведение — сигнал о неудовлетворённой потребности, "
             "а не «плохой характер».")]
    y = 238
    for i, (year, bold_part, rest) in enumerate(rows):
        t = tb(s, 656, y, 100, 30)
        para(t, year, 21, ADLER, font=SERIF, italic=True, line=1.0, first=True)
        t = tb(s, 776, y - 2, 424, 90)
        rich(t, [(bold_part, True, False, INK), (rest, False, False, None)],
             15, INK_SOFT, line=1.55, first=True)
        y += 122
        if i < len(rows) - 1:
            hline(s, 656, y - 26, 544)

    foot(s, [("Задача школы — ", False), ("развивать сотрудничество", True),
             (", а не соперничество и страх оценки.", False)])


# ─────────────────── слайд 5 · Юнг: бессознательное ───────────────────
def s05(prs):
    s = add_slide(prs)
    head(s, "Глава 2 · Юнг", "Коллективное бессознательное",
         "Глубинные слои психики и универсальные первообразы", JUNG)

    layers = [("Сознание", "опыт «здесь и сейчас»", L1_BG, INK, MUTED),
              ("Личное бессознательное", "вытесненное и забытое", L2_BG, INK,
               RGBColor(0x5C, 0x56, 0x4D)),
              ("Коллективное бессознательное", "общий для всех слой", JUNG,
               WHITE, RGBColor(0xD8, 0xD2, 0xE4))]
    y = 200
    for i, (name, desc, fill, fg, dcol) in enumerate(layers):
        rect(s, PADX, y, 574, 54, fill=fill, line=LINE)
        t = tb(s, 102, y + 14, 400, 28)
        para(t, name, 21, fg, font=SERIF, line=1.1, first=True)
        t = tb(s, 502, y + 19, 130, 22)
        para(t, desc, 12, dcol, align=PP_ALIGN.RIGHT, first=True)
        y += 54
    t = tb(s, PADX, y + 14, 574, 44)
    para(t, "Наследуется как структура опыта, а не как готовое содержание.",
         13, MUTED, font=SERIF, italic=True, line=1.5, first=True)

    t = tb(s, 710, 200, 490, 22)
    para(t, "Архетипы — «пустые формы»", 11, JUNG, bold=True, caps=True,
         spacing=2.4, first=True)
    chips = [("Самость", True), ("Тень", False), ("Персона", False),
             ("Анима · Анимус", False), ("Мудрец", False),
             ("Великая Мать", False), ("Герой", False)]
    cx, cy = 710, 234
    for text, primary in chips:
        w = 30 + len(text) * 8.4
        if cx + w > 1200:
            cx, cy = 710, cy + 40
        rect(s, cx, cy, w, 32, fill=(JUNG if primary else None),
             line=JUNG if primary else LINE_STRONG)
        t = tb(s, cx + 13, cy + 7, w - 26, 20)
        para(t, text, 14, WHITE if primary else INK_SOFT,
             bold=primary, first=True)
        cx += w + 8

    t = tb(s, 710, cy + 58, 490, 190)
    rich(t, [("Сказка, миф, игра, любимый герой — ", False, False, None),
             ("встреча ребёнка", False, True, INK),
             (" с общечеловеческими образами. Через них он переживает добро и "
              "зло, страх и преодоление в безопасной символической форме.",
              False, False, None)], 15, INK_SOFT, line=1.6, first=True)

    foot(s, [("Отсюда — ", False),
             ("ценность творчества и работы с символами", True),
             (" в детском возрасте.", False)])


# ─────────────────── слайд 6 · Юнг: типы ───────────────────
def s06(prs):
    s = add_slide(prs)
    head(s, "Глава 2 · Юнг", "Психологические типы и обучение",
         "Две установки и четыре функции", JUNG)

    for x, name, txt in [
        (PADX, "Экстраверсия",
         "Энергия направлена вовне. Учится через общение, движение, "
         "обсуждение и совместную работу."),
        (628, "Интроверсия",
         "Энергия направлена внутрь. Учится через размышление, паузу и "
         "спокойную сосредоточенность."),
    ]:
        hline(s, x, 200, 532, color=JUNG, lw=1.5)
        t = tb(s, x, 214, 532, 20)
        para(t, "Установка", 11, MUTED, bold=True, caps=True, spacing=1.5,
             first=True)
        t = tb(s, x, 238, 532, 40)
        para(t, name, 30, INK, font=SERIF, line=1.0, first=True)
        t = tb(s, x, 286, 532, 70)
        bold_part = "вовне" if x == PADX else "внутрь"
        pre, post = txt.split(bold_part, 1)
        rich(t, [(pre, False, False, None), (bold_part, True, False, INK),
                 (post, False, False, None)], 15, INK_SOFT, line=1.55,
             first=True)

    hline(s, PADX, 430, 1120)
    t = tb(s, PADX, 444, 1120, 22)
    para(t, "Четыре функции · как ребёнок воспринимает и решает", 11, JUNG,
         bold=True, caps=True, spacing=2.4, first=True)
    funcs = [("Мышление", "логика, структура, правила"),
             ("Чувство", "принятие, личное отношение"),
             ("Ощущение", "наглядность, практика"),
             ("Интуиция", "смысл, догадка, перспектива")]
    colw = 1120 / 4
    for i, (name, desc) in enumerate(funcs):
        x = PADX + i * colw
        if i:
            vline(s, x, 480, 90)
        tx = x if i == 0 else x + 20
        t = tb(s, tx, 484, colw - 20, 30)
        para(t, name, 19, INK, font=SERIF, line=1.1, first=True)
        t = tb(s, tx, 516, colw - 20, 50)
        para(t, desc, 13, MUTED, line=1.4, first=True)

    foot(s, [("Единый метод обучения не годится для всех — ", False),
             ("отсюда идея индивидуальных стилей обучения", True), (".", False)])


# ─────────────────── слайд 7 · Юнг: учитель и индивидуация ───────────────────
def s07(prs):
    s = add_slide(prs)
    head(s, "Глава 2 · Юнг", "Учитель как архетип и индивидуация",
         "Личность педагога и цель развития ребёнка", JUNG)

    t = tb(s, PADX, 200, 497, 22)
    para(t, "Учитель в детском восприятии", 11, JUNG, bold=True, caps=True,
         spacing=2.4, first=True)
    for i, chip in enumerate(["Мудрец", "Мастер"]):
        x = PADX + i * 96
        rect(s, x, 230, 88, 28, fill=None, line=JUNG)
        t = tb(s, x + 12, 236, 64, 18)
        para(t, chip, 13, JUNG, bold=True, first=True)
    t = tb(s, PADX, 276, 497, 200)
    rich(t, [("На учителя проецируются ", False, False, None),
             ("ожидания, надежды и страхи", False, True, INK),
             (" ребёнка. Педагог влияет прежде всего личностью, а не словами.",
              False, False, None)], 15, INK_SOFT, line=1.6, after=10,
         first=True)
    para(t, "Опасность — отождествить себя с идеальным, безошибочным "
            "Мудрецом: тогда неудача ребёнка воспринимается как личное "
            "оскорбление.", 15, INK_SOFT, line=1.6, after=10)
    rich(t, [("Самопознание учителя", False, True, INK),
             (" — часть его профессиональной работы.", False, False, None)],
         15, INK_SOFT, line=1.6)

    vline(s, 633, 200, 400)
    t = tb(s, 669, 200, 531, 22)
    para(t, "Цель развития", 11, JUNG, bold=True, caps=True, spacing=2.4,
         first=True)
    t = tb(s, 669, 228, 531, 110)
    para(t, "стать", 48, INK, font=SERIF, italic=True, line=1.05, first=True)
    para(t, "самим собой", 48, INK, font=SERIF, italic=True, line=1.05)
    t = tb(s, 669, 356, 531, 150)
    for i, point in enumerate([
        "Внутренняя целостность и согласие с собой",
        "Признание Тени — своих непризнанных качеств",
        "Раскрытие собственной неповторимости",
    ]):
        para(t, "—   " + point, 15, INK_SOFT, line=1.5, before=12,
             first=(i == 0))

    foot(s, [("Учитель влияет не тем, что говорит, а ", False),
             ("тем, кем он является", True), (".", False)])


# ─────────────────── слайд 8 · сравнение ───────────────────
def s08(prs):
    s = add_slide(prs)
    head(s, "Глава 3 · Сравнение", "Общее и различное", None, None)

    hline(s, PADX, 190, 1120)
    t = tb(s, PADX, 204, 120, 30)
    para(t, "Общее", 19, MUTED, font=SERIF, italic=True, first=True)
    t = tb(s, 214, 208, 986, 40)
    para(t, "критика авторитарной школы · интерес к бессознательному · "
            "связь психологии и педагогики · уважение к личности ребёнка",
         14, INK_SOFT, line=1.5, first=True)
    hline(s, PADX, 256, 1120)

    c1, c2, c3 = PADX, 358, 794
    w1, w2, w3 = 246, 404, 404
    t = tb(s, c1, 276, w1, 20)
    para(t, "Критерий", 11, MUTED, bold=True, caps=True, spacing=2.4, first=True)
    t = tb(s, c2, 276, w2, 20)
    para(t, "Адлер", 11, ADLER, bold=True, caps=True, spacing=2.4, first=True)
    t = tb(s, c3, 276, w3, 20)
    para(t, "Юнг", 11, JUNG, bold=True, caps=True, spacing=2.4, first=True)
    hline(s, PADX, 300, 1120, color=LINE_STRONG)

    rows = [("Движущая сила", "неполноценность и компенсация",
             "архетипы и индивидуация"),
            ("Ключевое понятие", "социальный интерес",
             "коллективное бессознательное"),
            ("Критерий здоровья", "сотрудничество, вклад",
             "целостность, согласие с собой"),
            ("Роль учителя", "партнёр и помощник", "носитель архетипа, пример"),
            ("Цель школы", "стать успешным вместе с другими",
             "стать самим собой")]
    y = 314
    for i, (a, b, c) in enumerate(rows):
        t = tb(s, c1, y, w1, 26)
        para(t, a, 14, INK, bold=True, first=True)
        t = tb(s, c2, y, w2, 26)
        para(t, b, 15, INK_SOFT, line=1.45, first=True)
        t = tb(s, c3, y, w3, 26)
        para(t, c, 15, INK_SOFT, line=1.45, first=True)
        y += 52
        if i < len(rows) - 1:
            hline(s, PADX, y - 14, 1120)

    foot(s, [("Формула: Адлер — «стать успешным вместе с другими» · Юнг — "
              "«стать самим собой».", False)])


# ─────────────────── слайд 9 · современная школа ───────────────────
def s09(prs):
    s = add_slide(prs)
    head(s, "Практика", "Значение для современной школы",
         "Что школа берёт у Адлера и у Юнга", None)

    for x, accent, sub, title, items in [
        (PADX, ADLER, "От Адлера", "Отношения и мотивация",
         ["«трудное» поведение — сигнал, а не характер",
          "поддержка и поощрение вместо наказания",
          "развитие сотрудничества и самоуправления",
          "внимание к самооценке и мотивации"]),
        (628, JUNG, "От Юнга", "Внутренний мир и тип",
         ["учёт индивидуальных стилей обучения",
          "ценность творчества, сказки, игры",
          "работа с символами и внутренним миром",
          "внимание к личности учителя"]),
    ]:
        hline(s, x, 200, 532, color=accent, lw=1.5)
        t = tb(s, x, 214, 532, 20)
        para(t, sub, 11, MUTED, bold=True, caps=True, spacing=2.2, first=True)
        t = tb(s, x, 238, 532, 36)
        para(t, title, 26, accent, font=SERIF, line=1.1, first=True)
        y = 288
        for i, it in enumerate(items):
            t = tb(s, x, y, 532, 44)
            para(t, "→   " + it, 15, INK_SOFT, line=1.55, first=True)
            y += 52
            if i < len(items) - 1:
                hline(s, x, y - 18, 532)

    foot(s, [("Двойная опора учителя: ", False),
             ("понимание отношений", True), (" и ", False),
             ("понимание внутреннего мира ребёнка", True), (".", False)])


# ─────────────────── слайд 10 · выводы ───────────────────
def s10(prs):
    s = add_slide(prs)
    head(s, "Итоги", "Выводы и источники", None, None)

    t = tb(s, PADX, 200, 567, 190)
    rich(t, [("Адлер:", True, False, ADLER),
             (" школа через отношения — неполноценность, компенсация, "
              "социальный интерес.", False, False, None)],
         26, INK, font=SERIF, line=1.3, after=18, first=True)
    rich(t, [("Юнг:", True, False, JUNG),
             (" школа через внутренний мир — архетипы, типы, индивидуация.",
              False, False, None)], 26, INK, font=SERIF, line=1.3)
    t = tb(s, PADX, 424, 567, 160)
    rich(t, [("Подходы ", False, False, None),
             ("не противоречат, а дополняют", True, False, INK),
             (" друг друга: «стать успешным вместе с другими» и «стать самим "
              "собой» — две стороны одной задачи — помочь ребёнку вырасти "
              "целостным человеком.", False, False, None)],
         15, INK_SOFT, line=1.6, first=True)

    vline(s, 703, 200, 400)
    t = tb(s, 735, 200, 465, 22)
    para(t, "Источники", 11, MUTED, bold=True, caps=True, spacing=2.4,
         first=True)
    srcs = [
        "Леонтьев Д. А. Адлер Альфред // БРЭ. URL: "
        "old.bigenc.ru/psychology/text/v/1800877",
        "Леонтьев Д. А. Индивидуальная психология // БРЭ. URL: "
        "old.bigenc.ru/psychology/text/v/2008679",
        "Мавлявиева Л. И. Личностно-социальная концепция воспитания "
        "А. Адлера : дис. … канд. пед. наук. URL: dslib.net/obw-pedagogika/"
        "lichnostno-socialnaja-koncepcija-vospitanija-alfreda-adlera.html",
        "Юнг (Jung) Карл Густав // Психологический словарь. URL: "
        "rus-big-psyho.slovaronline.com/3359-Юнг",
        "Хьелл Л., Зиглер Д. Теории личности (фрагмент). URL: "
        "gbs.spb.ru/ru/search/detail/?id=61e0492f12935ed6424cf239c090baf3",
        "Модель К. Юнга и современные стили обучения // КиберЛенинка. URL: "
        "cyberleninka.ru/article/n/model-k-yunga-i-sovremennye-stili-obucheniya",
        "Архетипическая педагогика (обзор) // Alphapedia. URL: "
        "alphapedia.ru/w/Archetypal_pedagogy",
    ]
    t = tb(s, 735, 228, 465, 360)
    for i, src in enumerate(srcs):
        para(t, f"{i + 1}.  {src}", 12.5, INK_SOFT, line=1.45, after=5,
             first=(i == 0))

    hline(s, PADX, 624, 1120)
    t = tb(s, PADX, 632, 700, 40)
    para(t, "Спасибо за внимание", 30, INK, font=SERIF, italic=True, first=True)
    t = tb(s, 700, 642, 500, 26)
    para(t, "Сабадаш П. А. · ИС-24 · 2026", 13, MUTED,
         align=PP_ALIGN.RIGHT, first=True)


# ─────────────────────────── общие блоки ───────────────────────────
def head(s, label, title, sub, accent):
    t = tb(s, PADX, PADY, 900, 22)
    para(t, label, 11, accent or MUTED, bold=True, caps=True, spacing=2.4,
         first=True)
    t = tb(s, PADX, 92, 1120, 56)
    para(t, title, 42, INK, font=SERIF, line=1.12, first=True)
    if sub:
        t = tb(s, PADX, 152, 1120, 30)
        para(t, sub, 19, MUTED, font=SERIF, italic=True, first=True)


def foot(s, parts):
    hline(s, PADX, 630, 1120)
    t = tb(s, PADX, 640, 1120, 46)
    rich(t, [(p, b, False, INK if b else INK_SOFT) for p, b in parts],
         15, INK_SOFT, line=1.45, first=True)


def main():
    prs = new_deck()
    for fn in (s01, s02, s03, s04, s05, s06, s07, s08, s09, s10):
        fn(prs)
    out = "presentation/psihologia_shkoly_adler_jung.pptx"
    prs.save(out)
    print("сохранено:", out, "| слайдов:", len(prs.slides.__iter__.__self__._sldIdLst))


if __name__ == "__main__":
    main()
