#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Сборка презентации PowerPoint (16:9) по теме
«Психология школы в подходах Альфреда Адлера и Карла Густава Юнга».

Дизайн повторяет HTML-версию presentation/index.html:
тёплый фон, Georgia в заголовках, охра — Адлер, сливовый — Юнг,
тонкие линии вместо карточек.

Запуск:
    pip install python-pptx
    python3 tools/build_pptx.py
Результат:
    presentation/psihobsh-adler-jung.pptx
"""

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls

# --------------------------------------------------------------------------
# Геометрия: 1280 x 720 px при 96 dpi, как в HTML-версии
# --------------------------------------------------------------------------
EMU_PER_PX = 9525
W, H = 1280, 720
MARGIN_X, MARGIN_TOP, MARGIN_BOTTOM = 80, 64, 64
CONTENT_W = W - 2 * MARGIN_X

# --------------------------------------------------------------------------
# Палитра
# --------------------------------------------------------------------------
BG = "FDFCF8"
INK = "1C1917"
INK_SOFT = "3F3A36"
MUTED = "6E6763"
LINE = "E5E1D8"
LINE_STRONG = "D6D0C2"
PANEL = "F5F1E8"
ADLER = "B45309"
ADLER_DARK = "9A4A06"
ADLER_TINT = "FDF0DF"
JUNG = "4C3575"
JUNG_TINT = "EDE7F5"

SERIF = "Georgia"
SANS = "Segoe UI"

OUT = Path(__file__).resolve().parent.parent / "presentation" / "psihobsh-adler-jung.pptx"


def px(value):
    return Emu(int(round(value * EMU_PER_PX)))


def rgb(hex_str):
    return RGBColor.from_string(hex_str)


# --------------------------------------------------------------------------
# Базовые помощники
# --------------------------------------------------------------------------
def no_shadow(shape):
    try:
        shape.shadow.inherit = False
    except Exception:
        pass
    return shape


def rect(slide, x, y, w, h, fill=None, line=None, line_w=1.0, shape=MSO_SHAPE.RECTANGLE):
    s = slide.shapes.add_shape(shape, px(x), px(y), px(w), px(h))
    no_shadow(s)
    if fill:
        s.fill.solid()
        s.fill.fore_color.rgb = rgb(fill)
    else:
        s.fill.background()
    if line:
        s.line.color.rgb = rgb(line)
        s.line.width = Pt(line_w)
    else:
        s.line.fill.background()
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        s.adjustments[0] = 0.12
    s.text_frame.text = ""
    return s


def rule(slide, x, y, w, color=LINE, thickness=1):
    return rect(slide, x, y, w, thickness, fill=color)


def _set_bullet(p, char, color, indent_px=18):
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(int(indent_px * EMU_PER_PX)))
    pPr.set("indent", str(-int(indent_px * EMU_PER_PX)))
    for tag in ("a:buNone", "a:buChar", "a:buAutoNum", "a:buClr", "a:buFont", "a:buSzPct"):
        for el in pPr.findall(qn(tag)):
            pPr.remove(el)
    bu_clr = parse_xml('<a:buClr %s><a:srgbClr val="%s"/></a:buClr>' % (nsdecls("a"), color))
    pPr.append(bu_clr)
    bu_font = parse_xml('<a:buFont %s typeface="Arial"/>' % nsdecls("a"))
    pPr.append(bu_font)
    bu_char = parse_xml('<a:buChar %s char="%s"/>' % (nsdecls("a"), char))
    pPr.append(bu_char)


def _set_no_bullet(p, indent_px, hanging=True):
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(int(indent_px * EMU_PER_PX)))
    pPr.set("indent", str(-int(indent_px * EMU_PER_PX) if hanging else 0))
    for tag in ("a:buChar", "a:buAutoNum", "a:buClr", "a:buFont", "a:buSzPct"):
        for el in pPr.findall(qn(tag)):
            pPr.remove(el)
    if pPr.find(qn("a:buNone")) is None:
        pPr.append(parse_xml("<a:buNone %s/>" % nsdecls("a")))


def _fill_run(run, spec):
    f = run.font
    f.size = Pt(spec.get("size", 16) * 0.75)
    f.bold = spec.get("bold", False)
    f.italic = spec.get("italic", False)
    f.name = spec.get("font", SANS)
    f.color.rgb = rgb(spec.get("color", INK))
    rPr = run._r.get_or_add_rPr()
    if spec.get("caps"):
        rPr.set("cap", "all")
    if spec.get("spc"):
        # межбуквенный интервал в px -> сотые доли пункта
        rPr.set("spc", str(int(round(spec["spc"] * 0.75 * 100))))


def add_text(slide, x, y, w, h, paras, anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(px(x), px(y), px(w), px(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = 0
    tf.margin_top = tf.margin_bottom = 0
    for i, spec in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = spec.get("align", PP_ALIGN.LEFT)
        if "line" in spec:
            p.line_spacing = spec["line"]
        if "before" in spec:
            p.space_before = Pt(spec["before"] * 0.75)
        if "after" in spec:
            p.space_after = Pt(spec["after"] * 0.75)
        if spec.get("bullet"):
            _set_bullet(p, spec["bullet"], spec.get("bullet_color", MUTED),
                        spec.get("indent", 18))
        elif spec.get("hanging"):
            _set_no_bullet(p, spec.get("indent", 18))

        base = {k: spec[k] for k in ("size", "font", "color", "bold", "italic") if k in spec}
        runs = spec.get("runs")
        if runs is None:
            runs = [{"text": spec.get("text", "")}]
        for r in runs:
            merged = dict(base)
            merged.update(r)
            run = p.add_run()
            run.text = r.get("text", "")
            _fill_run(run, merged)
    return box


def header(slide, label, title, sub=None, accent=MUTED):
    add_text(slide, MARGIN_X, 60, CONTENT_W, 18, [{
        "text": label, "size": 11, "bold": True, "color": accent,
        "caps": True, "spc": 2.4,
    }])
    add_text(slide, MARGIN_X, 86, CONTENT_W, 54, [{
        "text": title, "size": 42, "font": SERIF, "color": INK, "line": 1.05,
    }])
    if sub:
        add_text(slide, MARGIN_X, 143, CONTENT_W, 28, [{
            "text": sub, "size": 19, "font": SERIF, "italic": True, "color": MUTED,
        }])


def footer(slide, runs):
    rule(slide, MARGIN_X, 606, CONTENT_W, LINE, 1)
    add_text(slide, MARGIN_X, 622, CONTENT_W, 30, [{
        "runs": runs, "size": 15, "color": INK_SOFT, "line": 1.2,
    }])


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def new_slide(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = rgb(BG)
    return s


# --------------------------------------------------------------------------
# 1. Титульный слайд
# --------------------------------------------------------------------------
def slide_cover(prs):
    s = new_slide(prs)

    add_text(s, MARGIN_X, 62, 760, 20, [{
        "text": "Психология школы · Тема №8", "size": 12, "bold": True,
        "color": MUTED, "caps": True, "spc": 3,
    }])
    add_text(s, 1000, 40, 200, 70, [{
        "text": "08", "size": 52, "font": SERIF, "italic": True,
        "color": MUTED, "align": PP_ALIGN.RIGHT,
    }])

    add_text(s, MARGIN_X, 204, 630, 190, [
        {"text": "Психология", "size": 94, "font": SERIF, "color": INK, "line": 0.95},
        {"text": "школы", "size": 94, "font": SERIF, "italic": True,
         "color": MUTED, "line": 0.95},
    ])
    add_text(s, MARGIN_X, 416, 600, 80, [{
        "text": "в подходах Альфреда Адлера\nи Карла Густава Юнга",
        "size": 22, "font": SERIF, "italic": True, "color": INK_SOFT, "line": 1.4,
    }])

    rect(s, 748, 200, 1, 240, fill=LINE)
    for top, letter, lcolor, who, when in (
        (208, "А", ADLER, "Альфред Адлер", "индивидуальная психология"),
        (332, "Ю", JUNG, "Карл Густав Юнг", "аналитическая психология"),
    ):
        add_text(s, 772, top, 70, 84, [{
            "text": letter, "size": 70, "font": SERIF, "italic": True,
            "color": lcolor, "line": 0.85,
        }])
        add_text(s, 872, top + 20, 320, 64, [
            {"text": who, "size": 17, "bold": True, "color": INK},
            {"text": when, "size": 13, "color": MUTED, "before": 4},
        ])

    rule(s, MARGIN_X, 610, CONTENT_W, LINE, 1)
    add_text(s, MARGIN_X, 626, 700, 26, [{
        "text": "Сабадаш Павел Андреевич · группа ИС-24",
        "size": 14, "bold": True, "color": INK, "spc": 0.3,
    }])
    add_text(s, 800, 626, 400, 26, [{
        "text": "РКСИ · Ростов-на-Дону · 2026", "size": 14, "color": MUTED,
        "align": PP_ALIGN.RIGHT,
    }])

    notes(s, "Добрый день. Тема моего доклада — психология школы в подходах "
             "Альфреда Адлера и Карла Густава Юнга. Это два ученика Фрейда, "
             "которые разошлись с ним и построили собственные теории, но оба "
             "много занимались детством и школой. Я сравню их взгляды на ребёнка, "
             "учителя и обучение и покажу, что современная школа берёт у обоих.")


# --------------------------------------------------------------------------
# 2. Контекст
# --------------------------------------------------------------------------
def slide_context(prs):
    s = new_slide(prs)
    header(s, "Контекст", "Почему школа стала объектом психологии",
           "Начало XX века: массовая школа, детство и бессознательное")

    add_text(s, MARGIN_X, 200, 620, 390, [
        {"runs": [
            {"text": "К началу XX века в Европе складывается "},
            {"text": "массовая обязательная школа", "font": SERIF, "italic": True, "color": INK},
            {"text": ". Впервые огромное число детей одновременно оказывается в одном "
                     "пространстве, подчиняется общим правилам и сравнивается по единой мерке."},
        ], "size": 16, "color": INK_SOFT, "line": 1.62, "after": 13},
        {"runs": [
            {"text": "Школа становится "},
            {"text": "первым институтом после семьи", "font": SERIF, "italic": True, "color": INK},
            {"text": ", где ребёнок систематически переживает успех и неудачу, "
                     "сталкивается с авторитетом и оценкой. Именно здесь формируются его "
                     "представления о себе и о своём месте среди других."},
        ], "size": 16, "color": INK_SOFT, "line": 1.62, "after": 13},
        {"text": "Психоанализ выходит за пределы клиники. Интерес к детству и "
                 "бессознательным мотивам переносится на педагогику. Адлер и Юнг — "
                 "оба из круга Фрейда — предлагают свои ответы на вопрос: что происходит "
                 "с личностью ребёнка в школе.",
         "size": 16, "color": INK_SOFT, "line": 1.62},
    ])

    rect(s, 742, 200, 1, 350, fill=LINE)
    timeline = (
        ("1900-е", ADLER, "Массовая школа в Европе"),
        ("1911", ADLER, "А. Адлер создаёт индивидуальную психологию"),
        ("1913", JUNG, "К. Г. Юнг: аналитическая психология"),
        ("1919–20", ADLER, "Венские медико-педагогические консультации"),
    )
    y = 198
    for year, color, desc in timeline:
        add_text(s, 770, y, 430, 42, [{
            "text": year, "size": 30, "font": SERIF, "italic": True,
            "color": color, "line": 1.0,
        }])
        add_text(s, 770, y + 38, 420, 44, [{
            "text": desc, "size": 13, "color": MUTED, "line": 1.4,
        }])
        y += 88

    notes(s, "Сначала — почему школа вообще стала предметом психологии. "
             "В начале XX века появляется массовая обязательная школа: впервые "
             "большое число детей одновременно оценивают по общей мерке. Школа — "
             "первый после семьи институт, где ребёнок переживает успех и неудачу "
             "публично. Параллельно психоанализ выходит из клиники в педагогику. "
             "Адлер создаёт индивидуальную психологию в 1911 году, Юнг — "
             "аналитическую в 1913-м, а с 1919 года Адлер открывает венские "
             "медико-педагогические консультации.")


# --------------------------------------------------------------------------
# 3. Адлер: неполноценность и компенсация
# --------------------------------------------------------------------------
def slide_inferiority(prs):
    s = new_slide(prs)
    header(s, "Глава 1 · Адлер", "Чувство неполноценности и компенсация",
           "Как недостаток становится источником развития", accent=ADLER)

    rect(s, MARGIN_X, 194, CONTENT_W, 76, fill=ADLER_TINT)
    rect(s, MARGIN_X, 194, 3, 76, fill=ADLER)
    add_text(s, 104, 206, 1064, 56, [{
        "runs": [
            {"text": "Чувство неполноценности", "bold": True, "color": ADLER_DARK},
            {"text": " — нормальное переживание каждого ребёнка: он мал, слаб и зависим "
                     "от взрослых. Это не болезнь, а стимул к развитию."},
        ],
        "size": 16, "color": INK, "line": 1.45,
    }])

    columns = (
        (MARGIN_X, ADLER, "Конструктивный путь", "Рост и сотрудничество", (
            "преодоление трудности, рост усилий",
            "поиск другой области успеха",
            "уверенность и вклад в общее дело",
        )),
        (660, JUNG, "Деструктивный путь", "Защита и уход", (
            "агрессия, демонстративность, обман",
            "лень как защита самооценки",
            "уход в болезнь, отказ от деятельности",
        )),
    )
    for x, color, label, title, items in columns:
        rule(s, x, 344, 500, color, 2)
        add_text(s, x, 358, 500, 18, [{
            "text": label, "size": 11, "bold": True, "color": color,
            "caps": True, "spc": 2,
        }])
        add_text(s, x, 380, 500, 36, [{
            "text": title, "size": 24, "font": SERIF, "color": INK,
        }])
        add_text(s, x, 422, 500, 140, [{
            "text": item, "size": 16, "color": INK_SOFT, "line": 1.55,
            "bullet": "—", "bullet_color": color, "indent": 18, "after": 8,
        } for item in items])

    footer(s, [
        {"text": "Для учителя: сравнение и унижение толкают ребёнка к "},
        {"text": "деструктивной компенсации", "bold": True, "color": INK},
        {"text": "."},
    ])

    notes(s, "У Адлера ключевое понятие — чувство неполноценности. Оно есть у "
             "каждого ребёнка, потому что дети действительно малы, слабы и зависимы. "
             "Это не болезнь, а стимул. Дальше возможны два пути компенсации. "
             "Конструктивный: ребёнок преодолевает трудность, находит другую область "
             "успеха, становится уверенным и полезным другим. Деструктивный: агрессия, "
             "демонстративность, лень как защита самооценки, уход в болезнь. "
             "Вывод для учителя: сравнение и унижение толкают к деструктивному пути.")


# --------------------------------------------------------------------------
# 4. Адлер: социальный интерес
# --------------------------------------------------------------------------
def slide_social(prs):
    s = new_slide(prs)
    header(s, "Глава 1 · Адлер", "Социальный интерес и школьная практика",
           "Gemeinschaftsgefühl — чувство общности", accent=ADLER)

    add_text(s, MARGIN_X, 196, 470, 130, [{
        "text": "«Критерий психического здоровья — связь с людьми, а не превосходство»",
        "size": 24, "font": SERIF, "italic": True, "color": INK, "line": 1.35,
    }])
    add_text(s, MARGIN_X, 322, 470, 20, [{
        "text": "Альфред Адлер", "size": 11, "bold": True, "color": ADLER,
        "caps": True, "spc": 2,
    }])
    rule(s, MARGIN_X, 352, 90, ADLER, 2)

    add_text(s, MARGIN_X, 372, 470, 150, [{
        "runs": [
            {"text": "Социальный интерес — "},
            {"text": "врождённая склонность быть частью сообщества", "font": SERIF,
             "italic": True, "color": INK},
            {"text": ", сотрудничать и вносить вклад в общее дело. Его нужно "
                     "воспитывать, и школа — вторая среда после семьи, где это "
                     "происходит."},
        ],
        "size": 15.5, "color": INK_SOFT, "line": 1.62,
    }])

    rect(s, 600, 196, 1, 340, fill=LINE)
    add_text(s, 630, 196, 560, 18, [{
        "text": "Венская практика · 1919–1932", "size": 11, "bold": True,
        "color": MUTED, "caps": True, "spc": 2,
    }])
    practice = (
        ("1919–20", "Медико-педагогические консультации.",
         " Учитель, врач и психолог работают вместе с классом."),
        ("1920-е", "Экспериментальная школа.",
         " Без наказаний и соперничества ради отметки: общие дела и детское самоуправление."),
        ("Сегодня", "Школьная психологическая служба.",
         " «Трудное» поведение — сигнал о неудовлетворённой потребности, а не «плохой характер»."),
    )
    y = 230
    for year, lead, tail in practice:
        add_text(s, 630, y, 560, 20, [{
            "text": year, "size": 12, "bold": True, "color": ADLER, "spc": 0.5,
        }])
        add_text(s, 630, y + 22, 550, 90, [{
            "runs": [{"text": lead, "bold": True, "color": INK}, {"text": tail}],
            "size": 14.5, "color": INK_SOFT, "line": 1.5,
        }])
        y += 112

    footer(s, [
        {"text": "Задача школы — "},
        {"text": "развивать сотрудничество", "bold": True, "color": INK},
        {"text": ", а не соперничество и страх оценки."},
    ])

    notes(s, "Второе понятие Адлера — социальный интерес, по-немецки "
             "Gemeinschaftsgefühl. Это врождённая склонность быть частью сообщества, "
             "сотрудничать и вносить вклад. Критерий психического здоровья у Адлера — "
             "не превосходство, а связь с людьми. Социальный интерес нужно "
             "воспитывать, и школа здесь вторая среда после семьи. В Вене Адлер "
             "открыл медико-педагогические консультации, где учитель, врач и психолог "
             "работали вместе. В экспериментальной школе не было наказаний и "
             "соперничества ради отметки. Сегодня это логика школьной психологической "
             "службы: трудное поведение — сигнал потребности, а не плохой характер.")


# --------------------------------------------------------------------------
# 5. Юнг: коллективное бессознательное
# --------------------------------------------------------------------------
def slide_archetypes(prs):
    s = new_slide(prs)
    header(s, "Глава 2 · Юнг", "Коллективное бессознательное",
           "Глубинные слои психики и универсальные первообразы", accent=JUNG)

    layers = (
        (196, "Сознание", "опыт «здесь и сейчас»", "FFFFFF", LINE),
        (292, "Личное бессознательное", "вытесненное и забытое", PANEL, LINE),
        (388, "Коллективное бессознательное", "общий для всех слой", JUNG_TINT, None),
    )
    for y, name, desc, fill, border in layers:
        rect(s, MARGIN_X, y, 520, 88, fill=fill, line=border, line_w=1)
        rect(s, MARGIN_X, y, 3, 88, fill=JUNG)
        add_text(s, 104, y + 18, 480, 26, [{
            "text": name, "size": 17, "bold": True, "color": INK,
        }])
        add_text(s, 104, y + 50, 480, 22, [{
            "text": desc, "size": 12.5, "color": MUTED,
        }])

    add_text(s, MARGIN_X, 508, 520, 44, [{
        "text": "Наследуется как структура опыта, а не как готовое содержание.",
        "size": 13, "font": SERIF, "italic": True, "color": MUTED, "line": 1.45,
    }])

    add_text(s, 640, 196, 560, 18, [{
        "text": "Архетипы — «пустые формы»", "size": 11, "bold": True,
        "color": MUTED, "caps": True, "spc": 2,
    }])

    chips = (
        (640, 224, 128, "Самость", True),
        (776, 224, 96, "Тень", False),
        (880, 224, 124, "Персона", False),
        (640, 262, 186, "Анима · Анимус", False),
        (834, 262, 122, "Мудрец", False),
        (964, 262, 158, "Великая Мать", False),
        (640, 300, 108, "Герой", False),
    )
    for x, y, w, label, primary in chips:
        if primary:
            rect(s, x, y, w, 30, fill=JUNG_TINT, line=JUNG, line_w=1,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
            color = JUNG
        else:
            rect(s, x, y, w, 30, fill="FFFFFF", line=LINE_STRONG, line_w=1,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
            color = INK_SOFT
        add_text(s, x + 4, y + 7, w - 8, 18, [{
            "text": label, "size": 12.5, "color": color, "bold": primary,
            "align": PP_ALIGN.CENTER,
        }])

    add_text(s, 640, 354, 560, 130, [{
        "runs": [
            {"text": "Сказка, миф, игра, любимый герой — "},
            {"text": "встреча ребёнка", "font": SERIF, "italic": True, "color": INK},
            {"text": " с общечеловеческими образами. Через них он переживает добро "
                     "и зло, страх и преодоление в безопасной символической форме."},
        ],
        "size": 15, "color": INK_SOFT, "line": 1.6,
    }])

    footer(s, [
        {"text": "Отсюда — "},
        {"text": "ценность творчества и работы с символами", "bold": True, "color": INK},
        {"text": " в детском возрасте."},
    ])

    notes(s, "Теперь Юнг. Он делит психику на три слоя: сознание, личное "
             "бессознательное и коллективное бессознательное. Коллективное — общий "
             "для всех людей слой, который наследуется как структура опыта, а не как "
             "готовое содержание. Его единицы — архетипы, «пустые формы»: Самость, "
             "Тень, Персона, Анима и Анимус, Мудрец, Великая Мать, Герой. Ребёнок "
             "встречается с ними в сказке, мифе, игре, через любимого героя и "
             "безопасно переживает добро, зло, страх и преодоление. Поэтому так важны "
             "творчество и работа с символами.")


# --------------------------------------------------------------------------
# 6. Юнг: психологические типы
# --------------------------------------------------------------------------
def slide_types(prs):
    s = new_slide(prs)
    header(s, "Глава 2 · Юнг", "Психологические типы и обучение",
           "Две установки и четыре функции", accent=JUNG)

    settings = (
        (MARGIN_X, "Экстраверсия", [
            {"text": "Энергия направлена "},
            {"text": "вовне", "bold": True, "color": INK},
            {"text": ". Учится через общение, движение, обсуждение и совместную работу."},
        ]),
        (640, "Интроверсия", [
            {"text": "Энергия направлена "},
            {"text": "внутрь", "bold": True, "color": INK},
            {"text": ". Учится через размышление, паузу и спокойную сосредоточенность."},
        ]),
    )
    for x, name, runs in settings:
        rule(s, x, 200, 520, JUNG, 2)
        add_text(s, x, 213, 520, 16, [{
            "text": "Установка", "size": 10.5, "bold": True, "color": MUTED,
            "caps": True, "spc": 2,
        }])
        add_text(s, x, 234, 520, 34, [{
            "text": name, "size": 25, "font": SERIF, "color": INK,
        }])
        add_text(s, x, 276, 500, 80, [{
            "runs": runs, "size": 14.5, "color": INK_SOFT, "line": 1.5,
        }])

    add_text(s, MARGIN_X, 412, 800, 18, [{
        "text": "Четыре функции · как ребёнок воспринимает и решает",
        "size": 11, "bold": True, "color": MUTED, "caps": True, "spc": 2,
    }])

    funcs = (
        ("Мышление", "логика, структура, правила"),
        ("Чувство", "принятие, личное отношение"),
        ("Ощущение", "наглядность, практика"),
        ("Интуиция", "смысл, догадка, перспектива"),
    )
    for i, (name, desc) in enumerate(funcs):
        x = MARGIN_X + i * 285
        rule(s, x, 444, 255, JUNG, 2)
        add_text(s, x, 458, 255, 28, [{
            "text": name, "size": 19, "font": SERIF, "color": INK,
        }])
        add_text(s, x, 494, 255, 48, [{
            "text": desc, "size": 13, "color": MUTED, "line": 1.4,
        }])

    footer(s, [
        {"text": "Единый метод обучения не годится для всех — "},
        {"text": "отсюда идея индивидуальных стилей обучения", "bold": True, "color": INK},
        {"text": "."},
    ])

    notes(s, "Юнг описал две установки и четыре функции. Экстраверт направляет "
             "энергию вовне и учится через общение, движение, обсуждение; интроверт — "
             "внутрь, через размышление, паузу и сосредоточенность. Четыре функции — "
             "мышление, чувство, ощущение и интуиция — отвечают на вопрос, как ребёнок "
             "воспринимает материал и принимает решения. Отсюда важный практический "
             "вывод: единого метода для всех нет, нужны разные стили обучения и разные "
             "формы предъявления материала.")


# --------------------------------------------------------------------------
# 7. Юнг: учитель и индивидуация
# --------------------------------------------------------------------------
def slide_individuation(prs):
    s = new_slide(prs)
    header(s, "Глава 2 · Юнг", "Учитель как архетип и индивидуация",
           "Личность педагога и цель развития ребёнка", accent=JUNG)

    add_text(s, MARGIN_X, 196, 580, 18, [{
        "text": "Учитель в детском восприятии", "size": 11, "bold": True,
        "color": MUTED, "caps": True, "spc": 2,
    }])
    for x, w, label in ((MARGIN_X, 104, "Мудрец"), (MARGIN_X + 116, 104, "Мастер")):
        rect(s, x, 224, w, 30, fill=JUNG_TINT, line=None,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, x + 4, 231, w - 8, 18, [{
            "text": label, "size": 12.5, "color": JUNG, "align": PP_ALIGN.CENTER,
        }])

    paras = (
        [{"text": "На учителя проецируются "},
         {"text": "ожидания, надежды и страхи", "font": SERIF, "italic": True, "color": INK},
         {"text": " ребёнка. Педагог влияет прежде всего личностью, а не словами."}],
        [{"text": "Опасность — отождествить себя с идеальным, безошибочным Мудрецом: "
                  "тогда неудача ребёнка воспринимается как личное оскорбление."}],
        [{"text": "Самопознание учителя", "font": SERIF, "italic": True, "color": INK},
         {"text": " — часть его профессиональной работы."}],
    )
    add_text(s, MARGIN_X, 276, 580, 240, [{
        "runs": runs, "size": 15, "color": INK_SOFT, "line": 1.62, "after": 14,
    } for runs in paras])

    rect(s, 720, 196, 480, 300, fill=JUNG_TINT)
    rect(s, 720, 196, 3, 300, fill=JUNG)
    add_text(s, 752, 218, 420, 18, [{
        "text": "Цель развития", "size": 11, "bold": True, "color": MUTED,
        "caps": True, "spc": 2,
    }])
    add_text(s, 752, 246, 420, 100, [
        {"text": "стать", "size": 40, "font": SERIF, "italic": True, "color": MUTED, "line": 1.0},
        {"text": "самим собой", "size": 40, "font": SERIF, "italic": True, "color": INK, "line": 1.0},
    ])
    add_text(s, 752, 366, 416, 170, [{
        "text": item, "size": 14, "color": INK_SOFT, "line": 1.45,
        "bullet": "—", "bullet_color": JUNG, "indent": 16, "after": 8,
    } for item in (
        "Внутренняя целостность и согласие с собой",
        "Признание Тени — своих непризнанных качеств",
        "Раскрытие собственной неповторимости",
    )])

    footer(s, [
        {"text": "Учитель влияет не тем, что говорит, а "},
        {"text": "тем, кем он является", "bold": True, "color": INK},
        {"text": "."},
    ])

    notes(s, "Для Юнга учитель — тоже архетипическая фигура: на него проецируются "
             "ожидания, надежды и страхи ребёнка. Поэтому педагог влияет прежде всего "
             "личностью, а не словами. Опасность — отождествить себя с идеальным, "
             "безошибочным Мудрецом: тогда неудача ученика воспринимается как личное "
             "оскорбление. Самопознание учителя — часть профессии. Цель развития по "
             "Юнгу — индивидуация: стать самим собой, обрести внутреннюю целостность, "
             "признать свою Тень и раскрыть неповторимость.")


# --------------------------------------------------------------------------
# 8. Сравнение
# --------------------------------------------------------------------------
def slide_compare(prs):
    s = new_slide(prs)
    header(s, "Глава 3 · Сравнение", "Общее и различное")

    rect(s, MARGIN_X, 190, CONTENT_W, 46, fill=PANEL)
    add_text(s, 98, 203, 84, 20, [{
        "text": "Общее", "size": 11, "bold": True, "color": MUTED, "caps": True, "spc": 2,
    }])
    add_text(s, 186, 202, 990, 24, [{
        "text": "критика авторитарной школы · интерес к бессознательному · "
                "связь психологии и педагогики · уважение к личности ребёнка",
        "size": 14, "color": INK_SOFT,
    }])

    col_x = (MARGIN_X, 380, 780)
    col_w = (300, 400, 420)
    add_text(s, col_x[0], 260, col_w[0], 18, [{
        "text": "Критерий", "size": 11, "bold": True, "color": MUTED,
        "caps": True, "spc": 1.6,
    }])
    add_text(s, col_x[1], 260, col_w[1], 18, [{
        "text": "Адлер", "size": 11, "bold": True, "color": ADLER,
        "caps": True, "spc": 1.6,
    }])
    add_text(s, col_x[2], 260, col_w[2], 18, [{
        "text": "Юнг", "size": 11, "bold": True, "color": JUNG,
        "caps": True, "spc": 1.6,
    }])
    rule(s, MARGIN_X, 286, CONTENT_W, LINE_STRONG, 2)

    rows = (
        ("Движущая сила", "неполноценность и компенсация", "архетипы и индивидуация"),
        ("Ключевое понятие", "социальный интерес", "коллективное бессознательное"),
        ("Критерий здоровья", "сотрудничество, вклад", "целостность, согласие с собой"),
        ("Роль учителя", "партнёр и помощник", "носитель архетипа, пример"),
        ("Цель школы", "стать успешным вместе с другими", "стать самим собой"),
    )
    y = 288
    for i, (crit, a, j) in enumerate(rows):
        h = 52
        add_text(s, col_x[0], y + 15, col_w[0], 26, [{
            "text": crit, "size": 14, "bold": True, "color": INK,
        }])
        add_text(s, col_x[1], y + 15, col_w[1] - 20, 26, [{
            "text": a, "size": 14, "color": INK_SOFT,
        }])
        add_text(s, col_x[2], y + 15, col_w[2] - 20, 26, [{
            "text": j, "size": 14, "color": INK_SOFT,
        }])
        if i < len(rows) - 1:
            rule(s, MARGIN_X, y + h, CONTENT_W, LINE, 1)
        y += h

    footer(s, [
        {"text": "Формула: Адлер — "},
        {"text": "«стать успешным вместе с другими»", "bold": True, "color": INK},
        {"text": " · Юнг — "},
        {"text": "«стать самим собой»", "bold": True, "color": INK},
        {"text": "."},
    ])

    notes(s, "Сравним подходы. Общее: оба критикуют авторитарную школу, оба "
             "интересуются бессознательным, оба связывают психологию с педагогикой и "
             "уважают личность ребёнка. Различия — в движущей силе: у Адлера это "
             "неполноценность и компенсация, у Юнга — архетипы и индивидуация. "
             "Ключевые понятия — социальный интерес и коллективное бессознательное. "
             "Критерий здоровья: сотрудничество и вклад у Адлера, целостность у Юнга. "
             "Учитель — партнёр или носитель архетипа. Короткая формула: Адлер — "
             "стать успешным вместе с другими, Юнг — стать самим собой.")


# --------------------------------------------------------------------------
# 9. Современная школа
# --------------------------------------------------------------------------
def slide_modern(prs):
    s = new_slide(prs)
    header(s, "Практика", "Значение для современной школы",
           "Что школа берёт у Адлера и у Юнга")

    columns = (
        (MARGIN_X, ADLER, "От Адлера", "Отношения и мотивация", (
            "«трудное» поведение — сигнал, а не характер",
            "поддержка и поощрение вместо наказания",
            "развитие сотрудничества и самоуправления",
            "внимание к самооценке и мотивации",
        )),
        (640, JUNG, "От Юнга", "Внутренний мир и тип", (
            "учёт индивидуальных стилей обучения",
            "ценность творчества, сказки, игры",
            "работа с символами и внутренним миром",
            "внимание к личности учителя",
        )),
    )
    for x, color, label, title, items in columns:
        rule(s, x, 200, 520, color, 2)
        add_text(s, x, 214, 520, 18, [{
            "text": label, "size": 11, "bold": True, "color": color,
            "caps": True, "spc": 2,
        }])
        add_text(s, x, 236, 520, 40, [{
            "text": title, "size": 26, "font": SERIF, "color": INK,
        }])
        add_text(s, x, 302, 500, 210, [{
            "text": item, "size": 16, "color": INK_SOFT, "line": 1.55,
            "bullet": "—", "bullet_color": color, "indent": 18, "after": 15,
        } for item in items])

    footer(s, [
        {"text": "Двойная опора учителя: "},
        {"text": "понимание отношений", "bold": True, "color": INK},
        {"text": " и "},
        {"text": "понимание внутреннего мира ребёнка", "bold": True, "color": INK},
        {"text": "."},
    ])

    notes(s, "Что из этого берёт современная школа. От Адлера — работу с "
             "отношениями и мотивацией: трудное поведение понимаем как сигнал, "
             "поддерживаем вместо наказания, развиваем сотрудничество и "
             "самоуправление, следим за самооценкой. От Юнга — внимание к внутреннему "
             "миру: учитываем индивидуальные стили обучения, ценим творчество, сказку "
             "и игру, работаем с символами и с личностью самого учителя. В итоге у "
             "педагога две опоры: понимание отношений и понимание внутреннего мира.")


# --------------------------------------------------------------------------
# 10. Выводы и источники
# --------------------------------------------------------------------------
def slide_conclusion(prs):
    s = new_slide(prs)
    header(s, "Итоги", "Выводы и источники")

    add_text(s, MARGIN_X, 196, 520, 150, [{
        "runs": [
            {"text": "Адлер:", "bold": True, "color": ADLER},
            {"text": " школа через отношения — неполноценность, компенсация, "
                     "социальный интерес.\n"},
            {"text": "Юнг:", "bold": True, "color": JUNG},
            {"text": " школа через внутренний мир — архетипы, типы, индивидуация."},
        ],
        "size": 15.5, "color": INK_SOFT, "line": 1.55,
    }])

    add_text(s, MARGIN_X, 372, 520, 130, [{
        "runs": [
            {"text": "Подходы "},
            {"text": "не противоречат, а дополняют", "bold": True, "color": INK},
            {"text": " друг друга: «стать успешным вместе с другими» и «стать самим "
                     "собой» — две стороны одной задачи — помочь ребёнку вырасти "
                     "целостным человеком."},
        ],
        "size": 15, "color": INK_SOFT, "line": 1.6,
    }])

    add_text(s, 660, 196, 540, 18, [{
        "text": "Источники", "size": 11, "bold": True, "color": MUTED,
        "caps": True, "spc": 2,
    }])
    sources = (
        "Леонтьев Д. А. Адлер Альфред // БРЭ. URL: old.bigenc.ru/psychology/text/v/1800877",
        "Леонтьев Д. А. Индивидуальная психология // БРЭ. URL: old.bigenc.ru/psychology/text/v/2008679",
        "Мавлявиева Л. И. Личностно-социальная концепция воспитания А. Адлера : дис. … канд. пед. наук. URL: dslib.net/obw-pedagogika/lichnostno-socialnaja-koncepcija-vospitanija-alfreda-adlera.html",
        "Юнг (Jung) Карл Густав // Психологический словарь. URL: rus-big-psyho.slovaronline.com/3359-Юнг",
        "Хьелл Л., Зиглер Д. Теории личности (фрагмент). URL: gbs.spb.ru/ru/search/detail/?id=61e0492f12935ed6424cf239c090baf3",
        "Модель К. Юнга и современные стили обучения // КиберЛенинка. URL: cyberleninka.ru/article/n/model-k-yunga-i-sovremennye-stili-obucheniya",
        "Архетипическая педагогика (обзор) // Alphapedia. URL: alphapedia.ru/w/Archetypal_pedagogy",
    )
    add_text(s, 660, 224, 540, 330, [{
        "text": "%d. %s" % (i + 1, text),
        "size": 11, "color": INK_SOFT, "line": 1.32, "after": 7,
        "hanging": True, "indent": 16,
    } for i, text in enumerate(sources)])

    rule(s, MARGIN_X, 556, CONTENT_W, LINE, 1)
    add_text(s, MARGIN_X, 576, 640, 46, [{
        "text": "Спасибо за внимание", "size": 30, "font": SERIF,
        "italic": True, "color": INK,
    }])
    add_text(s, 760, 590, 440, 24, [{
        "text": "Сабадаш П. А. · ИС-24 · 2026", "size": 12.5,
        "color": MUTED, "align": PP_ALIGN.RIGHT,
    }])

    notes(s, "Итог. У Адлера школа — это отношения: неполноценность, компенсация, "
             "социальный интерес. У Юнга школа — это внутренний мир: архетипы, типы, "
             "индивидуация. Эти подходы не противоречат, а дополняют друг друга: "
             "стать успешным вместе с другими и стать самим собой — две стороны одной "
             "задачи, помочь ребёнку вырасти целостным человеком. Спасибо за внимание, "
             "готов ответить на вопросы.")


# --------------------------------------------------------------------------
def build():
    prs = Presentation()
    prs.slide_width = px(W)
    prs.slide_height = px(H)

    slide_cover(prs)
    slide_context(prs)
    slide_inferiority(prs)
    slide_social(prs)
    slide_archetypes(prs)
    slide_types(prs)
    slide_individuation(prs)
    slide_compare(prs)
    slide_modern(prs)
    slide_conclusion(prs)

    core = prs.core_properties
    core.title = "Психология школы в подходах Альфреда Адлера и Карла Густава Юнга"
    core.subject = "Доклад по психологии, тема №8"
    core.author = "Сабадаш Павел Андреевич, группа ИС-24, РКСИ, 2026"
    core.keywords = "Адлер, Юнг, психология школы, индивидуальная психология, аналитическая психология"
    core.comments = "10 слайдов, 16:9. Дизайн повторяет HTML-версию presentation/index.html."

    OUT.parent.mkdir(parents=True, exist_ok=True)
    prs.save(OUT)
    print("Готово:", OUT, "—", len(prs.slides._sldIdLst), "слайдов")


if __name__ == "__main__":
    build()
