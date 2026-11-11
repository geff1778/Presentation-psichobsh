# -*- coding: utf-8 -*-
"""Сборка презентации: presentation/index.html + presentation/slides/*.svg

Дизайн рассчитан на светлый кабинет: белый фон, тёмный текст,
акцентные цвета #0D6EFD (синий) и #6C757D (серый), плюс бирюзовый
#0F766E для раздела Юнга. Шрифт без засечек (system-ui).
Всё работает офлайн: SVG встроены прямо в index.html, внешних
зависимостей и CDN нет.

Запуск:  python3 tools/build_slides.py
"""

import os
from xml.sax.saxutils import escape

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_HTML = os.path.join(ROOT, "presentation", "index.html")
OUT_SLIDES = os.path.join(ROOT, "presentation", "slides")

W, H = 1280, 720
FONT = "system-ui,-apple-system,'Segoe UI',Roboto,Arial,sans-serif"

BG = "#FFFFFF"
SOFT = "#F8F9FA"
PANEL = "#F1F5F9"
BORDER = "#E2E8F0"
INK = "#212529"
INK2 = "#343A40"
MUTED = "#6C757D"
BLUE = "#0D6EFD"
BLUE_D = "#0A58CA"
BLUE_L = "#E7F1FF"
TEAL = "#0F766E"
TEAL_L = "#E4F3F0"
AMBER = "#B45309"
AMBER_L = "#FEF3E2"


def t(x, y, s, size=22, fill=INK2, weight="normal", anchor="start",
      spacing=None, opacity=None):
    a = ' text-anchor="%s"' % anchor if anchor != "start" else ""
    w = ' font-weight="%s"' % weight if weight != "normal" else ""
    sp = ' letter-spacing="%s"' % spacing if spacing else ""
    op = ' opacity="%s"' % opacity if opacity else ""
    return ('<text x="%s" y="%s" font-size="%s" fill="%s"%s%s%s>%s</text>'
            % (x, y, size, fill, w, a, sp + op, escape(s)))


def lines(x, y, items, size=22, fill=INK2, lh=32, weight="normal",
          anchor="start"):
    """Несколько строк одним <text> через tspan (ручной перенос)."""
    a = ' text-anchor="%s"' % anchor if anchor != "start" else ""
    w = ' font-weight="%s"' % weight if weight != "normal" else ""
    out = ['<text x="%s" y="%s" font-size="%s" fill="%s"%s%s>'
           % (x, y, size, fill, w, a)]
    for i, s in enumerate(items):
        out.append('<tspan x="%s" dy="%s">%s</tspan>'
                   % (x, 0 if i == 0 else lh, escape(s)))
    out.append("</text>")
    return "".join(out)


def card(x, y, w, h, fill="#FFFFFF", rx=18, stroke=BORDER, shadow=True):
    out = []
    if shadow:
        out.append('<rect x="%s" y="%s" width="%s" height="%s" rx="%s" '
                   'fill="#E9EEF3"/>' % (x + 5, y + 6, w, h, rx))
    out.append('<rect x="%s" y="%s" width="%s" height="%s" rx="%s" '
               'fill="%s" stroke="%s" stroke-width="1.5"/>'
               % (x, y, w, h, rx, fill, stroke))
    return "".join(out)


def chip(x, y, label, fill=BLUE_L, color=BLUE_D, size=19, pad=18, h=42):
    w = int(len(label) * size * 0.56) + pad * 2
    return ('<rect x="%s" y="%s" width="%s" height="%s" rx="%s" fill="%s"/>'
            % (x, y, w, h, h // 2, fill)
            + t(x + pad, y + h - 13, label, size=size, fill=color,
                weight="600"))


def bullet(x, y, s, size=21, fill=INK2, color=BLUE, lh=None):
    out = ['<circle cx="%s" cy="%s" r="5" fill="%s"/>'
           % (x + 5, y - 7, color)]
    out.append(t(x + 26, y, s, size=size, fill=fill))
    return "".join(out)


def bullets(x, y, items, size=21, lh=44, color=BLUE, fill=INK2):
    out = []
    for i, s in enumerate(items):
        out.append(bullet(x, y + i * lh, s, size=size, fill=fill, color=color))
    return "".join(out)


def header(title, sub=None, accent=BLUE, kicker=None):
    out = ['<rect x="0" y="0" width="%s" height="7" fill="%s"/>' % (W, accent)]
    if kicker:
        out.append(t(72, 62, kicker, size=17, fill=MUTED, weight="600",
                     spacing="1.5"))
    out.append(t(72, 118 if kicker else 104, title, size=40, fill=INK,
                 weight="700"))
    if sub:
        out.append(t(72, (162 if kicker else 148), sub, size=21, fill=MUTED))
    return "".join(out)


def svg_wrap(name, inner):
    return ('<svg xmlns="http://www.w3.org/2000/svg" width="%s" height="%s" '
            'viewBox="0 0 %s %s" font-family="%s" lang="ru-RU" '
            'role="img" aria-label="%s">%s</svg>'
            % (W, H, W, H, FONT, escape(name), inner))


# ---------------------------------------------------------------------------
# Слайды
# ---------------------------------------------------------------------------

def slide1():
    s = []
    s.append('<rect x="0" y="0" width="%s" height="%s" fill="%s"/>'
             % (W, H, BG))
    s.append('<circle cx="1180" cy="90" r="230" fill="%s" opacity="0.55"/>'
             % BLUE_L)
    s.append('<circle cx="120" cy="660" r="200" fill="%s" opacity="0.5"/>'
             % TEAL_L)
    s.append('<rect x="0" y="0" width="%s" height="8" fill="%s"/>' % (W, BLUE))
    s.append('<rect x="0" y="8" width="640" height="4" fill="%s"/>' % TEAL)
    s.append(t(72, 96, "Доклад по психологии · тема №8", size=19, fill=MUTED,
               weight="600", spacing="1.6"))
    s.append(lines(72, 210, [
        "Психология школы",
        "в подходах Альфреда Адлера",
        "и Карла Густава Юнга",
    ], size=54, fill=INK, lh=70, weight="700"))
    s.append('<rect x="76" y="452" width="120" height="6" rx="3" fill="%s"/>'
             % BLUE)
    s.append(lines(72, 508, [
        "Индивидуальная психология А. Адлера и аналитическая психология",
        "К. Г. Юнга о ребёнке, учителе и обучении",
    ], size=24, fill=MUTED, lh=36))
    s.append(t(72, 640, "Сабадаш Павел Андреевич", size=23, fill=INK2,
               weight="600"))
    s.append(t(72, 674, "Группа ИС-24 · РКСИ · 2026", size=20, fill=MUTED))
    return svg_wrap("Титульный слайд", "".join(s))


def slide2():
    s = [header("Почему школа стала объектом психологии",
                "Начало XX века: массовая школа, детство и бессознательное",
                kicker="КОНТЕКСТ", accent=BLUE)]
    x, y, w, h = 72, 196, 552, 396
    s.append(card(x, y, w, h))
    s.append(t(x + 32, y + 56, "Школа в центре внимания", size=25, fill=INK,
               weight="700"))
    s.append(bullets(x + 32, y + 112, [
        "Массовая школа и обязательное обучение —",
        "огромный поток детей впервые рядом друг с другом.",
        "Школа — первый институт после семьи, где",
        "ребёнок переживает успех и неудачу.",
        "Психоанализ выходит за пределы клиники:",
        "возникает интерес к детству и обучению.",
    ], size=20, lh=38))
    # timeline
    tx, ty, tw = 664, 196, 544
    s.append(card(tx, ty, tw, 396, fill=SOFT))
    s.append(t(tx + 32, ty + 56, "Ключевые вехи", size=25, fill=INK,
               weight="700"))
    nodes = [
        ("1900-е", "Массовая школа в Европе", BLUE),
        ("1911", "А. Адлер создаёт индивидуальную психологию", TEAL),
        ("1913", "К. Г. Юнг: аналитическая психология", TEAL),
        ("1919–20", "Венские медико-педагогические консультации", BLUE),
    ]
    ly = ty + 120
    s.append('<line x1="%s" y1="%s" x2="%s" y2="%s" stroke="%s" '
             'stroke-width="3"/>' % (tx + 74, ly - 4, tx + 74, ly + 232,
                                     BORDER))
    for i, (year, txt, col) in enumerate(nodes):
        cy = ly + i * 72
        s.append('<circle cx="%s" cy="%s" r="11" fill="%s"/>'
                 % (tx + 74, cy, col))
        s.append(t(tx + 110, cy - 2, year, size=21, fill=col, weight="700"))
        s.append(t(tx + 110, cy + 24, txt, size=18, fill=MUTED))
    s.append(t(72, 656, "Вывод: школа становится пространством, где "
                        "формируется личность ребёнка.", size=21,
               fill=INK2, weight="600"))
    return svg_wrap("Почему школа стала объектом психологии", "".join(s))


def slide3():
    s = [header("Адлер: чувство неполноценности",
                "Компенсация как источник развития", kicker="ГЛАВА 1 · АДЛЕР",
                accent=BLUE)]
    s.append(card(72, 196, 1136, 116, fill=BLUE_L, stroke=BLUE_L))
    s.append('<rect x="72" y="196" width="8" height="116" rx="4" fill="%s"/>'
             % BLUE)
    s.append(lines(104, 238, [
        "Чувство неполноценности — нормальное переживание каждого",
        "ребёнка: он мал, слаб и зависим от взрослых.",
    ], size=23, fill=INK, lh=34, weight="600"))
    s.append(t(104, 300, "Это не болезнь, а стимул к развитию.",
               size=21, fill=BLUE_D, weight="600"))
    # two paths
    s.append(card(72, 344, 552, 250))
    s.append('<rect x="72" y="344" width="552" height="10" rx="5" fill="%s"/>'
             % TEAL)
    s.append(t(104, 400, "Конструктивная компенсация", size=23, fill=TEAL,
               weight="700"))
    s.append(bullets(104, 448, [
        "преодоление трудности, рост усилий;",
        "поиск другой области успеха;",
        "уверенность и вклад в общее дело.",
    ], size=20, lh=42, color=TEAL))
    s.append(card(656, 344, 552, 250))
    s.append('<rect x="656" y="344" width="552" height="10" rx="5" '
             'fill="%s"/>' % AMBER)
    s.append(t(688, 400, "Деструктивная компенсация", size=23, fill=AMBER,
               weight="700"))
    s.append(bullets(688, 448, [
        "агрессия, демонстративность, обман;",
        "лень как защита: «способный, просто ленится»;",
        "уход в болезнь, отказ от деятельности.",
    ], size=20, lh=42, color=AMBER))
    s.append(t(72, 654, "Для учителя: сравнение и унижение толкают ребёнка "
                        "к деструктивной компенсации.", size=21, fill=INK2,
               weight="600"))
    return svg_wrap("Адлер: чувство неполноценности и компенсация",
                    "".join(s))


def slide4():
    s = [header("Адлер: социальный интерес и школа",
                "Gemeinschaftsgefühl — чувство общности",
                kicker="ГЛАВА 1 · АДЛЕР", accent=BLUE)]
    s.append(card(72, 196, 500, 396, fill=BLUE_L, stroke=BLUE_L))
    s.append(t(104, 252, "Социальный интерес", size=25, fill=BLUE_D,
               weight="700"))
    s.append(lines(104, 300, [
        "Врождённая склонность быть частью",
        "сообщества, сотрудничать и вносить",
        "вклад в общее дело.",
    ], size=21, fill=INK, lh=34))
    s.append('<line x1="104" y1="418" x2="540" y2="418" stroke="%s" '
             'stroke-width="2"/>' % BLUE)
    s.append(t(104, 458, "Критерий психического здоровья:", size=20,
               fill=MUTED))
    s.append(t(104, 496, "связь с людьми, а не превосходство.", size=22,
               fill=BLUE_D, weight="600"))
    s.append(t(104, 552, "Задача школы — воспитывать", size=20, fill=MUTED))
    s.append(t(104, 582, "сотрудничество, а не соперничество.", size=22,
               fill=INK, weight="600"))
    # practice
    s.append(card(612, 196, 596, 396))
    s.append(t(644, 252, "Школьная практика Адлера", size=25, fill=INK,
               weight="700"))
    items = [
        ("1919–1920-е", ["Венские медико-педагогические консультации:",
                         "учитель + врач + психолог"]),
        ("Экспериментальная школа", ["без наказаний, с самоуправлением",
                                     "и общими делами"]),
        ("Сегодня", ["поддержка вместо наказания;",
                     "«трудное» поведение — сигнал"]),
    ]
    yy = 306
    for head, body_lines in items:
        s.append('<circle cx="656" cy="%s" r="7" fill="%s"/>' % (yy - 6, BLUE))
        s.append(t(680, yy, head, size=21, fill=BLUE_D, weight="700"))
        s.append(lines(680, yy + 28, body_lines, size=19, fill=MUTED, lh=26))
        yy += 100
    s.append(t(72, 654, "Адлер: школа должна дать ребёнку переживание "
                        "нужности и принадлежности.", size=21, fill=INK2,
               weight="600"))
    return svg_wrap("Адлер: социальный интерес и школьная практика",
                    "".join(s))


def slide5():
    s = [header("Юнг: коллективное бессознательное",
                "Архетипы как универсальные первообразы",
                kicker="ГЛАВА 2 · ЮНГ", accent=TEAL)]
    # layers
    s.append(card(72, 196, 520, 396))
    s.append(t(104, 248, "Структура психики", size=24, fill=INK,
               weight="700"))
    layers = [
        ("Сознание", "#FFFFFF", INK, BORDER),
        ("Личное бессознательное", TEAL_L, TEAL, TEAL_L),
        ("Коллективное бессознательное", TEAL, "#FFFFFF", TEAL),
    ]
    ly = 292
    for i, (label, fill, col, stroke) in enumerate(layers):
        s.append('<rect x="104" y="%s" width="456" height="78" rx="12" '
                 'fill="%s" stroke="%s" stroke-width="1.5"/>'
                 % (ly, fill, stroke))
        s.append(t(132, ly + 49, label, size=21, fill=col,
                   weight="600" if i != 2 else "700"))
        ly += 92
    s.append(t(104, 572, "Слой, общий для всех людей,", size=19, fill=MUTED))
    s.append(t(104, 596, "наследуется как структура опыта.", size=19,
               fill=MUTED))
    # archetypes
    s.append(card(624, 196, 584, 396, fill=SOFT))
    s.append(t(656, 248, "Архетипы — «пустые формы»", size=24, fill=INK,
               weight="700"))
    s.append(t(656, 286, "наполняются личным опытом ребёнка", size=19,
               fill=MUTED))
    arch = ["Тень", "Персона", "Самость", "Мудрец", "Великая Мать",
            "Анима / Анимус"]
    ax, ay = 656, 330
    for i, a in enumerate(arch):
        cx = ax + (i % 2) * 268
        cy = ay + (i // 2) * 66
        s.append(chip(cx, cy, a, fill=TEAL_L, color=TEAL, size=19))
    s.append('<line x1="656" y1="524" x2="1176" y2="524" stroke="%s" '
             'stroke-width="2"/>' % BORDER)
    s.append(lines(656, 560, [
        "Сказка, миф, игра, любимый герой —",
        "встреча ребёнка с общечеловеческими образами.",
    ], size=19, fill=INK2, lh=30))
    s.append(t(72, 654, "Отсюда — ценность творчества и работы с символами "
                        "в детском возрасте.", size=21, fill=INK2,
               weight="600"))
    return svg_wrap("Юнг: коллективное бессознательное и архетипы",
                    "".join(s))


def slide6():
    s = [header("Юнг: психологические типы и обучение",
                "Экстраверсия и интроверсия · четыре функции",
                kicker="ГЛАВА 2 · ЮНГ", accent=TEAL)]
    s.append(card(72, 196, 552, 396))
    s.append(t(104, 248, "Две установки", size=24, fill=INK, weight="700"))
    s.append(chip(104, 288, "Экстраверсия", fill=TEAL_L, color=TEAL))
    s.append(lines(104, 372, [
        "Энергия направлена вовне.",
        "Учится через общение, движение,",
        "обсуждение и совместную работу.",
    ], size=20, fill=INK2, lh=32))
    s.append(chip(104, 470, "Интроверсия", fill=BLUE_L, color=BLUE_D))
    s.append(lines(104, 552, [
        "Энергия направлена внутрь.",
        "Учится через размышление и паузу.",
    ], size=20, fill=INK2, lh=32))
    s.append(card(656, 196, 552, 396, fill=SOFT))
    s.append(t(688, 248, "Четыре функции", size=24, fill=INK, weight="700"))
    funcs = [
        ("Мышление", "логика, структура, правила", BLUE),
        ("Чувство", "принятие, личное отношение", TEAL),
        ("Ощущение", "наглядность, практика", BLUE),
        ("Интуиция", "смысл, догадка, перспектива", TEAL),
    ]
    fy = 300
    for name, desc, col in funcs:
        s.append('<circle cx="700" cy="%s" r="7" fill="%s"/>' % (fy - 6, col))
        s.append(t(724, fy, name, size=21, fill=INK, weight="700"))
        s.append(t(724, fy + 28, desc, size=19, fill=MUTED))
        fy += 72
    s.append(t(72, 654, "Вывод: единый метод обучения не годится для всех — "
                        "отсюда идея стилей обучения.", size=21, fill=INK2,
               weight="600"))
    return svg_wrap("Юнг: психологические типы и обучение", "".join(s))


def slide7():
    s = [header("Юнг: учитель как архетип и индивидуация",
                "Личность учителя и цель развития ребёнка",
                kicker="ГЛАВА 2 · ЮНГ", accent=TEAL)]
    s.append(card(72, 196, 552, 396))
    s.append(t(104, 248, "Учитель как архетип", size=24, fill=INK,
               weight="700"))
    s.append(chip(104, 288, "Мудрец", fill=TEAL_L, color=TEAL))
    s.append(t(240, 314, "— носитель знания", size=19, fill=MUTED))
    s.append(chip(104, 344, "Мастер", fill=BLUE_L, color=BLUE_D))
    s.append(t(240, 370, "— носитель умения", size=19, fill=MUTED))
    s.append('<line x1="104" y1="416" x2="592" y2="416" stroke="%s" '
             'stroke-width="2"/>' % BORDER)
    s.append(bullets(104, 460, [
        "на учителя проецируются ожидания и страхи;",
        "влияет прежде всего личностью, не словами;",
        "самопознание учителя — часть работы.",
    ], size=19, lh=40, color=TEAL))
    s.append(card(656, 196, 552, 396, fill=TEAL_L, stroke=TEAL_L))
    s.append(t(688, 248, "Индивидуация", size=24, fill=TEAL, weight="700"))
    s.append(lines(688, 296, [
        "Цель развития — стать самим собой",
        "и обрести внутреннюю целостность.",
    ], size=21, fill=INK, lh=34))
    s.append('<line x1="688" y1="378" x2="1176" y2="378" stroke="%s" '
             'stroke-width="2"/>' % TEAL)
    s.append(t(688, 416, "Включает признание Тени —", size=20, fill=INK2))
    s.append(t(688, 446, "своих непризнанных качеств.", size=20, fill=INK2))
    s.append(t(688, 512, "Цель школы — помочь каждому", size=20, fill=MUTED))
    s.append(t(688, 542, "раскрыть собственную", size=20, fill=MUTED))
    s.append(t(688, 572, "неповторимость.", size=22, fill=TEAL,
               weight="700"))
    s.append(t(72, 654, "Юнг: учитель влияет не тем, что говорит, а тем, "
                        "кем он является.", size=21, fill=INK2, weight="600"))
    return svg_wrap("Юнг: учитель как архетип и индивидуация", "".join(s))


def slide8():
    s = [header("Общее и различное", "Адлер и Юнг: сопоставление",
                kicker="ГЛАВА 3 · СРАВНЕНИЕ", accent=BLUE)]
    # common band
    s.append(card(72, 184, 1136, 104, fill=BLUE_L, stroke=BLUE_L))
    s.append(t(104, 224, "Общее:", size=21, fill=BLUE_D, weight="700"))
    s.append(lines(230, 216, [
        "критика авторитарной школы · интерес к бессознательному ·",
        "связь психологии и педагогики · уважение к ребёнку",
    ], size=19, fill=INK2, lh=28))
    # table
    tx, ty, tw = 72, 300, 1136
    col1, col2 = 260, 438
    rows = [
        ("Критерий", "АДЛЕР", "ЮНГ"),
        ("Движущая сила", "неполноценность и компенсация",
         "архетипы и индивидуация"),
        ("Ключевое понятие", "социальный интерес",
         "коллективное бессознательное"),
        ("Критерий здоровья", "сотрудничество, вклад",
         "целостность, согласие с собой"),
        ("Роль учителя", "партнёр и помощник",
         "носитель архетипа, пример"),
        ("Цель школы", "стать успешным вместе с другими",
         "стать самим собой"),
    ]
    s.append(card(tx, ty, tw, 300, fill="#FFFFFF", rx=14))
    # header row
    s.append('<rect x="%s" y="%s" width="%s" height="52" rx="14" fill="%s"/>'
             % (tx, ty, tw, PANEL))
    s.append('<rect x="%s" y="%s" width="%s" height="30" fill="%s"/>'
             % (tx, ty + 22, tw, PANEL))
    s.append(t(tx + 24, ty + 34, rows[0][0], size=19, fill=MUTED,
               weight="700"))
    s.append(t(tx + col1 + 24, ty + 34, rows[0][1], size=20, fill=BLUE_D,
               weight="700"))
    s.append(t(tx + col1 + col2 + 24, ty + 34, rows[0][2], size=20, fill=TEAL,
               weight="700"))
    ry = ty + 52
    for i, (c, a, j) in enumerate(rows[1:]):
        if i % 2 == 1:
            s.append('<rect x="%s" y="%s" width="%s" height="49" fill="%s"/>'
                     % (tx, ry, tw, SOFT))
        s.append(t(tx + 24, ry + 32, c, size=19, fill=INK2, weight="600"))
        s.append(t(tx + col1 + 24, ry + 32, a, size=19, fill=INK2))
        s.append(t(tx + col1 + col2 + 24, ry + 32, j, size=19, fill=INK2))
        ry += 49
        if i < len(rows[1:]) - 1:
            s.append('<line x1="%s" y1="%s" x2="%s" y2="%s" stroke="%s" '
                     'stroke-width="1"/>' % (tx + 20, ry, tx + tw - 20, ry,
                                             BORDER))
    s.append('<line x1="%s" y1="%s" x2="%s" y2="%s" stroke="%s" '
             'stroke-width="2"/>' % (tx + col1, ty, tx + col1, ty + 300,
                                     BORDER))
    s.append('<line x1="%s" y1="%s" x2="%s" y2="%s" stroke="%s" '
             'stroke-width="2"/>' % (tx + col1 + col2, ty, tx + col1 + col2,
                                     ty + 300, BORDER))
    s.append(t(72, 654, "Адлер: «стать успешным вместе с другими» · Юнг: "
                        "«стать самим собой»", size=21, fill=INK2,
               weight="600"))
    return svg_wrap("Общее и различное", "".join(s))


def slide9():
    s = [header("Значение для современной школы",
                "Что школа берёт у Адлера и у Юнга",
                kicker="ПРАКТИКА", accent=BLUE)]
    s.append(card(72, 196, 552, 396))
    s.append('<rect x="72" y="196" width="552" height="10" rx="5" fill="%s"/>'
             % BLUE)
    s.append(t(104, 254, "От Адлера", size=25, fill=BLUE_D, weight="700"))
    s.append(bullets(104, 312, [
        "«трудное» поведение — сигнал, а не характер;",
        "поддержка и поощрение вместо наказания;",
        "развитие сотрудничества и самоуправления;",
        "внимание к самооценке и мотивации.",
    ], size=20, lh=56, color=BLUE))
    s.append(card(656, 196, 552, 396))
    s.append('<rect x="656" y="196" width="552" height="10" rx="5" '
             'fill="%s"/>' % TEAL)
    s.append(t(688, 254, "От Юнга", size=25, fill=TEAL, weight="700"))
    s.append(bullets(688, 312, [
        "учёт индивидуальных стилей обучения;",
        "ценность творчества, сказки и игры;",
        "работа с символами и внутренним миром;",
        "внимание к личности учителя.",
    ], size=20, lh=56, color=TEAL))
    s.append(card(72, 612, 1136, 68, fill=SOFT, stroke=SOFT))
    s.append(t(104, 656, "Двойная опора учителя: понимание отношений и "
                         "понимание внутреннего мира ребёнка.", size=21,
               fill=INK, weight="600"))
    return svg_wrap("Значение для современной школы", "".join(s))


def slide10():
    s = [header("Выводы и источники", "Итоги доклада", kicker="ИТОГИ",
                accent=BLUE)]
    s.append(card(72, 196, 660, 396))
    s.append(t(104, 250, "Выводы", size=25, fill=INK, weight="700"))
    s.append(bullets(104, 306, [
        "Адлер: школа через отношения —",
        "неполноценность, компенсация, социальный интерес.",
        "Юнг: школа через внутренний мир —",
        "архетипы, психологические типы, индивидуация.",
        "Подходы не противоречат, а дополняют друг друга:",
        "успех вместе с другими и верность себе.",
    ], size=20, lh=44, color=BLUE))
    s.append(card(756, 196, 452, 396, fill=SOFT))
    s.append(t(788, 250, "Источники", size=25, fill=INK, weight="700"))
    src = [
        "Д. А. Леонтьев. Адлер Альфред (БРЭ)",
        "Леонтьев. Индивидуальная психология",
        "Мавлявиева. Дис. о концепции Адлера",
        "К. Г. Юнг (Психологический словарь)",
        "Л. Хьелл, Д. Зиглер. Теории личности",
        "Модель Юнга и стили обучения",
        "Архетипическая педагогика",
    ]
    sy = 306
    for i, item in enumerate(src, 1):
        s.append('<circle cx="800" cy="%s" r="4" fill="%s"/>'
                 % (sy - 6, BLUE))
        s.append(t(820, sy, "%d. %s" % (i, item), size=18, fill=INK2))
        sy += 38
    s.append(t(72, 654, "Спасибо за внимание!", size=24, fill=BLUE_D,
               weight="700"))
    return svg_wrap("Выводы и источники", "".join(s))


SLIDES = [
    ("01_cover", slide1),
    ("02_why", slide2),
    ("03_adler_inferiority", slide3),
    ("04_adler_social", slide4),
    ("05_jung_unconscious", slide5),
    ("06_jung_types", slide6),
    ("07_jung_teacher", slide7),
    ("08_compare", slide8),
    ("09_modern_school", slide9),
    ("10_conclusion", slide10),
]

TITLES = [
    "Титульный", "Почему школа стала объектом психологии",
    "Адлер: чувство неполноценности и компенсация",
    "Адлер: социальный интерес и школьная практика",
    "Юнг: коллективное бессознательное и архетипы",
    "Юнг: психологические типы и обучение",
    "Юнг: учитель как архетип и индивидуация",
    "Общее и различное", "Значение для современной школы",
    "Выводы и источники",
]


HTML = """<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Психология школы: Адлер и Юнг — Сабадаш Павел Андреевич, ИС-24</title>
<style>
*{{margin:0;padding:0;box-sizing:border-box}}
html,body{{height:100%;background:#EEF2F6;overflow:hidden;
  font-family:{font}}}
#stage{{position:absolute;inset:0;display:flex;align-items:center;
  justify-content:center}}
.slide{{position:absolute;width:{w}px;height:{h}px;transform-origin:center;
  opacity:0;visibility:hidden;transition:opacity .35s ease;
  border-radius:10px;box-shadow:0 10px 40px rgba(15,23,42,.14);
  overflow:hidden;background:#fff}}
.slide.active{{opacity:1;visibility:visible}}
.slide svg{{display:block;width:{w}px;height:{h}px}}
.slide.active svg{{animation:rise .4s ease both}}
@keyframes rise{{from{{opacity:0;transform:translateY(14px)}}
  to{{opacity:1;transform:translateY(0)}}}}
#bar{{position:fixed;left:0;top:0;height:4px;background:{blue};width:0;
  z-index:60;transition:width .3s ease}}
#hud{{position:fixed;right:16px;bottom:14px;z-index:60;font-size:14px;
  color:#334155;background:rgba(255,255,255,.92);border:1px solid #E2E8F0;
  padding:6px 14px;border-radius:999px;font-variant-numeric:tabular-nums}}
#hud b{{color:{blue}}}
#help{{position:fixed;left:16px;bottom:14px;z-index:60;font-size:13px;
  color:#64748B;background:rgba(255,255,255,.9);border:1px solid #E2E8F0;
  padding:6px 12px;border-radius:999px}}
#help kbd{{font-family:inherit;background:#F1F5F9;border:1px solid #E2E8F0;
  border-radius:5px;padding:1px 5px;font-size:12px;color:#334155}}
@media print{{
  html,body{{overflow:visible;background:#fff}}
  #stage{{position:static;display:block}}
  .slide{{position:relative;opacity:1;visibility:visible;transform:none!important;
    box-shadow:none;border-radius:0;page-break-after:always}}
  #bar,#hud,#help{{display:none}}
}}
</style>
</head>
<body>
<div id="bar"></div>
<div id="stage">
{slides}
</div>
<div id="hud"><b id="cur">1</b> / {n}</div>
<div id="help"><kbd>&larr;</kbd> <kbd>&rarr;</kbd> слайды ·
  <kbd>F</kbd> полный экран · <kbd>Home</kbd>/<kbd>End</kbd></div>
<script>
var idx=0,slides=document.querySelectorAll('.slide'),n=slides.length;
var bar=document.getElementById('bar'),cur=document.getElementById('cur');
function fit(){{var s=Math.min(innerWidth/{w},innerHeight/{h})*0.94;
  slides.forEach(function(el){{el.style.transform='scale('+s+')';}});}}
function show(i){{idx=Math.max(0,Math.min(n-1,i));
  slides.forEach(function(el,k){{el.classList.toggle('active',k===idx);}});
  bar.style.width=((idx+1)/n*100)+'%';cur.textContent=idx+1;
  try{{history.replaceState(null,'','#'+(idx+1));}}catch(e){{}}}}
addEventListener('resize',fit);
addEventListener('keydown',function(e){{
  if(e.key==='ArrowRight'||e.key==='PageDown'||e.key===' '){{
    show(idx+1);e.preventDefault();}}
  else if(e.key==='ArrowLeft'||e.key==='PageUp'){{show(idx-1);e.preventDefault();}}
  else if(e.key==='Home'){{show(0);}}
  else if(e.key==='End'){{show(n-1);}}
  else if(e.key==='f'||e.key==='F'){{
    if(document.fullscreenElement){{document.exitFullscreen();}}
    else{{document.documentElement.requestFullscreen();}}}}}});
document.addEventListener('click',function(e){{
  if(e.target.closest('#hud')||e.target.closest('#help'))return;
  show(idx+(e.clientX<innerWidth*0.3?-1:1));}});
fit();show((parseInt((location.hash||'#1').slice(1),10)||1)-1);
</script>
</body>
</html>
"""


def main():
    os.makedirs(OUT_SLIDES, exist_ok=True)
    blocks = []
    for i, (name, fn) in enumerate(SLIDES, 1):
        svg = fn()
        with open(os.path.join(OUT_SLIDES, "%s.svg" % name), "w",
                  encoding="utf-8") as f:
            f.write('<?xml version="1.0" encoding="UTF-8"?>\n' + svg + "\n")
        blocks.append('<div class="slide" data-name="%s" '
                      'aria-label="%s">%s</div>'
                      % (name, escape(TITLES[i - 1]), svg))
    html = HTML.format(font=FONT, w=W, h=H, blue=BLUE, n=len(SLIDES),
                       slides="\n".join(blocks))
    with open(OUT_HTML, "w", encoding="utf-8") as f:
        f.write(html)
    print("saved", OUT_HTML, "(%d slides)" % len(SLIDES))
    print("saved %d svg -> %s" % (len(SLIDES), OUT_SLIDES))


if __name__ == "__main__":
    main()
