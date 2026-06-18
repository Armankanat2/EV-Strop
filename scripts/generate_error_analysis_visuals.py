# -*- coding: utf-8 -*-
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Iterable
import textwrap

from PIL import Image, ImageDraw, ImageFont


WIDTH = 1600
HEIGHT = 900
MARGIN = 48
HEADER_H = 110
FOOTER_H = 54
RIGHT_W = 500
CANVAS_BG = "#f6f2ea"
TEXT = "#1f2428"
MUTED = "#64707d"
WHITE = "#ffffff"
GRID = "#ded6ca"
DANGER = "#d23b3b"
SUCCESS = "#1f8a4c"
WARN = "#d97c17"
SAFE = "#2d6aa0"


ROOT = Path(__file__).resolve().parent.parent
ASSETS = ROOT / "assets" / "course-media"


FONT_CANDIDATES = [
    ("C:/Windows/Fonts/arialbd.ttf", True),
    ("C:/Windows/Fonts/arial.ttf", False),
    ("C:/Windows/Fonts/segoeuib.ttf", True),
    ("C:/Windows/Fonts/segoeui.ttf", False),
]


def load_font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    preferred = [p for p, is_bold in FONT_CANDIDATES if is_bold == bold] + [
        p for p, _ in FONT_CANDIDATES
    ]
    for path in preferred:
        if Path(path).exists():
            return ImageFont.truetype(path, size=size)
    return ImageFont.load_default()


TITLE_FONT = load_font(34, bold=True)
SUBTITLE_FONT = load_font(20, bold=False)
PANEL_TITLE_FONT = load_font(26, bold=True)
BODY_FONT = load_font(24, bold=False)
SMALL_FONT = load_font(19, bold=False)
TAG_FONT = load_font(18, bold=True)
BIG_FONT = load_font(52, bold=True)
ID_FONT = load_font(40, bold=True)


@dataclass
class SlideSpec:
    slide_id: str
    test_id: str
    module_slug: str
    module_title: str
    filename: str
    title: str
    kind: str
    accent: str
    bullets: list[str]


SPECS: list[SlideSpec] = [
    SlideSpec("S018-P01", "S018", "module-01-stropovka-gruzov", "Тема 2.1. Строповка грузов", "S018-P01_proverka-gruza_checklist.png", "Нет проверки груза перед подъемом", "checklist", "#9d6a1b", ["Оценить груз", "Проверить массу", "Проверить схему", "Проверить условия"]),
    SlideSpec("S018-P02", "S018", "module-01-stropovka-gruzov", "Тема 2.1. Строповка грузов", "S018-P02_massa-neizvestna_warning.png", "Работа с неизвестной массой груза", "warning_load", "#b25a26", ["Масса не подтверждена", "Подъем запрещен", "Сначала уточнить данные"]),
    SlideSpec("S018-P03", "S018", "module-01-stropovka-gruzov", "Тема 2.1. Строповка грузов", "S018-P03_tsentr-tyazhesti_diagram.png", "Не учтен центр тяжести", "center_gravity", "#8a5a18", ["Устойчивый подъем", "Смещенный центр тяжести", "Риск крена и разворота"]),
    SlideSpec("S018-P04", "S018", "module-01-stropovka-gruzov", "Тема 2.1. Строповка грузов", "S018-P04_ugol-vetvey_nagruzka-diagram.png", "Не учтен угол между ветвями стропа", "angle_load", "#8b5a18", ["Малый угол безопаснее", "Большой угол повышает нагрузку", "Схема подбирается заранее"]),
    SlideSpec("S018-P05", "S018", "module-01-stropovka-gruzov", "Тема 2.1. Строповка грузов", "S018-P05_ostrye-rebra_zashchita-slinga.png", "Нет защиты от острых ребер", "edge_protection", "#98591d", ["Без защиты нельзя", "Нужны прокладки", "Защита ставится до подъема"]),
    SlideSpec("S018-P06", "S018", "module-01-stropovka-gruzov", "Тема 2.1. Строповка грузов", "S018-P06_zatsepka-osnastka_checkpoints.png", "Не проверена зацепка и оснастка", "inspection_points", "#926323", ["Крюк", "Стропы", "Бирка", "Место зацепки"]),
    SlideSpec("S018-P07", "S018", "module-01-stropovka-gruzov", "Тема 2.1. Строповка грузов", "S018-P07_stop-pri-opasnosti_poster.png", "Работа не остановлена при опасности", "stop_poster", "#a64832", ["Опасность выявлена", "Работу остановить", "Риск устранить до продолжения"]),
    SlideSpec("S018-P08", "S018", "module-01-stropovka-gruzov", "Тема 2.1. Строповка грузов", "S018-P08_somnitelnaia-zatsepka_compare.png", "Сомнительная зацепка принята как допустимая", "compare", "#8f5d20", ["Сомнительно", "Допустимо", "Нужна проверенная схема"]),
    SlideSpec("S018-P09", "S018", "module-01-stropovka-gruzov", "Тема 2.1. Строповка грузов", "S018-P09_manipuliatsionnye-znaki_set.png", "Игнорируются манипуляционные знаки", "signs", "#896019", ["Верх", "Хрупкое", "Беречь от влаги", "Центр тяжести", "Не кантовать", "Место строповки"]),
    SlideSpec("S018-P10", "S018", "module-01-stropovka-gruzov", "Тема 2.1. Строповка грузов", "S018-P10_proverka-skhemy_steps.png", "Проверка схемы считается необязательной", "flow_steps", "#88601c", ["Проверка", "Сигнал", "Подъем", "Сначала схема, потом действие"]),
    SlideSpec("S018-P11", "S018", "module-01-stropovka-gruzov", "Тема 2.1. Строповка грузов", "S018-P11_massa-naverniaka_case.png", "Груз поднимают без подтвержденной массы", "mass_case", "#90561c", ["Нет маркировки", "Нельзя действовать наверняка", "Сначала уточнить массу"]),
    SlideSpec("S018-P12", "S018", "module-01-stropovka-gruzov", "Тема 2.1. Строповка грузов", "S018-P12_kontakt-s-ostrym-rebrom_detail.png", "Ветвь идет по острому ребру без защиты", "edge_detail", "#9b5b18", ["Опасный контакт", "Риск повреждения", "Подъем остановить"]),
    SlideSpec("S031-P01", "S031", "module-02-tekhnologiya-raboty-stropalshchika", "Тема 2.2. Технология работы стропальщика", "S031-P01_podgotovka-k-rabote_checklist.png", "Нет подготовки к работе", "checklist", "#2c6a8b", ["Уточнить задание", "Проверить условия", "Осмотреть оснастку", "Проверить зону"]),
    SlideSpec("S031-P02", "S031", "module-02-tekhnologiya-raboty-stropalshchika", "Тема 2.2. Технология работы стропальщика", "S031-P02_plohaia-osveshchennost_warning.png", "Работа начинается при плохой освещенности", "low_light", "#375f89", ["Видимость недостаточна", "Работу начинать нельзя", "Устранить условия до старта"]),
    SlideSpec("S031-P03", "S031", "module-02-tekhnologiya-raboty-stropalshchika", "Тема 2.2. Технология работы стропальщика", "S031-P03_obviazka-pered-podem_steps.png", "Не проверена обвязка перед подъемом", "flow_steps", "#2d6488", ["Обвязка", "Зацепка", "Контроль", "Сигнал"]),
    SlideSpec("S031-P04", "S031", "module-02-tekhnologiya-raboty-stropalshchika", "Тема 2.2. Технология работы стропальщика", "S031-P04_istochnik-signala_roles.png", "Сигналы подает неуполномоченный работник", "signal_roles", "#2a6384", ["Один источник сигнала", "Остальные не командуют", "Путаница исключается"]),
    SlideSpec("S031-P05", "S031", "module-02-tekhnologiya-raboty-stropalshchika", "Тема 2.2. Технология работы стропальщика", "S031-P05_stop-prioritet_poster.png", "Команда Стоп недооценена", "stop_poster", "#35607a", ["Команда обязательна", "Выполняется сразу", "Приоритет безопасности"]),
    SlideSpec("S031-P06", "S031", "module-02-tekhnologiya-raboty-stropalshchika", "Тема 2.2. Технология работы стропальщика", "S031-P06_neiasnyi-signal_stop.png", "Работа продолжается при неясном сигнале", "unclear_signal", "#356780", ["Сигнал неясен", "Видимость плохая", "Работу остановить"]),
    SlideSpec("S031-P07", "S031", "module-02-tekhnologiya-raboty-stropalshchika", "Тема 2.2. Технология работы стропальщика", "S031-P07_ottiagka_safe-guiding-diagram.png", "Оттяжка используется неправильно", "tagline_guiding", "#22698b", ["Сопровождение на дистанции", "Контроль раскачивания", "Не тянуть руками под подвесом"]),
    SlideSpec("S031-P08", "S031", "module-02-tekhnologiya-raboty-stropalshchika", "Тема 2.2. Технология работы стропальщика", "S031-P08_opasnaia-zona_movement-diagram.png", "Человек находится под грузом или в опасной зоне", "danger_zone", "#26698f", ["Под грузом находиться нельзя", "Траектория должна быть свободна", "Безопасная позиция вне зоны"]),
    SlideSpec("S031-P09", "S031", "module-02-tekhnologiya-raboty-stropalshchika", "Тема 2.2. Технология работы стропальщика", "S031-P09_skladovanie_podkladki-diagram.png", "Место укладки не подготовлено", "stacking", "#22698b", ["Основание", "Подкладки", "Устойчивое положение"]),
    SlideSpec("S031-P10", "S031", "module-02-tekhnologiya-raboty-stropalshchika", "Тема 2.2. Технология работы стропальщика", "S031-P10_rasstropovka_rano-pravilno.png", "Расстроповка начата слишком рано", "compare", "#2a6288", ["Рано", "Правильно", "Сначала устойчивая установка"]),
    SlideSpec("S031-P11", "S031", "module-02-tekhnologiya-raboty-stropalshchika", "Тема 2.2. Технология работы стропальщика", "S031-P11_raskachivanie-liudi_case-diagram.png", "Не остановлено перемещение при раскачивании и людях в зоне", "swing_people", "#296488", ["Груз раскачивается", "Люди в зоне", "Остановить и вывести людей"]),
    SlideSpec("S031-P12", "S031", "module-02-tekhnologiya-raboty-stropalshchika", "Тема 2.2. Технология работы стропальщика", "S031-P12_neustoichivyi-gruz_unhooking-ban.png", "Расстроповка при неустойчивом положении груза", "unstable_load", "#2b5e80", ["Груз перекошен", "Стропы не снимать", "Сначала выровнять"]),
    SlideSpec("S044-P01", "S044", "module-03-proizvodstvo-rabot", "Тема 2.3. Производство работ", "S044-P01_opasnaia-zona_diagram.png", "Опасная зона не определена до начала работ", "danger_zone", "#b24635", ["Опасную зону определить заранее", "Людей вывести", "Работу без этого не начинать"]),
    SlideSpec("S044-P02", "S044", "module-03-proizvodstvo-rabot", "Тема 2.3. Производство работ", "S044-P02_opasnaia-obstanovka_case.png", "Опасная обстановка оценивается слишком узко", "obstacles_case", "#b44c36", ["Траектория", "Раскачивание", "Помехи", "Люди вокруг"]),
    SlideSpec("S044-P03", "S044", "module-03-proizvodstvo-rabot", "Тема 2.3. Производство работ", "S044-P03_bezopasnye-usloviia_stop-card.png", "Работа продолжается без безопасных условий", "stop_card", "#b64c39", ["Нет безопасных условий", "Работу прекратить", "Риск устранить"]),
    SlideSpec("S044-P04", "S044", "module-03-proizvodstvo-rabot", "Тема 2.3. Производство работ", "S044-P04_lep-special-conditions_diagram.png", "Работа рядом с ЛЭП ведется на глаз", "lep_special", "#b54c31", ["ЛЭП — особая опасность", "Нужны точные границы", "Требуются специальные меры"]),
    SlideSpec("S044-P05", "S044", "module-03-proizvodstvo-rabot", "Тема 2.3. Производство работ", "S044-P05_lep-unclear-distance_stop.png", "Работа у ЛЭП продолжается при неясных дистанциях", "lep_stop", "#b84931", ["Дистанции неясны", "Работу остановить", "Уточнить безопасную зону"]),
    SlideSpec("S051-P01", "S051", "module-04-okhrana-truda", "Тема 3. Охрана труда", "S051-P01_rabochee-mesto-siz_card.png", "Охрана труда воспринимается слишком узко", "card_grid", "#2b7a59", ["Рабочее место", "СИЗ", "Правила", "Право остановить опасную работу"]),
    SlideSpec("S051-P02", "S051", "module-04-okhrana-truda", "Тема 3. Охрана труда", "S051-P02_obiazannosti-rabotnika_checklist.png", "Обязанности на рабочем месте понимаются неполно", "duties_checklist", "#2e7d56", ["Использовать СИЗ", "Выполнять инструкции", "Сообщать об опасности", "Не продолжать при риске"]),
    SlideSpec("S051-P03", "S051", "module-04-okhrana-truda", "Тема 3. Охрана труда", "S051-P03_dopusk-obuchenie_flow.png", "Допуск к работе понимается неправильно", "admission_flow", "#2f845c", ["Обучение", "Инструктаж", "Проверка знаний", "Допуск"]),
    SlideSpec("S051-P04", "S051", "module-04-okhrana-truda", "Тема 3. Охрана труда", "S051-P04_opasnost-stop-message_card.png", "При опасности действия откладываются", "stop_message", "#317f5a", ["Остановить работу", "Предупредить людей", "Сообщить ответственному"]),
    SlideSpec("S051-P05", "S051", "module-04-okhrana-truda", "Тема 3. Охрана труда", "S051-P05_neispravnost-people-stop_case.png", "Неисправность и люди в зоне не приводят к остановке", "defect_people_case", "#2d7b57", ["Есть неисправность", "Люди в зоне", "Работу остановить сразу"]),
]


def wrap(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.ImageFont, width: int) -> list[str]:
    words = text.split()
    lines: list[str] = []
    current = ""
    for word in words:
        trial = f"{current} {word}".strip()
        if draw.textbbox((0, 0), trial, font=font)[2] <= width:
            current = trial
        else:
            if current:
                lines.append(current)
            current = word
    if current:
        lines.append(current)
    return lines


def draw_wrapped(draw: ImageDraw.ImageDraw, xy: tuple[int, int], text: str, font: ImageFont.ImageFont, fill: str, width: int, line_gap: int = 8) -> int:
    x, y = xy
    for line in wrap(draw, text, font, width):
        draw.text((x, y), line, font=font, fill=fill)
        y += font.size + line_gap
    return y


def create_canvas(accent: str) -> tuple[Image.Image, ImageDraw.ImageDraw]:
    image = Image.new("RGB", (WIDTH, HEIGHT), CANVAS_BG)
    draw = ImageDraw.Draw(image)
    draw.rounded_rectangle((MARGIN, MARGIN, WIDTH - MARGIN, HEIGHT - MARGIN), 36, fill=WHITE, outline=GRID, width=2)
    draw.rounded_rectangle((MARGIN, MARGIN, WIDTH - MARGIN, MARGIN + HEADER_H), 36, fill=accent)
    draw.rounded_rectangle((MARGIN, HEIGHT - MARGIN - FOOTER_H, WIDTH - MARGIN, HEIGHT - MARGIN), 18, fill="#f2eee7")
    return image, draw


def header(draw: ImageDraw.ImageDraw, spec: SlideSpec) -> None:
    pill = (MARGIN + 28, MARGIN + 22, MARGIN + 238, MARGIN + 82)
    draw.rounded_rectangle(pill, 24, fill=WHITE)
    draw.text((MARGIN + 44, MARGIN + 29), spec.slide_id, font=ID_FONT, fill=spec.accent)
    draw.text((MARGIN + 272, MARGIN + 22), spec.title, font=TITLE_FONT, fill=WHITE)
    draw.text((MARGIN + 272, MARGIN + 65), spec.module_title, font=SUBTITLE_FONT, fill="#f6f2ea")
    footer_text = f"Возврат: {spec.slide_id} -> {spec.test_id}"
    draw.text((MARGIN + 26, HEIGHT - MARGIN - 38), footer_text, font=SMALL_FONT, fill=MUTED)


def panels() -> tuple[tuple[int, int, int, int], tuple[int, int, int, int]]:
    left = (MARGIN + 30, MARGIN + HEADER_H + 25, WIDTH - RIGHT_W - 60, HEIGHT - MARGIN - FOOTER_H - 24)
    right = (WIDTH - RIGHT_W - 10, MARGIN + HEADER_H + 25, WIDTH - MARGIN - 30, HEIGHT - MARGIN - FOOTER_H - 24)
    return left, right


def draw_panel_frame(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], title: str, accent: str) -> None:
    x1, y1, x2, y2 = box
    draw.rounded_rectangle(box, 28, fill="#faf8f4", outline="#e2ddd4", width=2)
    draw.rounded_rectangle((x1 + 18, y1 + 16, x1 + 220, y1 + 54), 16, fill=accent)
    draw.text((x1 + 34, y1 + 23), title, font=TAG_FONT, fill=WHITE)


def draw_bullets(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], title: str, bullets: Iterable[str], accent: str) -> None:
    draw_panel_frame(draw, box, title, accent)
    x1, y1, x2, _ = box
    y = y1 + 82
    width = x2 - x1 - 70
    for bullet in bullets:
        draw.ellipse((x1 + 28, y + 8, x1 + 40, y + 20), fill=accent)
        y = draw_wrapped(draw, (x1 + 56, y), bullet, BODY_FONT, TEXT, width) + 18


def draw_load(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], fill: str, label: str = "ГРУЗ") -> tuple[int, int, int, int]:
    x1, y1, x2, y2 = box
    pad_x = min(60, max(18, int((x2 - x1) * 0.14)))
    pad_y_top = min(54, max(18, int((y2 - y1) * 0.18)))
    pad_y_bottom = min(48, max(18, int((y2 - y1) * 0.12)))
    load = (x1 + pad_x, y1 + pad_y_top, x2 - pad_x, y2 - pad_y_bottom)
    draw.rounded_rectangle(load, 16, fill=fill, outline="#7b5e34", width=4)
    tx = (load[0] + load[2]) // 2
    ty = (load[1] + load[3]) // 2 - 20
    bbox = draw.textbbox((0, 0), label, font=TITLE_FONT)
    draw.text((tx - (bbox[2] - bbox[0]) // 2, ty), label, font=TITLE_FONT, fill=WHITE)
    return load


def draw_hook(draw: ImageDraw.ImageDraw, center_x: int, top_y: int, bottom_y: int, accent: str) -> None:
    draw.line((center_x, top_y, center_x, bottom_y - 40), fill=accent, width=10)
    draw.arc((center_x - 40, bottom_y - 80, center_x + 40, bottom_y), start=180, end=25, fill=accent, width=10)


def draw_person(draw: ImageDraw.ImageDraw, x: int, y: int, color: str) -> None:
    draw.ellipse((x - 18, y, x + 18, y + 36), fill=color)
    draw.line((x, y + 36, x, y + 96), fill=color, width=8)
    draw.line((x, y + 54, x - 32, y + 88), fill=color, width=7)
    draw.line((x, y + 54, x + 32, y + 88), fill=color, width=7)
    draw.line((x, y + 96, x - 28, y + 140), fill=color, width=7)
    draw.line((x, y + 96, x + 28, y + 140), fill=color, width=7)


def arrow(draw: ImageDraw.ImageDraw, start: tuple[int, int], end: tuple[int, int], color: str, width: int = 6) -> None:
    draw.line((*start, *end), fill=color, width=width)
    ex, ey = end
    sx, sy = start
    dx = ex - sx
    dy = ey - sy
    if dx == 0 and dy == 0:
        return
    scale = max((dx * dx + dy * dy) ** 0.5, 1)
    ux = dx / scale
    uy = dy / scale
    px = -uy
    py = ux
    tip = (ex, ey)
    left = (int(ex - ux * 28 + px * 14), int(ey - uy * 28 + py * 14))
    right = (int(ex - ux * 28 - px * 14), int(ey - uy * 28 - py * 14))
    draw.polygon([tip, left, right], fill=color)


def stop_octagon(draw: ImageDraw.ImageDraw, cx: int, cy: int, radius: int = 95) -> None:
    pts = []
    for i in range(8):
        ang = (22.5 + i * 45) * 3.14159265 / 180
        pts.append((int(cx + radius * __import__("math").cos(ang)), int(cy + radius * __import__("math").sin(ang))))
    draw.polygon(pts, fill=DANGER, outline="#8d1f1f")
    bbox = draw.textbbox((0, 0), "СТОП", font=BIG_FONT)
    draw.text((cx - (bbox[2] - bbox[0]) // 2, cy - (bbox[3] - bbox[1]) // 2 - 4), "СТОП", font=BIG_FONT, fill=WHITE)


def render_checklist(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], accent: str, bullets: list[str]) -> None:
    draw_panel_frame(draw, box, "Порядок", accent)
    x1, y1, x2, _ = box
    card_w = (x2 - x1 - 80) // 2
    card_h = 120
    positions = [
        (x1 + 28, y1 + 90),
        (x1 + 48 + card_w, y1 + 90),
        (x1 + 28, y1 + 232),
        (x1 + 48 + card_w, y1 + 232),
    ]
    for idx, (label, (cx, cy)) in enumerate(zip(bullets[:4], positions), start=1):
        draw.rounded_rectangle((cx, cy, cx + card_w, cy + card_h), 24, fill=WHITE, outline=accent, width=3)
        draw.ellipse((cx + 18, cy + 18, cx + 62, cy + 62), fill=accent)
        num = str(idx)
        nb = draw.textbbox((0, 0), num, font=PANEL_TITLE_FONT)
        draw.text((cx + 40 - (nb[2] - nb[0]) // 2, cy + 22), num, font=PANEL_TITLE_FONT, fill=WHITE)
        draw_wrapped(draw, (cx + 82, cy + 26), label, BODY_FONT, TEXT, card_w - 100)


def render_warning_load(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], accent: str) -> None:
    draw_panel_frame(draw, box, "Запрет", accent)
    load = draw_load(draw, box, accent)
    cx = (load[0] + load[2]) // 2
    top = box[1] + 80
    draw_hook(draw, cx, top, load[1] - 10, "#444444")
    tri = [(cx + 210, load[1] + 30), (cx + 310, load[1] + 210), (cx + 110, load[1] + 210)]
    draw.polygon(tri, fill=WARN, outline="#8f5511")
    draw.text((cx + 175, load[1] + 108), "!", font=BIG_FONT, fill=WHITE)
    draw.text((load[0] + 50, load[3] - 90), "масса ?", font=TITLE_FONT, fill=WHITE)


def render_center_gravity(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], accent: str) -> None:
    draw_panel_frame(draw, box, "Схема", accent)
    left = (box[0] + 40, box[1] + 90, box[0] + (box[2] - box[0]) // 2 - 20, box[3] - 40)
    right = (box[0] + (box[2] - box[0]) // 2 + 20, box[1] + 90, box[2] - 40, box[3] - 40)
    for area, ok in ((left, True), (right, False)):
        lx, ly, rx, by = area
        fill = "#5a94c2" if ok else "#c97a4a"
        load = draw_load(draw, area, fill, "ГРУЗ")
        cx = (load[0] + load[2]) // 2 + (0 if ok else 70)
        cy = (load[1] + load[3]) // 2
        draw.ellipse((cx - 18, cy - 18, cx + 18, cy + 18), fill="#f2f2f2", outline="#333333")
        arrow(draw, (cx, cy), (cx + (0 if ok else 80), cy + (0 if ok else 110)), SAFE if ok else DANGER, 5)
        draw.text((lx + 20, ly + 10), "правильно" if ok else "риск", font=PANEL_TITLE_FONT, fill=SUCCESS if ok else DANGER)


def render_angle_load(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], accent: str) -> None:
    draw_panel_frame(draw, box, "Нагрузка", accent)
    left = (box[0] + 40, box[1] + 90, box[0] + (box[2] - box[0]) // 2 - 20, box[3] - 40)
    right = (box[0] + (box[2] - box[0]) // 2 + 20, box[1] + 90, box[2] - 40, box[3] - 40)
    for area, wide in ((left, False), (right, True)):
        lx, ly, rx, by = area
        cx = (lx + rx) // 2
        hook_y = ly + 40
        draw_hook(draw, cx, ly + 5, hook_y, "#4f5964")
        load = (lx + 80, by - 150, rx - 80, by - 60)
        draw.rounded_rectangle(load, 12, fill="#8aa86d", outline="#5d7049", width=4)
        left_top = (load[0] + 22, load[1])
        right_top = (load[2] - 22, load[1])
        if wide:
            anchor_left = (cx - 160, hook_y)
            anchor_right = (cx + 160, hook_y)
        else:
            anchor_left = (cx - 70, hook_y)
            anchor_right = (cx + 70, hook_y)
        draw.line((*anchor_left, *left_top), fill=accent, width=8)
        draw.line((*anchor_right, *right_top), fill=accent, width=8)
        arrow(draw, (load[0] + 22, load[1] + 10), (load[0] - 50, load[1] - (40 if wide else 10)), DANGER if wide else SAFE, 5)
        arrow(draw, (load[2] - 22, load[1] + 10), (load[2] + 50, load[1] - (40 if wide else 10)), DANGER if wide else SAFE, 5)
        draw.text((lx + 24, ly + 14), "большой угол" if wide else "малый угол", font=SMALL_FONT, fill=DANGER if wide else SUCCESS)


def render_edge_protection(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], accent: str) -> None:
    draw_panel_frame(draw, box, "Контакт", accent)
    left = (box[0] + 42, box[1] + 95, box[0] + (box[2] - box[0]) // 2 - 16, box[3] - 48)
    right = (box[0] + (box[2] - box[0]) // 2 + 16, box[1] + 95, box[2] - 42, box[3] - 48)
    for area, protected in ((left, False), (right, True)):
        lx, ly, rx, by = area
        draw.rounded_rectangle(area, 24, fill=WHITE, outline=GRID, width=2)
        edge = [(lx + 120, by - 80), (rx - 80, by - 80), (rx - 80, ly + 90)]
        draw.line(edge, fill="#666666", width=10)
        sling_y = ly + 60
        draw.line((lx + 60, sling_y, rx - 120, by - 80), fill=accent, width=14)
        if protected:
            draw.rounded_rectangle((rx - 180, by - 122, rx - 72, by - 78), 14, fill=SUCCESS)
            draw.text((lx + 24, ly + 18), "правильно", font=PANEL_TITLE_FONT, fill=SUCCESS)
        else:
            draw.line((rx - 170, by - 140, rx - 50, by - 20), fill=DANGER, width=10)
            draw.line((rx - 50, by - 140, rx - 170, by - 20), fill=DANGER, width=10)
            draw.text((lx + 24, ly + 18), "нельзя", font=PANEL_TITLE_FONT, fill=DANGER)


def render_inspection_points(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], accent: str, bullets: list[str]) -> None:
    draw_panel_frame(draw, box, "Контроль", accent)
    cx = (box[0] + box[2]) // 2
    draw_hook(draw, cx, box[1] + 95, box[1] + 240, "#4f5964")
    load = draw_load(draw, (box[0] + 120, box[1] + 160, box[2] - 120, box[3] - 70), "#7b9bb4")
    points = [
        (cx, box[1] + 170),
        (load[0] + 40, load[1] + 20),
        (load[2] - 40, load[1] + 20),
        (load[0] + 90, load[3] - 20),
    ]
    labels = bullets[:4]
    anchors = [(box[0] + 52, box[1] + 120), (box[2] - 210, box[1] + 120), (box[2] - 220, box[1] + 260), (box[0] + 48, box[3] - 120)]
    for idx, (pt, anchor_pt, text) in enumerate(zip(points, anchors, labels), start=1):
        draw.ellipse((pt[0] - 8, pt[1] - 8, pt[0] + 8, pt[1] + 8), fill=accent)
        arrow(draw, anchor_pt, pt, accent, 4)
        ax, ay = anchor_pt
        draw.rounded_rectangle((ax - 8, ay - 10, ax + 150, ay + 55), 16, fill=WHITE, outline=accent, width=2)
        draw.text((ax + 12, ay), f"{idx}. {text}", font=SMALL_FONT, fill=TEXT)


def render_stop_poster(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], accent: str) -> None:
    draw_panel_frame(draw, box, "Решение", accent)
    cx = (box[0] + box[2]) // 2
    cy = (box[1] + box[3]) // 2 + 20
    stop_octagon(draw, cx, cy)
    draw.text((box[0] + 84, box[1] + 94), "опасную операцию прерывают сразу", font=PANEL_TITLE_FONT, fill=DANGER)


def render_compare(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], accent: str) -> None:
    draw_panel_frame(draw, box, "Сравнение", accent)
    left = (box[0] + 40, box[1] + 90, box[0] + (box[2] - box[0]) // 2 - 20, box[3] - 40)
    right = (box[0] + (box[2] - box[0]) // 2 + 20, box[1] + 90, box[2] - 40, box[3] - 40)
    for area, ok in ((left, False), (right, True)):
        lx, ly, rx, by = area
        draw.rounded_rectangle(area, 24, fill=WHITE, outline=GRID, width=2)
        draw_hook(draw, (lx + rx) // 2, ly + 18, ly + 100, "#4f5964")
        load = draw_load(draw, (lx + 40, ly + 90, rx - 40, by - 30), "#87a968" if ok else "#b38162")
        if ok:
            draw.line((load[0] + 30, load[1], (lx + rx) // 2, ly + 98), fill=SUCCESS, width=10)
            draw.line((load[2] - 30, load[1], (lx + rx) // 2, ly + 98), fill=SUCCESS, width=10)
        else:
            draw.line((load[0] + 30, load[2] - 20, (lx + rx) // 2, ly + 98), fill=DANGER, width=10)
            draw.line((load[2] - 30, load[1] + 100, (lx + rx) // 2, ly + 98), fill=DANGER, width=10)
        draw.text((lx + 22, ly + 20), "правильно" if ok else "ошибка", font=PANEL_TITLE_FONT, fill=SUCCESS if ok else DANGER)


def render_signs(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], accent: str, bullets: list[str]) -> None:
    draw_panel_frame(draw, box, "Знаки", accent)
    x1, y1, x2, _ = box
    card_w = (x2 - x1 - 90) // 3
    card_h = 130
    for idx, label in enumerate(bullets[:6]):
        row, col = divmod(idx, 3)
        cx = x1 + 24 + col * (card_w + 18)
        cy = y1 + 96 + row * (card_h + 18)
        draw.rounded_rectangle((cx, cy, cx + card_w, cy + card_h), 20, fill=WHITE, outline=accent, width=3)
        draw.rectangle((cx + 24, cy + 18, cx + 94, cy + 88), outline=accent, width=4)
        if idx == 0:
            arrow(draw, (cx + 59, cy + 80), (cx + 59, cy + 26), accent, 6)
        elif idx == 1:
            draw.line((cx + 24, cy + 88, cx + 94, cy + 18), fill=accent, width=5)
            draw.line((cx + 24, cy + 18, cx + 94, cy + 88), fill=accent, width=5)
        elif idx == 2:
            draw.arc((cx + 26, cy + 22, cx + 92, cy + 86), 200, 340, fill=accent, width=6)
            draw.line((cx + 26, cy + 60, cx + 92, cy + 60), fill=accent, width=6)
        elif idx == 3:
            draw.ellipse((cx + 36, cy + 30, cx + 84, cy + 78), outline=accent, width=4)
            draw.line((cx + 60, cy + 18, cx + 60, cy + 90), fill=accent, width=4)
        elif idx == 4:
            draw.line((cx + 28, cy + 80, cx + 90, cy + 26), fill=accent, width=6)
            draw.line((cx + 28, cy + 30, cx + 90, cy + 86), fill=accent, width=6)
        else:
            draw.line((cx + 32, cy + 80, cx + 60, cy + 28), fill=accent, width=6)
            draw.line((cx + 88, cy + 80, cx + 60, cy + 28), fill=accent, width=6)
        draw_wrapped(draw, (cx + 18, cy + 96), label, SMALL_FONT, TEXT, card_w - 36, 5)


def render_flow_steps(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], accent: str, bullets: list[str]) -> None:
    draw_panel_frame(draw, box, "Этапы", accent)
    x1, y1, x2, _ = box
    usable = x2 - x1 - 80
    step_w = usable // max(len(bullets), 1)
    cy = y1 + 200
    prev_end = None
    for idx, label in enumerate(bullets):
        sx = x1 + 36 + idx * step_w
        ex = sx + step_w - 24
        draw.rounded_rectangle((sx, cy, ex, cy + 140), 26, fill=WHITE, outline=accent, width=3)
        draw.ellipse((sx + 22, cy + 22, sx + 66, cy + 66), fill=accent)
        num = str(idx + 1)
        nb = draw.textbbox((0, 0), num, font=SMALL_FONT)
        draw.text((sx + 44 - (nb[2] - nb[0]) // 2, cy + 28), num, font=SMALL_FONT, fill=WHITE)
        draw_wrapped(draw, (sx + 24, cy + 82), label, SMALL_FONT, TEXT, step_w - 56, 4)
        if prev_end:
            arrow(draw, prev_end, (sx - 8, cy + 70), accent, 5)
        prev_end = (ex + 8, cy + 70)


def render_mass_case(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], accent: str) -> None:
    draw_panel_frame(draw, box, "Кейс", accent)
    load = draw_load(draw, box, "#9c7c58")
    draw.text((load[0] + 32, load[1] + 30), "масса ?", font=TITLE_FONT, fill=WHITE)
    tag = (load[2] - 260, load[3] - 140, load[2] - 36, load[3] - 58)
    draw.rounded_rectangle(tag, 18, fill="#fff2f2", outline=DANGER, width=3)
    draw.text((tag[0] + 24, tag[1] + 20), "не наверняка", font=PANEL_TITLE_FONT, fill=DANGER)
    draw.line((tag[0] + 8, tag[1] + 8, tag[2] - 8, tag[3] - 8), fill=DANGER, width=6)


def render_edge_detail(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], accent: str) -> None:
    draw_panel_frame(draw, box, "Деталь", accent)
    x1, y1, x2, y2 = box
    cx = (x1 + x2) // 2
    draw.line((x1 + 220, y2 - 150, x2 - 220, y1 + 160), fill=accent, width=20)
    draw.line((cx + 40, y1 + 180, cx + 40, y2 - 180), fill="#666666", width=16)
    draw.rectangle((cx + 28, (y1 + y2) // 2 - 70, cx + 200, (y1 + y2) // 2 + 70), outline=DANGER, width=8)
    draw.text((x1 + 70, y1 + 90), "опасный контакт", font=TITLE_FONT, fill=DANGER)


def render_low_light(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], accent: str) -> None:
    draw_panel_frame(draw, box, "Условия", accent)
    x1, y1, x2, y2 = box
    draw.rounded_rectangle((x1 + 24, y1 + 78, x2 - 24, y2 - 24), 24, fill="#223141")
    draw.ellipse((x1 + 170, y1 + 120, x1 + 300, y1 + 250), fill="#f9d96a")
    draw_load(draw, (x1 + 120, y1 + 200, x2 - 120, y2 - 80), "#50667a")
    draw.text((x1 + 80, y1 + 92), "работу не начинать", font=PANEL_TITLE_FONT, fill="#f9d96a")


def render_signal_roles(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], accent: str) -> None:
    draw_panel_frame(draw, box, "Роли", accent)
    x1, y1, x2, _ = box
    xs = [x1 + 160, (x1 + x2) // 2, x2 - 160]
    colors = [DANGER, SUCCESS, DANGER]
    labels = ["лишний", "назначен", "лишний"]
    for x, color, label in zip(xs, colors, labels):
        draw_person(draw, x, y1 + 150, color)
        draw.text((x - 52, y1 + 316), label, font=SMALL_FONT, fill=color)
    draw.text((x1 + 96, y1 + 92), "команду подает только назначенное лицо", font=PANEL_TITLE_FONT, fill=TEXT)


def render_unclear_signal(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], accent: str) -> None:
    draw_panel_frame(draw, box, "Стоп", accent)
    cx = (box[0] + box[2]) // 2
    stop_octagon(draw, cx, box[1] + 250, 80)
    arrow(draw, (box[0] + 120, box[1] + 170), (cx - 120, box[1] + 210), MUTED, 5)
    arrow(draw, (box[2] - 120, box[1] + 170), (cx + 120, box[1] + 210), MUTED, 5)
    draw.text((box[0] + 98, box[1] + 110), "неясный сигнал = остановка", font=PANEL_TITLE_FONT, fill=TEXT)


def render_tagline_guiding(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], accent: str) -> None:
    draw_panel_frame(draw, box, "Оттяжка", accent)
    load = draw_load(draw, (box[0] + 120, box[1] + 120, box[2] - 200, box[3] - 120), "#6888a7")
    cx = (load[0] + load[2]) // 2
    draw_hook(draw, cx, box[1] + 90, load[1] - 14, "#4f5964")
    rope_start = (load[2] - 10, load[1] + 40)
    rope_end = (box[2] - 150, load[3] - 30)
    draw.line((*rope_start, *rope_end), fill=accent, width=8)
    draw_person(draw, box[2] - 120, load[3] - 130, SUCCESS)
    draw.text((box[0] + 78, box[1] + 96), "человек держит груз на дистанции", font=PANEL_TITLE_FONT, fill=TEXT)


def render_danger_zone(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], accent: str) -> None:
    draw_panel_frame(draw, box, "Опасная зона", accent)
    x1, y1, x2, y2 = box
    load = draw_load(draw, (x1 + 220, y1 + 180, x2 - 280, y2 - 180), "#7d9b6b")
    cx = (load[0] + load[2]) // 2
    draw_hook(draw, cx, y1 + 120, load[1] - 14, "#4f5964")
    draw.ellipse((x1 + 90, y1 + 120, x2 - 90, y2 - 120), outline=DANGER, width=8)
    draw.text((x1 + 110, y1 + 128), "людей в красной зоне быть не должно", font=PANEL_TITLE_FONT, fill=DANGER)
    draw_person(draw, x1 + 160, y2 - 260, DANGER)
    draw_person(draw, x2 - 170, y1 + 210, DANGER)
    draw_person(draw, x2 - 120, y2 - 220, SUCCESS)
    draw.text((x2 - 180, y2 - 68), "безопасно", font=SMALL_FONT, fill=SUCCESS)


def render_stacking(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], accent: str) -> None:
    draw_panel_frame(draw, box, "Складирование", accent)
    x1, y1, x2, y2 = box
    draw.line((x1 + 80, y2 - 100, x2 - 80, y2 - 100), fill="#6d6d6d", width=8)
    for px in (x1 + 280, x2 - 280):
        draw.rectangle((px - 60, y2 - 118, px + 60, y2 - 78), fill="#8b6d46")
    load = (x1 + 220, y2 - 280, x2 - 220, y2 - 120)
    draw.rounded_rectangle(load, 12, fill="#7291b0", outline="#4d6175", width=4)
    draw.text((x1 + 110, y1 + 110), "основание и подкладки готовят заранее", font=PANEL_TITLE_FONT, fill=TEXT)


def render_swing_people(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], accent: str) -> None:
    draw_panel_frame(draw, box, "Кейс", accent)
    x1, y1, x2, y2 = box
    cx = (x1 + x2) // 2
    load = draw_load(draw, (x1 + 260, y1 + 220, x2 - 220, y2 - 160), "#8e9fb0")
    draw_hook(draw, cx, y1 + 90, load[1] - 12, "#4f5964")
    draw.arc((load[0] - 120, load[1] - 100, load[2] + 80, load[3] + 10), 250, 325, fill=DANGER, width=8)
    draw_person(draw, x1 + 150, y2 - 260, DANGER)
    draw_person(draw, x2 - 180, y2 - 220, DANGER)
    stop_octagon(draw, x2 - 170, y1 + 170, 60)


def render_unstable_load(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], accent: str) -> None:
    draw_panel_frame(draw, box, "Запрет", accent)
    x1, y1, x2, y2 = box
    load = Image.new("RGBA", (420, 220), (0, 0, 0, 0))
    ld = ImageDraw.Draw(load)
    ld.rounded_rectangle((0, 0, 420, 220), 16, fill="#9a8a74", outline="#675845", width=4)
    rotated = load.rotate(-12, expand=True)
    draw.bitmap((x1 + 250, y1 + 220), rotated)
    draw.line((x1 + 240, y2 - 110, x2 - 220, y2 - 110), fill="#6b6b6b", width=8)
    draw.line((x1 + 220, y1 + 160, x2 - 200, y2 - 40), fill=DANGER, width=10)
    draw.line((x2 - 200, y1 + 160, x1 + 220, y2 - 40), fill=DANGER, width=10)


def render_obstacles_case(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], accent: str) -> None:
    draw_panel_frame(draw, box, "Обстановка", accent)
    x1, y1, x2, y2 = box
    load = draw_load(draw, (x1 + 220, y1 + 180, x2 - 240, y2 - 160), "#8b9e74")
    draw_hook(draw, (load[0] + load[2]) // 2, y1 + 90, load[1] - 10, "#4f5964")
    arrow(draw, (load[2] + 30, load[1] + 30), (x2 - 120, y1 + 200), DANGER, 6)
    draw.rectangle((x1 + 90, y1 + 250, x1 + 180, y1 + 360), outline="#7b7b7b", width=5)
    draw_person(draw, x2 - 150, y2 - 260, DANGER)
    draw.text((x1 + 90, y1 + 112), "нужно смотреть не только на груз", font=PANEL_TITLE_FONT, fill=TEXT)


def render_stop_card(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], accent: str, bullets: list[str]) -> None:
    draw_panel_frame(draw, box, "Решение", accent)
    stop_octagon(draw, box[0] + 230, box[1] + 260, 82)
    y = box[1] + 150
    for bullet in bullets:
        draw.rounded_rectangle((box[0] + 360, y, box[2] - 80, y + 72), 16, fill="#fff7f4", outline=accent, width=2)
        draw.text((box[0] + 390, y + 20), bullet, font=BODY_FONT, fill=TEXT)
        y += 90


def render_lep_special(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], accent: str) -> None:
    draw_panel_frame(draw, box, "ЛЭП", accent)
    x1, y1, x2, y2 = box
    for px in (x1 + 180, x2 - 180):
        draw.line((px, y1 + 120, px, y2 - 140), fill="#5b5b5b", width=10)
    for offset in (0, 44, 88):
        draw.line((x1 + 180, y1 + 140 + offset, x2 - 180, y1 + 120 + offset), fill="#444444", width=4)
    crane_base = (x1 + 220, y2 - 160, x1 + 420, y2 - 80)
    draw.rounded_rectangle(crane_base, 12, fill="#d0a04f", outline="#885b10", width=4)
    draw.line((crane_base[0] + 40, crane_base[1], crane_base[0] + 220, y1 + 280), fill=accent, width=16)
    draw.ellipse((x1 + 120, y1 + 180, x2 - 120, y2 - 180), outline=DANGER, width=7)
    draw.text((x1 + 90, y1 + 86), "расстояние на глаз не оценивают", font=PANEL_TITLE_FONT, fill=DANGER)


def render_lep_stop(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], accent: str) -> None:
    render_lep_special(draw, box, accent)
    stop_octagon(draw, box[2] - 170, box[3] - 150, 64)
    draw.text((box[0] + 94, box[3] - 120), "при неясных границах работу останавливают", font=PANEL_TITLE_FONT, fill=TEXT)


def render_card_grid(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], accent: str, bullets: list[str]) -> None:
    render_checklist(draw, box, accent, bullets[:4])


def render_duties_checklist(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], accent: str, bullets: list[str]) -> None:
    draw_panel_frame(draw, box, "Обязанности", accent)
    x1, y1, x2, _ = box
    y = y1 + 96
    for idx, bullet in enumerate(bullets, start=1):
        draw.rounded_rectangle((x1 + 28, y, x2 - 28, y + 92), 20, fill=WHITE, outline=accent, width=2)
        draw.ellipse((x1 + 48, y + 24, x1 + 92, y + 68), fill=accent)
        draw.text((x1 + 62, y + 30), str(idx), font=SMALL_FONT, fill=WHITE)
        draw_wrapped(draw, (x1 + 118, y + 22), bullet, BODY_FONT, TEXT, x2 - x1 - 168)
        y += 106


def render_admission_flow(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], accent: str, bullets: list[str]) -> None:
    render_flow_steps(draw, box, accent, bullets[:4])


def render_stop_message(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], accent: str, bullets: list[str]) -> None:
    draw_panel_frame(draw, box, "Алгоритм", accent)
    y = box[1] + 110
    for idx, bullet in enumerate(bullets, start=1):
        draw.rounded_rectangle((box[0] + 90, y, box[2] - 90, y + 96), 18, fill=WHITE, outline=accent, width=2)
        draw.text((box[0] + 120, y + 28), f"{idx}.", font=PANEL_TITLE_FONT, fill=accent)
        draw_wrapped(draw, (box[0] + 180, y + 22), bullet, BODY_FONT, TEXT, box[2] - box[0] - 300)
        y += 118


def render_defect_people_case(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], accent: str) -> None:
    draw_panel_frame(draw, box, "Критично", accent)
    x1, y1, x2, y2 = box
    load = draw_load(draw, (x1 + 220, y1 + 180, x2 - 260, y2 - 160), "#899b73")
    draw_hook(draw, (load[0] + load[2]) // 2, y1 + 90, load[1] - 14, "#4f5964")
    draw.line((load[0] + 16, load[1] + 12, load[0] - 80, load[1] - 80), fill=DANGER, width=10)
    draw.line((load[0] - 80, load[1] + 12, load[0] + 16, load[1] - 80), fill=DANGER, width=10)
    draw_person(draw, x2 - 160, y2 - 250, DANGER)
    stop_octagon(draw, x2 - 180, y1 + 170, 60)


KIND_MAP = {
    "checklist": render_checklist,
    "warning_load": render_warning_load,
    "center_gravity": render_center_gravity,
    "angle_load": render_angle_load,
    "edge_protection": render_edge_protection,
    "inspection_points": render_inspection_points,
    "stop_poster": render_stop_poster,
    "compare": render_compare,
    "signs": render_signs,
    "flow_steps": render_flow_steps,
    "mass_case": render_mass_case,
    "edge_detail": render_edge_detail,
    "low_light": render_low_light,
    "signal_roles": render_signal_roles,
    "unclear_signal": render_unclear_signal,
    "tagline_guiding": render_tagline_guiding,
    "danger_zone": render_danger_zone,
    "stacking": render_stacking,
    "swing_people": render_swing_people,
    "unstable_load": render_unstable_load,
    "obstacles_case": render_obstacles_case,
    "stop_card": render_stop_card,
    "lep_special": render_lep_special,
    "lep_stop": render_lep_stop,
    "card_grid": render_card_grid,
    "duties_checklist": render_duties_checklist,
    "admission_flow": render_admission_flow,
    "stop_message": render_stop_message,
    "defect_people_case": render_defect_people_case,
}


def render(spec: SlideSpec) -> Path:
    image, draw = create_canvas(spec.accent)
    header(draw, spec)
    left, right = panels()
    renderer = KIND_MAP[spec.kind]
    if spec.kind in {"checklist", "signs", "flow_steps", "stop_card", "card_grid", "duties_checklist", "admission_flow", "stop_message"}:
        renderer(draw, left, spec.accent, spec.bullets)
    elif spec.kind == "inspection_points":
        renderer(draw, left, spec.accent, spec.bullets)
    else:
        renderer(draw, left, spec.accent)
    draw_bullets(draw, right, "Что запомнить", spec.bullets, spec.accent)
    out_dir = ASSETS / spec.module_slug / "diagrams" / "error-analysis"
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / spec.filename
    image.save(out_path, format="PNG")
    return out_path


def ensure_module_dirs() -> None:
    for slug, title in [
        ("module-03-proizvodstvo-rabot", "Модуль 3. Производство работ"),
        ("module-04-okhrana-truda", "Модуль 4. Охрана труда"),
    ]:
        base = ASSETS / slug
        for sub in ("diagrams", "images", "references", "video"):
            (base / sub).mkdir(parents=True, exist_ok=True)
        readme = base / "README.md"
        if not readme.exists():
            readme.write_text(
                "\n".join(
                    [
                        f"# Media Pack: {title}",
                        "",
                        f"Сюда складываем визуальные материалы для `{title}`.",
                        "",
                        "## Папки",
                        "",
                        "- `images` - фото и ситуационные материалы.",
                        "- `diagrams` - схемы, алгоритмы, карточки и учебные кейсы.",
                        "- `video` - ролики и короткие фрагменты, если появятся.",
                        "- `references` - ссылки, исходники и служебные описания.",
                        "",
                    ]
                ),
                encoding="utf-8",
            )


def main() -> None:
    ensure_module_dirs()
    outputs = [render(spec) for spec in SPECS]
    print(f"generated: {len(outputs)}")
    for path in outputs:
        print(path.relative_to(ROOT))


if __name__ == "__main__":
    main()
