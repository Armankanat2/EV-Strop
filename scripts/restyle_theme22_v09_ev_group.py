from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt


SRC_PATH = Path(
    "presentations/ispring-course/module-01-stropovka-gruzov/live-preview/"
    "S029_theme22_preview_2026-06-29_v09.pptx"
)
OUT_PATH = Path(
    "presentations/ispring-course/module-01-stropovka-gruzov/live-preview/"
    "S029_theme22_preview_2026-06-30_v10_ev-group-restyled.pptx"
)

PNG_OUT_DIR = Path(
    "presentations/ispring-course/module-01-stropovka-gruzov/live-preview/"
    "S029_theme22_preview_2026-06-30_v10_ev-group-restyled_png"
)


BLUE = RGBColor(0x00, 0x57, 0xFF)
DEEP_BLUE = RGBColor(0x00, 0x38, 0xB8)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GRAPHITE = RGBColor(0x11, 0x11, 0x11)
DARK_GRAY = RGBColor(0x4B, 0x55, 0x63)
LIGHT_GRAY = RGBColor(0xE5, 0xE7, 0xEB)
PALE_BLUE = RGBColor(0xEA, 0xF0, 0xFF)
RED = RGBColor(0xC5, 0x1F, 0x1F)


ASSETS = {
    "S029": Path("assets/course-media/module-01-stropovka-gruzov/images/S004_avatar-1_signal-podnyat-gruz_white-bg.png"),
    "S031": Path("assets/course-media/module-02-tekhnologiya-raboty-stropalshchika/diagrams/error-analysis/S031-P01_podgotovka-k-rabote_checklist.png"),
    "S032": Path("assets/course-media/module-02-tekhnologiya-raboty-stropalshchika/diagrams/error-analysis/S031-P02_plohaia-osveshchennost_warning.png"),
    "S035": Path("assets/course-media/module-01-stropovka-gruzov/images/S004_avatar-1_signal-podnyat-gruz_white-bg.png"),
    "S036": Path("assets/course-media/module-02-tekhnologiya-raboty-stropalshchika/diagrams/error-analysis/S031-P05_stop-prioritet_poster.png"),
    "S037": Path("assets/course-media/module-02-tekhnologiya-raboty-stropalshchika/diagrams/error-analysis/S031-P08_opasnaia-zona_movement-diagram.png"),
    "S038": Path("assets/course-media/module-02-tekhnologiya-raboty-stropalshchika/diagrams/error-analysis/S031-P11_raskachivanie-liudi_case-diagram.png"),
    "S039": Path("assets/course-media/module-02-tekhnologiya-raboty-stropalshchika/diagrams/error-analysis/S031-P09_skladovanie_podkladki-diagram.png"),
    "S031-P01": Path("assets/course-media/module-02-tekhnologiya-raboty-stropalshchika/diagrams/error-analysis/S031-P01_podgotovka-k-rabote_checklist.png"),
    "S034-P01": Path("assets/course-media/module-02-tekhnologiya-raboty-stropalshchika/diagrams/error-analysis/S031-P03_obviazka-pered-podem_steps.png"),
    "S035-P01": Path("assets/course-media/module-01-stropovka-gruzov/images/S004_avatar-1_signal-palm-down-source.png"),
    "S035-PP01": Path("assets/course-media/module-01-stropovka-gruzov/images/S004_avatar-1_signal-podnyat-gruz_white-bg.png"),
    "S037-P01": Path("assets/course-media/module-02-tekhnologiya-raboty-stropalshchika/diagrams/error-analysis/S031-P08_opasnaia-zona_movement-diagram.png"),
    "S037-PP01": Path("assets/course-media/module-02-tekhnologiya-raboty-stropalshchika/diagrams/error-analysis/S031-P07_ottiagka_safe-guiding-diagram.png"),
    "S039-P01": Path("assets/course-media/module-02-tekhnologiya-raboty-stropalshchika/diagrams/error-analysis/S031-P10_rasstropovka_rano-pravilno.png"),
    "S039-PP01": Path("assets/course-media/module-02-tekhnologiya-raboty-stropalshchika/diagrams/error-analysis/S031-P09_skladovanie_podkladki-diagram.png"),
}


CAPTIONS = {
    "S029": "Стропальщик подает команду перед началом грузоподъемной операции",
    "S031": "Памятка по подготовке до начала работ",
    "S032": "Пример условия, при котором работу начинать нельзя",
    "S035": "Ключевой ручной сигнал для взаимодействия с крановщиком",
    "S036": "Команда «Стоп» имеет безусловный приоритет",
    "S037": "Опасная зона и безопасная позиция при перемещении груза",
    "S038": "Нештатная ситуация: перемещение нужно остановить",
    "S039": "Подкладки и устойчивое положение перед расстроповкой",
    "S031-P01": "Подготовка: задание, условия, оснастка и рабочая зона",
    "S034-P01": "Раскрытие маршрута операции по этапам",
    "S035-P01": "Способы подачи сигналов и пример ручной команды",
    "S035-PP01": "Галерея жестов должна читаться как единая система",
    "S037-P01": "Схема опасной зоны для отдельного разбора",
    "S037-PP01": "Работа с оттяжкой только на безопасной дистанции",
    "S039-P01": "Расстроповка допустима только после устойчивой установки",
    "S039-PP01": "Варианты безопасного складирования",
}


def set_fill(shape, rgb, transparency=None):
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb
    if transparency is not None:
        shape.fill.transparency = transparency


def set_line(shape, rgb, width_pt=1.0):
    shape.line.color.rgb = rgb
    shape.line.width = Pt(width_pt)


def style_text_shape(shape, size, color, bold=False, align=PP_ALIGN.LEFT, wrap=True):
    tf = shape.text_frame
    tf.word_wrap = wrap
    tf.margin_left = Pt(8)
    tf.margin_right = Pt(8)
    tf.margin_top = Pt(6)
    tf.margin_bottom = Pt(6)
    for p in tf.paragraphs:
        p.alignment = align
        for run in p.runs:
            run.font.name = "Arial"
            run.font.size = Pt(size)
            run.font.bold = bold
            run.font.color.rgb = color


def contain_image(slide, image_path, left, top, width, height):
    with Image.open(image_path) as img:
        img_w, img_h = img.size

    box_ratio = width / height
    img_ratio = img_w / img_h
    if img_ratio > box_ratio:
        pic_w = width
        pic_h = width / img_ratio
    else:
        pic_h = height
        pic_w = height * img_ratio

    pic_left = left + (width - pic_w) / 2
    pic_top = top + (height - pic_h) / 2
    return slide.shapes.add_picture(str(image_path), pic_left, pic_top, width=pic_w, height=pic_h)


def add_small_accent(slide):
    accent = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.RECTANGLE,
        Inches(0.56),
        Inches(0.74),
        Inches(1.12),
        Inches(0.05),
    )
    set_fill(accent, BLUE)
    accent.line.fill.background()


def restyle_slide(slide):
    code = slide.shapes[2].text.strip()

    # Background band
    set_fill(slide.shapes[0], WHITE)
    slide.shapes[0].line.fill.background()
    add_small_accent(slide)

    # Titles
    slide.shapes[1].text = "Тема 2.2\nТехнология работы"
    style_text_shape(slide.shapes[1], 10, BLUE, bold=True, wrap=False)
    style_text_shape(slide.shapes[3], 27 if len(slide.shapes[3].text) < 34 else 24, GRAPHITE, bold=True)
    slide.shapes[1].top = Inches(0.30)
    slide.shapes[1].left = Inches(0.64)
    slide.shapes[1].width = Inches(1.45)
    slide.shapes[1].height = Inches(0.42)
    slide.shapes[3].top = Inches(0.93)
    slide.shapes[3].left = Inches(0.64)
    slide.shapes[3].width = Inches(6.2)
    slide.shapes[3].height = Inches(0.8)

    # Code badge
    set_fill(slide.shapes[2], DEEP_BLUE)
    slide.shapes[2].line.fill.background()
    badge_width = 1.15
    badge_font = 13
    if len(code) >= 8:
        badge_width = 1.32
        badge_font = 12
    if len(code) >= 9:
        badge_width = 1.48
        badge_font = 11
    style_text_shape(slide.shapes[2], badge_font, WHITE, bold=True, align=PP_ALIGN.CENTER, wrap=False)
    slide.shapes[2].left = Inches(12.20) - Inches(badge_width)
    slide.shapes[2].top = Inches(0.34)
    slide.shapes[2].width = Inches(badge_width)
    slide.shapes[2].height = Inches(0.34)

    # Panels
    for panel_idx in [4, 7]:
        set_fill(slide.shapes[panel_idx], WHITE)
        set_line(slide.shapes[panel_idx], LIGHT_GRAY, width_pt=1.4)
    for header_idx in [5]:
        set_fill(slide.shapes[header_idx], BLUE)
        slide.shapes[header_idx].line.fill.background()
        style_text_shape(slide.shapes[header_idx], 14, WHITE, bold=True)

    # Body text
    left_size = 16
    right_size = 14
    if "-P01" in code or "-PP01" in code:
        left_size = 13
        right_size = 13
    if code in {"S031-P01", "S034-P01", "S035-P01", "S035-PP01", "S039-P01", "S039-PP01"}:
        left_size = 12
        right_size = 12

    style_text_shape(slide.shapes[6], left_size, BLUE)
    style_text_shape(slide.shapes[8], right_size, DARK_GRAY)

    # Subtitle and content blocks spacing
    slide.shapes[4].top = Inches(2.0)
    slide.shapes[4].left = Inches(0.74)
    slide.shapes[4].width = Inches(5.62)
    slide.shapes[4].height = Inches(4.28)
    slide.shapes[5].top = Inches(2.0)
    slide.shapes[5].left = Inches(0.74)
    slide.shapes[5].width = Inches(5.62)
    slide.shapes[5].height = Inches(0.45)
    slide.shapes[6].top = Inches(2.52)
    slide.shapes[6].left = Inches(0.98)
    slide.shapes[6].width = Inches(5.08)
    slide.shapes[6].height = Inches(3.38)

    slide.shapes[7].top = Inches(2.0)
    slide.shapes[7].left = Inches(6.62)
    slide.shapes[7].width = Inches(5.62)
    slide.shapes[7].height = Inches(4.28)

    # Right panel: text-only vs image-driven
    image_path = ASSETS.get(code)
    if image_path and image_path.exists():
        slide.shapes[8].top = Inches(5.94)
        slide.shapes[8].left = Inches(6.86)
        slide.shapes[8].width = Inches(4.95)
        slide.shapes[8].height = Inches(0.28)
        style_text_shape(slide.shapes[8], 10, DARK_GRAY)
        slide.shapes[8].text = CAPTIONS.get(code, "")
        contain_image(slide, image_path, Inches(6.9), Inches(2.18), Inches(5.02), Inches(3.58))
    else:
        slide.shapes[8].top = Inches(2.4)
        slide.shapes[8].left = Inches(6.9)
        slide.shapes[8].width = Inches(4.85)
        slide.shapes[8].height = Inches(3.55)

    # Navigation
    set_fill(slide.shapes[9], WHITE)
    set_line(slide.shapes[9], LIGHT_GRAY, width_pt=1.0)
    style_text_shape(slide.shapes[9], 12, DARK_GRAY, bold=True, align=PP_ALIGN.CENTER)
    slide.shapes[9].left = Inches(0.74)
    slide.shapes[9].top = Inches(6.55)
    slide.shapes[9].width = Inches(1.85)
    slide.shapes[9].height = Inches(0.38)

    set_fill(slide.shapes[10], BLUE)
    slide.shapes[10].line.fill.background()
    style_text_shape(slide.shapes[10], 12, WHITE, bold=True, align=PP_ALIGN.CENTER)
    slide.shapes[10].left = Inches(10.15)
    slide.shapes[10].top = Inches(6.55)
    slide.shapes[10].width = Inches(2.08)
    slide.shapes[10].height = Inches(0.38)


def main():
    prs = Presentation(str(SRC_PATH))
    for slide in prs.slides:
        restyle_slide(slide)
    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    prs.save(OUT_PATH)
    print(OUT_PATH)


if __name__ == "__main__":
    main()
