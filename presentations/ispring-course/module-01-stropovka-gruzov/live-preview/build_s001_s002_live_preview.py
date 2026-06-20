from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt


OUT_DIR = Path(__file__).resolve().parent
OUT_FILE = OUT_DIR / "S001-S002_live_preview_working.pptx"


def u(text):
    return text.encode("ascii").decode("unicode_escape")


NAVY = RGBColor(18, 43, 66)
BLUE = RGBColor(38, 90, 137)
GREEN = RGBColor(66, 133, 95)
STEEL = RGBColor(83, 119, 149)
PALE = RGBColor(240, 244, 247)
ORANGE = RGBColor(222, 116, 54)
WHITE = RGBColor(255, 255, 255)
TEXT = RGBColor(33, 43, 54)
GRAY = RGBColor(110, 122, 134)


prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)


def add_bg(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_text(slide, x, y, w, h, text, size=18, bold=False, color=TEXT, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    return box


def add_band(slide, title, code, subtitle=""):
    band = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, 0, 0, prs.slide_width, Inches(0.72))
    band.fill.solid()
    band.fill.fore_color.rgb = NAVY
    band.line.fill.background()

    add_text(slide, Inches(0.45), Inches(0.14), Inches(8.6), Inches(0.3), title, 24, True, WHITE)

    tag = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Inches(11.45), Inches(0.14), Inches(1.35), Inches(0.35)
    )
    tag.fill.solid()
    tag.fill.fore_color.rgb = ORANGE
    tag.line.fill.background()
    tf = tag.text_frame
    tf.clear()
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = code
    r.font.size = Pt(16)
    r.font.bold = True
    r.font.color.rgb = WHITE

    if subtitle:
        add_text(slide, Inches(0.5), Inches(0.82), Inches(6.5), Inches(0.25), subtitle, 11, False, GRAY)


def add_button(slide, x, y, w, h, text, fill):
    shape = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, x, y, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.color.rgb = fill
    tf = shape.text_frame
    tf.clear()
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = text
    r.font.size = Pt(18)
    r.font.bold = True
    r.font.color.rgb = WHITE
    return shape


def add_panel(slide, x, y, w, h, title, lines, accent):
    panel = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, x, y, w, h)
    panel.fill.solid()
    panel.fill.fore_color.rgb = WHITE
    panel.line.color.rgb = accent
    panel.line.width = Pt(1.5)

    head = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, x, y, w, Inches(0.72))
    head.fill.solid()
    head.fill.fore_color.rgb = accent
    head.line.color.rgb = accent
    tf = head.text_frame
    tf.clear()
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = title
    r.font.size = Pt(22)
    r.font.bold = True
    r.font.color.rgb = WHITE

    body = slide.shapes.add_textbox(x + Inches(0.28), y + Inches(1.0), w - Inches(0.56), h - Inches(1.2))
    tf = body.text_frame
    tf.clear()
    tf.word_wrap = True
    for idx, line in enumerate(lines):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.text = line
        p.level = 0
        p.bullet = True
        p.space_after = Pt(8)
        if p.runs:
            p.runs[0].font.size = Pt(17)
            p.runs[0].font.color.rgb = TEXT
    return panel


# S001
s1 = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s1, PALE)
add_band(s1, u(r"\u0422\u0438\u0442\u0443\u043b\u044c\u043d\u044b\u0439 \u044d\u043a\u0440\u0430\u043d \u043a\u0443\u0440\u0441\u0430"), "S001")
add_text(s1, Inches(0.65), Inches(1.15), Inches(6.5), Inches(0.55), u(r"\u042d\u0412 \u0421\u0442\u0440\u043e\u043f\u0430\u043b\u044c\u0449\u0438\u043a"), 30, True)
add_text(
    s1,
    Inches(0.67),
    Inches(1.95),
    Inches(4.5),
    Inches(0.3),
    u(r"\u0421\u0442\u0440\u043e\u043f\u0430\u043b\u044c\u0449\u0438\u043a 2-4 \u0440\u0430\u0437\u0440\u044f\u0434\u0430"),
    18,
    True,
    BLUE,
)
add_text(
    s1,
    Inches(0.67),
    Inches(2.28),
    Inches(4.5),
    Inches(0.3),
    u(r"\u0421\u0442\u0440\u043e\u043f\u0430\u043b\u044c\u0449\u0438\u043a 5-6 \u0440\u0430\u0437\u0440\u044f\u0434\u0430"),
    18,
    True,
    ORANGE,
)
add_text(
    s1,
    Inches(0.67),
    Inches(2.85),
    Inches(6.2),
    Inches(0.42),
    u(r"\u041f\u0440\u0430\u043a\u0442\u0438\u0447\u0435\u0441\u043a\u0438\u0439 \u043a\u0443\u0440\u0441 \u043f\u043e \u0431\u0435\u0437\u043e\u043f\u0430\u0441\u043d\u043e\u0439 \u0440\u0430\u0431\u043e\u0442\u0435 \u0441 \u0433\u0440\u0443\u0437\u0430\u043c\u0438"),
    20,
    True,
)
note = s1.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Inches(0.67), Inches(3.52), Inches(6.2), Inches(1.08))
note.fill.solid()
note.fill.fore_color.rgb = WHITE
note.line.color.rgb = RGBColor(210, 218, 226)
tf = note.text_frame
tf.clear()
tf.vertical_anchor = MSO_ANCHOR.MIDDLE
p = tf.paragraphs[0]
r = p.add_run()
r.text = u(r"\u0411\u0430\u0437\u043e\u0432\u0430\u044f \u0447\u0430\u0441\u0442\u044c \u043a\u0443\u0440\u0441\u0430 - \u0434\u043b\u044f 2-4 \u0440\u0430\u0437\u0440\u044f\u0434\u0430. \u0414\u043b\u044f 5-6 \u0440\u0430\u0437\u0440\u044f\u0434\u0430 \u043e\u0442\u043a\u0440\u044b\u0432\u0430\u0439\u0442\u0435 \u0434\u043e\u043f\u043e\u043b\u043d\u0438\u0442\u0435\u043b\u044c\u043d\u044b\u0435 \u0431\u043b\u043e\u043a\u0438.")
r.font.size = Pt(17)
r.font.color.rgb = TEXT

right = s1.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(7.9), Inches(0.72), Inches(5.43), Inches(6.78))
right.fill.solid()
right.fill.fore_color.rgb = STEEL
right.line.fill.background()
add_text(s1, Inches(8.35), Inches(1.2), Inches(4.3), Inches(0.3), u(r"\u0416\u0438\u0432\u0430\u044f demo-\u0441\u0431\u043e\u0440\u043a\u0430"), 18, True, WHITE, PP_ALIGN.CENTER)
add_text(
    s1,
    Inches(8.3),
    Inches(1.8),
    Inches(4.45),
    Inches(2.7),
    u(
        r"\u0421\u0435\u0439\u0447\u0430\u0441 \u043c\u043e\u0436\u043d\u043e \u043f\u043e\u0441\u043c\u043e\u0442\u0440\u0435\u0442\u044c:\n\n\u2022 S001 \u0442\u0438\u0442\u0443\u043b\n\u2022 S002 \u044d\u043a\u0440\u0430\u043d \u0432\u044b\u0431\u043e\u0440\u0430\n\u2022 S002-P01 \u0432\u0438\u0434\u0435\u043e\u0438\u043d\u0441\u0442\u0440\u0443\u043a\u0446\u0438\u044f\n\u2022 S002-P02 \u0441\u043b\u043e\u0432\u0430\u0440\u044c \u0442\u0435\u0440\u043c\u0438\u043d\u043e\u0432"
    ),
    18,
    False,
    WHITE,
)
start_btn = add_button(s1, Inches(0.67), Inches(5.15), Inches(2.6), Inches(0.62), u(r"\u041d\u0410\u0427\u0410\u0422\u042c \u041a\u0423\u0420\u0421"), ORANGE)


# S002
s2 = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s2, PALE)
add_band(
    s2,
    u(r"\u0412\u0438\u0434\u0435\u043e\u0438\u043d\u0441\u0442\u0440\u0443\u043a\u0446\u0438\u044f \u043f\u043e \u043a\u0443\u0440\u0441\u0443 \u0438 \u0433\u043b\u043e\u0441\u0430\u0440\u0438\u0439"),
    "S002",
    u(r"\u0422\u0435\u043c\u0430 1 \u2022 \u0432\u0432\u043e\u0434\u043d\u044b\u0439 \u044d\u043a\u0440\u0430\u043d"),
)
add_text(s2, Inches(0.65), Inches(1.0), Inches(4.5), Inches(0.3), u(r"\u0412\u044b\u0431\u0435\u0440\u0438\u0442\u0435, \u0447\u0442\u043e \u043e\u0442\u043a\u0440\u044b\u0442\u044c:"), 20, True)
card_video = add_panel(
    s2,
    Inches(0.7),
    Inches(1.65),
    Inches(5.7),
    Inches(3.9),
    u(r"\u0412\u0438\u0434\u0435\u043e\u0438\u043d\u0441\u0442\u0440\u0443\u043a\u0446\u0438\u044f"),
    [
        u(r"\u042d\u0442\u043e \u043d\u0430\u0432\u0438\u0433\u0430\u0446\u0438\u044f \u043f\u043e \u043a\u0443\u0440\u0441\u0443."),
        u(r"\u041e\u0442\u043a\u0440\u043e\u0439\u0442\u0435 \u0438 \u043e\u0437\u043d\u0430\u043a\u043e\u043c\u044c\u0442\u0435\u0441\u044c."),
        u(r"\u041f\u0435\u0440\u0435\u0445\u043e\u0434 \u0432 \u043f\u043e\u0434\u0432\u0430\u043b S002-P01."),
    ],
    BLUE,
)
card_gloss = add_panel(
    s2,
    Inches(6.9),
    Inches(1.65),
    Inches(5.7),
    Inches(3.9),
    u(r"\u0421\u043b\u043e\u0432\u0430\u0440\u044c \u0442\u0435\u0440\u043c\u0438\u043d\u043e\u0432"),
    [
        u(r"\u041e\u0442\u043a\u0440\u043e\u0439\u0442\u0435 \u0438 \u043f\u0440\u043e\u0447\u0442\u0438\u0442\u0435."),
        u(r"\u0422\u0435\u0440\u043c\u0438\u043d\u044b \u0432\u044b\u043d\u0435\u0441\u0435\u043d\u044b \u0432 \u043e\u0442\u0434\u0435\u043b\u044c\u043d\u044b\u0439 \u043f\u043e\u0434\u0432\u0430\u043b."),
        u(r"\u041f\u0435\u0440\u0435\u0445\u043e\u0434 \u0432 \u043f\u043e\u0434\u0432\u0430\u043b S002-P02."),
    ],
    GREEN,
)
btn_video = add_button(s2, Inches(1.75), Inches(5.0), Inches(3.55), Inches(0.62), u(r"\u041e\u0422\u041a\u0420\u042b\u0422\u042c S002-P01"), BLUE)
btn_gloss = add_button(s2, Inches(7.95), Inches(5.0), Inches(3.55), Inches(0.62), u(r"\u041e\u0422\u041a\u0420\u042b\u0422\u042c S002-P02"), GREEN)


# S002-P01
s3 = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s3, WHITE)
add_band(
    s3,
    u(r"\u0412\u0438\u0434\u0435\u043e\u0438\u043d\u0441\u0442\u0440\u0443\u043a\u0446\u0438\u044f \u043f\u043e \u043a\u0443\u0440\u0441\u0443"),
    "S002-P01",
    u(r"\u041f\u043e\u0434\u0432\u0430\u043b \u043e\u0442 S002"),
)
preview = s3.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Inches(0.7), Inches(1.25), Inches(4.8), Inches(3.85))
preview.fill.solid()
preview.fill.fore_color.rgb = RGBColor(224, 232, 239)
preview.line.color.rgb = STEEL
add_text(s3, Inches(1.05), Inches(1.58), Inches(4.1), Inches(0.45), u(r"\u0417\u0434\u0435\u0441\u044c \u0431\u0443\u0434\u0435\u0442 \u0432\u0438\u0434\u0435\u043e\u0440\u043e\u043b\u0438\u043a"), 24, True, NAVY, PP_ALIGN.CENTER)
add_text(
    s3,
    Inches(0.95),
    Inches(2.35),
    Inches(4.25),
    Inches(1.1),
    u(r"\u0420\u0430\u0431\u043e\u0447\u0430\u044f \u043e\u0437\u0432\u0443\u0447\u043a\u0430 \u0443\u0436\u0435 \u043f\u043e\u0434\u0433\u043e\u0442\u043e\u0432\u043b\u0435\u043d\u0430:\nS002-P01_ozvuchka_osnovnaya.mp3"),
    17,
    False,
    TEXT,
    PP_ALIGN.CENTER,
)
add_text(s3, Inches(5.95), Inches(1.3), Inches(5.9), Inches(0.3), u(r"\u0427\u0442\u043e \u0434\u043e\u043b\u0436\u0435\u043d \u043f\u043e\u043d\u044f\u0442\u044c \u0441\u043b\u0443\u0448\u0430\u0442\u0435\u043b\u044c"), 20, True)
add_panel(
    s3,
    Inches(5.9),
    Inches(1.8),
    Inches(6.0),
    Inches(3.55),
    u(r"\u041a\u043e\u0440\u043e\u0442\u043a\u043e \u043e \u043b\u043e\u0433\u0438\u043a\u0435 \u043a\u0443\u0440\u0441\u0430"),
    [
        u(r"\u041a\u0443\u0440\u0441 \u0438\u043d\u0442\u0435\u0440\u0430\u043a\u0442\u0438\u0432\u043d\u044b\u0439."),
        u(r"\u0422\u0435\u043c\u044b \u0438\u0434\u0443\u0442 \u043f\u043e \u043f\u043e\u0440\u044f\u0434\u043a\u0443."),
        u(r"\u041f\u043e\u0441\u043b\u0435 \u0442\u0435\u043c \u0438 \u0431\u043b\u043e\u043a\u043e\u0432 \u0431\u0443\u0434\u0443\u0442 \u0442\u0435\u0441\u0442\u044b."),
        u(r"\u041f\u043e \u043a\u043d\u043e\u043f\u043a\u0430\u043c \u043e\u0442\u043a\u0440\u044b\u0432\u0430\u0439\u0442\u0435 \u043f\u043e\u044f\u0441\u043d\u0435\u043d\u0438\u044f."),
        u(r"\u041f\u043e\u0441\u043b\u0435 \u043f\u0440\u043e\u0441\u043c\u043e\u0442\u0440\u0430 \u0432\u0435\u0440\u043d\u0438\u0442\u0435\u0441\u044c \u0432 \u043a\u0443\u0440\u0441."),
    ],
    BLUE,
)
back1 = add_button(s3, Inches(9.15), Inches(6.55), Inches(3.05), Inches(0.5), u(r"\u041d\u0410\u0417\u0410\u0414 \u041a \u041a\u0423\u0420\u0421\u0423"), NAVY)


# S002-P02
s4 = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s4, WHITE)
add_band(
    s4,
    u(r"\u0421\u043b\u043e\u0432\u0430\u0440\u044c \u0442\u0435\u0440\u043c\u0438\u043d\u043e\u0432"),
    "S002-P02",
    u(r"\u041f\u043e\u0434\u0432\u0430\u043b \u043e\u0442 S002"),
)
add_panel(
    s4,
    Inches(0.7),
    Inches(1.4),
    Inches(5.8),
    Inches(4.9),
    u(r"\u041b\u0435\u0432\u0430\u044f \u043a\u043e\u043b\u043e\u043d\u043a\u0430"),
    [
        u(r"\u0421\u0442\u0440\u043e\u043f\u0430\u043b\u044c\u0449\u0438\u043a - \u0440\u0430\u0431\u043e\u0447\u0438\u0439, \u043a\u043e\u0442\u043e\u0440\u044b\u0439 \u0432\u044b\u043f\u043e\u043b\u043d\u044f\u0435\u0442 \u0441\u0442\u0440\u043e\u043f\u043e\u0432\u043a\u0443 \u0433\u0440\u0443\u0437\u0430."),
        u(r"\u0421\u0442\u0440\u043e\u043f - \u0433\u0440\u0443\u0437\u043e\u0437\u0430\u0445\u0432\u0430\u0442\u043d\u043e\u0435 \u0441\u0440\u0435\u0434\u0441\u0442\u0432\u043e."),
        u(r"\u0421\u0442\u0440\u043e\u043f\u043e\u0432\u043a\u0430 - \u0441\u043f\u043e\u0441\u043e\u0431 \u043e\u0431\u0432\u044f\u0437\u043a\u0438 \u0438 \u0437\u0430\u0446\u0435\u043f\u043a\u0438."),
    ],
    GREEN,
)
add_panel(
    s4,
    Inches(6.85),
    Inches(1.4),
    Inches(5.8),
    Inches(4.9),
    u(r"\u041f\u0440\u0430\u0432\u0430\u044f \u043a\u043e\u043b\u043e\u043d\u043a\u0430"),
    [
        u(r"\u0413\u0440\u0443\u0437 - \u043f\u0440\u0435\u0434\u043c\u0435\u0442 \u0438\u043b\u0438 \u043a\u043e\u043d\u0441\u0442\u0440\u0443\u043a\u0446\u0438\u044f \u0434\u043b\u044f \u043f\u043e\u0434\u044a\u0435\u043c\u0430."),
        u(r"\u0413\u0440\u0443\u0437\u043e\u0437\u0430\u0445\u0432\u0430\u0442\u043d\u043e\u0435 \u043f\u0440\u0438\u0441\u043f\u043e\u0441\u043e\u0431\u043b\u0435\u043d\u0438\u0435 - \u0443\u0441\u0442\u0440\u043e\u0439\u0441\u0442\u0432\u043e \u0434\u043b\u044f \u0437\u0430\u0445\u0432\u0430\u0442\u0430 \u0433\u0440\u0443\u0437\u0430."),
        u(r"\u0421\u0418\u0417 - \u0441\u0440\u0435\u0434\u0441\u0442\u0432\u0430 \u0438\u043d\u0434\u0438\u0432\u0438\u0434\u0443\u0430\u043b\u044c\u043d\u043e\u0439 \u0437\u0430\u0449\u0438\u0442\u044b."),
    ],
    GREEN,
)
back2 = add_button(s4, Inches(9.15), Inches(6.55), Inches(3.05), Inches(0.5), u(r"\u041d\u0410\u0417\u0410\u0414 \u041a \u041a\u0423\u0420\u0421\u0423"), NAVY)


start_btn.click_action.target_slide = s2
card_video.click_action.target_slide = s3
card_gloss.click_action.target_slide = s4
btn_video.click_action.target_slide = s3
btn_gloss.click_action.target_slide = s4
back1.click_action.target_slide = s2
back2.click_action.target_slide = s2


prs.save(OUT_FILE)
print(OUT_FILE)
