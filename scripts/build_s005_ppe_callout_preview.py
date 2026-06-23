from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parents[1]
AVATAR_SRC = ROOT / "assets" / "reference_visuals" / "visual-bank" / "images" / "VIS-0005_avatar-1_worker-ppe_white-bg.png"
OUT_DIR = ROOT / "temp"
OUT_PATH = OUT_DIR / "s005_ppe_callout_preview_v8.png"

CANVAS_W = 1430
CANVAS_H = 910
HEADER_H = 144
PANEL_PAD = 24
BODY_PAD_X = 34
BODY_PAD_TOP = 28
BODY_PAD_BOTTOM = 26

AVATAR_H = 644
AVATAR_X = 488
AVATAR_Y = HEADER_H + 54

BG = (246, 248, 250)
WHITE = (255, 255, 255)
NAVY = (18, 43, 66)
TEXT = (30, 41, 54)
MUTED = (95, 109, 123)
ORANGE = (230, 126, 57)
FRAME = (204, 214, 224)
TITLE_FONT = ImageFont.truetype(r"C:\Windows\Fonts\ARIALNB.TTF", 24)
BODY_FONT = ImageFont.truetype(r"C:\Windows\Fonts\ARIALN.TTF", 18)
HEADER_FONT = ImageFont.truetype(r"C:\Windows\Fonts\arialbd.ttf", 32)
TAG_FONT = ImageFont.truetype(r"C:\Windows\Fonts\ARIALN.TTF", 18)

LABELS = [
    {
        "title": "КАСКА ЗАЩИТНАЯ",
        "body": "ГОСТ EN 397-2012",
        "top": 214,
        "x": 174,
        "w": 238,
        "align": "right",
        "target": (548, 222),
    },
    {
        "title": "ОЧКИ",
        "body": "ГОСТ 12.4.253-2013 (EN 166:2001)",
        "top": 314,
        "x": 92,
        "w": 320,
        "align": "right",
        "target": (548, 272),
    },
    {
        "title": "ОДЕЖДА ЗАЩИТНАЯ",
        "body": "ГОСТ 12.4.280-2014",
        "top": 330,
        "x": 828,
        "w": 300,
        "align": "left",
        "target": (638, 426),
    },
    {
        "title": "ПЕРЧАТКИ",
        "body": "ГОСТ 12.4.252-2013 и ГОСТ EN 388-2012",
        "top": 480,
        "x": 842,
        "w": 352,
        "align": "left",
        "target": (632, 582),
    },
    {
        "title": "ОБУВЬ ЗАЩИТНАЯ",
        "body": "ГОСТ Р 12.4.310-2020",
        "top": 630,
        "x": 844,
        "w": 290,
        "align": "left",
        "target": (603, 811),
    },
]


def crop_avatar() -> Image.Image:
    image = Image.open(AVATAR_SRC).convert("RGBA")
    rgb = image.convert("RGB")
    bbox = rgb.point(lambda value: 0 if value > 245 else 255).getbbox()
    if not bbox:
        return image

    left, top, right, bottom = bbox
    return image.crop(
        (
            max(0, left - 48),
            max(0, top - 36),
            min(image.width, right + 48),
            min(image.height, bottom + 28),
        )
    )


def line_height(font: ImageFont.FreeTypeFont) -> int:
    bbox = font.getbbox("Ag")
    return bbox[3] - bbox[1]


def wrap_text(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.FreeTypeFont, max_width: int) -> list[str]:
    words = text.split()
    lines: list[str] = []
    current = ""
    for word in words:
        candidate = word if not current else f"{current} {word}"
        if draw.textlength(candidate, font=font) <= max_width:
            current = candidate
            continue
        if current:
            lines.append(current)
        current = word
    if current:
        lines.append(current)
    return lines


def draw_background(draw: ImageDraw.ImageDraw) -> None:
    draw.rounded_rectangle(
        [PANEL_PAD, PANEL_PAD, CANVAS_W - PANEL_PAD, CANVAS_H - PANEL_PAD],
        radius=24,
        fill=WHITE,
        outline=FRAME,
        width=2,
    )

    draw.rounded_rectangle(
        [PANEL_PAD, PANEL_PAD, CANVAS_W - PANEL_PAD, PANEL_PAD + HEADER_H],
        radius=24,
        fill=NAVY,
        outline=NAVY,
        width=2,
    )
    draw.rectangle([PANEL_PAD, PANEL_PAD + 86, CANVAS_W - PANEL_PAD, PANEL_PAD + HEADER_H], fill=NAVY)

    draw.text((62, 62), "Общепроизводственные СИЗ", font=HEADER_FONT, fill=WHITE)
    draw.text((62, 102), "S005-P05 / PPE LOADOUT", font=TAG_FONT, fill=(199, 214, 226))


def draw_arrow(draw: ImageDraw.ImageDraw, start: tuple[int, int], end: tuple[int, int]) -> None:
    draw.line([start, end], fill=ORANGE, width=2)

    dx = end[0] - start[0]
    dy = end[1] - start[1]
    length = max((dx * dx + dy * dy) ** 0.5, 1.0)
    ux = dx / length
    uy = dy / length

    arrow_len = 12
    wing = 5
    base_x = end[0] - ux * arrow_len
    base_y = end[1] - uy * arrow_len
    perp_x = -uy
    perp_y = ux
    arrow = [
        end,
        (int(base_x + perp_x * wing), int(base_y + perp_y * wing)),
        (int(base_x - perp_x * wing), int(base_y - perp_y * wing)),
    ]
    draw.polygon(arrow, fill=ORANGE)


def draw_label(
    draw: ImageDraw.ImageDraw, top: int, x: int, width: int, title: str, body: str, align: str
) -> tuple[tuple[int, int], tuple[int, int]]:
    body_lines = wrap_text(draw, body, BODY_FONT, width)

    title_w = draw.textlength(title, font=TITLE_FONT)
    title_x = x if align == "left" else int(x + width - title_w)
    draw.text((title_x, top), title, font=TITLE_FONT, fill=TEXT)
    title_bottom = top + line_height(TITLE_FONT)
    current_y = title_bottom + 4

    body_widths = []
    for idx, line in enumerate(body_lines):
        body_w = draw.textlength(line, font=BODY_FONT)
        body_x = x if align == "left" else int(x + width - body_w)
        draw.text((body_x, current_y), line, font=BODY_FONT, fill=MUTED)
        body_widths.append(body_w)
        current_y += line_height(BODY_FONT)
        if idx != len(body_lines) - 1:
            current_y += 2

    underline_y = current_y + 6
    widest = max(title_w, max(body_widths))
    if align == "left":
        left_edge = x
        right_edge = int(x + widest + 8)
    else:
        right_edge = x + width
        left_edge = int(right_edge - widest - 8)
    draw.line([(left_edge, underline_y), (right_edge, underline_y)], fill=ORANGE, width=2)
    return (left_edge, underline_y), (right_edge, underline_y)


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    canvas = Image.new("RGBA", (CANVAS_W, CANVAS_H), BG + (255,))
    draw = ImageDraw.Draw(canvas)
    draw_background(draw)

    avatar = crop_avatar()
    avatar_w = int(AVATAR_H * avatar.width / avatar.height)
    avatar = avatar.resize((avatar_w, AVATAR_H), Image.Resampling.LANCZOS)
    canvas.alpha_composite(avatar, (AVATAR_X, AVATAR_Y))

    for label in LABELS:
        left_edge, right_edge = draw_label(
            draw,
            label["top"],
            label["x"],
            label["w"],
            label["title"],
            label["body"],
            label["align"],
        )
        start = right_edge if label["align"] == "right" else left_edge
        draw_arrow(draw, start, label["target"])

    canvas.convert("RGB").save(OUT_PATH, quality=95)
    print(OUT_PATH)


if __name__ == "__main__":
    main()
