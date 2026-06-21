from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Callable
import math

from PIL import Image, ImageDraw, ImageFilter, ImageFont


ROOT = Path(__file__).resolve().parent.parent
OUTPUT_DIR = (
    ROOT
    / "presentations"
    / "ispring-course"
    / "module-01-stropovka-gruzov"
    / "live-preview"
    / "video-build"
    / "poster_animation_preview_2026-06-21"
)

SOURCE_IMAGE = Path(r"C:\Users\Дмитрий\Pictures\стропаль 20 см.png")

CANVAS_W = 1600
CANVAS_H = 900
FPS = 12
FRAME_MS = int(1000 / FPS)

BG = "#ede7dc"
TEXT = "#18222c"
MUTED = "#4f5b66"
ACCENT = "#cf3d2e"
ACCENT_SOFT = "#f5b2aa"
SAFE = "#0d8b61"
PANEL = "#fffdf8"


def load_font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    candidates = [
        ("C:/Windows/Fonts/arialbd.ttf", True),
        ("C:/Windows/Fonts/arial.ttf", False),
        ("C:/Windows/Fonts/segoeuib.ttf", True),
        ("C:/Windows/Fonts/segoeui.ttf", False),
    ]
    preferred = [path for path, is_bold in candidates if is_bold == bold]
    fallback = [path for path, _ in candidates if path not in preferred]
    for path in preferred + fallback:
        if Path(path).exists():
            return ImageFont.truetype(path, size=size)
    return ImageFont.load_default()


TITLE_FONT = load_font(34, bold=True)
BODY_FONT = load_font(24, bold=False)
SMALL_FONT = load_font(20, bold=False)


@dataclass(frozen=True)
class Segment:
    name: str
    frames: int
    renderer: Callable[[Image.Image, Image.Image, float], None]


def ease_out_cubic(t: float) -> float:
    t = max(0.0, min(1.0, t))
    return 1 - pow(1 - t, 3)


def ease_in_out(t: float) -> float:
    t = max(0.0, min(1.0, t))
    return 0.5 - 0.5 * math.cos(math.pi * t)


def lerp(a: float, b: float, t: float) -> float:
    return a + (b - a) * t


def poster_rect() -> tuple[int, int, int, int]:
    return (220, 140, 1380, 820)


def fit_poster(source: Image.Image) -> Image.Image:
    x0, y0, x1, y1 = poster_rect()
    target_w = x1 - x0
    target_h = y1 - y0
    ratio = min(target_w / source.width, target_h / source.height)
    size = (max(1, int(source.width * ratio)), max(1, int(source.height * ratio)))
    return source.resize(size, Image.Resampling.LANCZOS)


def paste_poster(canvas: Image.Image, poster: Image.Image, *, x: int = 220, y: int = 160) -> tuple[int, int, int, int]:
    shadow = Image.new("RGBA", (poster.width + 24, poster.height + 24), (0, 0, 0, 0))
    shadow_draw = ImageDraw.Draw(shadow)
    shadow_draw.rounded_rectangle((12, 12, poster.width + 12, poster.height + 12), 16, fill=(0, 0, 0, 100))
    shadow = shadow.filter(ImageFilter.GaussianBlur(12))
    canvas.alpha_composite(shadow, (x - 12, y - 4))
    canvas.alpha_composite(poster, (x, y))
    return (x, y, x + poster.width, y + poster.height)


def add_header(draw: ImageDraw.ImageDraw, subtitle: str) -> None:
    draw.text((80, 54), "Оживление учебного плаката", font=TITLE_FONT, fill=TEXT)
    draw.text((80, 98), subtitle, font=SMALL_FONT, fill=MUTED)


def add_footer(draw: ImageDraw.ImageDraw, text: str) -> None:
    draw.rounded_rectangle((70, 836, 1530, 882), 16, fill=PANEL, outline="#d9d0c4", width=2)
    draw.text((94, 847), text, font=BODY_FONT, fill=TEXT)


def dim_region(canvas: Image.Image, box: tuple[int, int, int, int], opacity: int) -> None:
    overlay = Image.new("RGBA", (box[2] - box[0], box[3] - box[1]), (255, 255, 255, opacity))
    canvas.alpha_composite(overlay, (box[0], box[1]))


def draw_focus_outline(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], color: str, width: int = 6) -> None:
    draw.rounded_rectangle(box, 18, outline=color, width=width)


def draw_arrow(draw: ImageDraw.ImageDraw, x: int, y_top: int, y_bottom: int, color: str, width: int = 10) -> None:
    draw.line((x, y_top, x, y_bottom), fill=color, width=width)
    head = 22
    draw.polygon(
        [(x, y_top), (x - head, y_top + head + 4), (x + head, y_top + head + 4)],
        fill=color,
    )
    draw.polygon(
        [(x, y_bottom), (x - head, y_bottom - head - 4), (x + head, y_bottom - head - 4)],
        fill=color,
    )


def draw_glow_circle(canvas: Image.Image, center: tuple[int, int], radius: int, color: tuple[int, int, int], alpha: int) -> None:
    glow = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    glow_draw = ImageDraw.Draw(glow)
    for step, scale in enumerate((2.8, 2.2, 1.6, 1.0), start=1):
        current_alpha = max(20, int(alpha / step))
        r = int(radius * scale)
        glow_draw.ellipse(
            (center[0] - r, center[1] - r, center[0] + r, center[1] + r),
            fill=(color[0], color[1], color[2], current_alpha),
        )
    glow = glow.filter(ImageFilter.GaussianBlur(16))
    canvas.alpha_composite(glow)


def draw_callout(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], title: str, body: str) -> None:
    draw.rounded_rectangle(box, 20, fill=PANEL, outline="#d3cabf", width=2)
    draw.text((box[0] + 22, box[1] + 16), title, font=BODY_FONT, fill=TEXT)
    draw.multiline_text((box[0] + 22, box[1] + 54), body, font=SMALL_FONT, fill=MUTED, spacing=6)


def create_base_canvas() -> Image.Image:
    return Image.new("RGBA", (CANVAS_W, CANVAS_H), BG)


def render_intro(canvas: Image.Image, poster: Image.Image, t: float) -> None:
    draw = ImageDraw.Draw(canvas)
    add_header(draw, "Прототип мягкой анимации для статического плаката")
    scale = lerp(0.96, 1.0, ease_out_cubic(t))
    scaled = poster.resize((int(poster.width * scale), int(poster.height * scale)), Image.Resampling.LANCZOS)
    x = 220 + (poster.width - scaled.width) // 2
    y = 160 + (poster.height - scaled.height) // 2
    paste_poster(canvas, scaled, x=x, y=y)
    fade = int(lerp(0, 210, ease_out_cubic(t)))
    banner = Image.new("RGBA", (930, 74), (255, 253, 248, fade))
    canvas.alpha_composite(banner, (340, 738))
    if t > 0.1:
        draw.text((378, 758), "Сначала фиксируем общий смысл плаката", font=BODY_FONT, fill=TEXT)


def render_left_focus(canvas: Image.Image, poster: Image.Image, t: float) -> None:
    draw = ImageDraw.Draw(canvas)
    add_header(draw, "Шаг 1. Сфокусировать внимание на проверочном подъеме")
    box = paste_poster(canvas, poster)
    left_panel = (box[0] + 10, box[1] + 10, box[0] + poster.width // 2 - 6, box[3] - 10)
    right_panel = (box[0] + poster.width // 2 + 6, box[1] + 10, box[2] - 10, box[3] - 10)
    dim_region(canvas, right_panel, int(lerp(220, 90, ease_out_cubic(t))))
    draw_focus_outline(draw, left_panel, ACCENT, width=8)
    callout_alpha = int(lerp(0, 255, ease_out_cubic(max(0.0, (t - 0.2) / 0.8))))
    if callout_alpha > 0:
        callout = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
        callout_draw = ImageDraw.Draw(callout)
        draw_callout(
            callout_draw,
            (1010, 210, 1475, 356),
            "Прием",
            "Не двигаем весь плакат.\nПросто приглушаем лишнее и оставляем один главный смысловой блок.",
        )
        callout.putalpha(callout_alpha)
        canvas.alpha_composite(callout)
    add_footer(draw, "Зритель сразу понимает: сейчас речь про пробный подъем и контроль состояния груза.")


def render_height_emphasis(canvas: Image.Image, poster: Image.Image, t: float) -> None:
    draw = ImageDraw.Draw(canvas)
    add_header(draw, "Шаг 2. Оживить цифру 0,2-0,3 м через акцент, а не через мультфильм")
    box = paste_poster(canvas, poster)
    left_panel = (box[0] + 10, box[1] + 10, box[0] + poster.width // 2 - 6, box[3] - 10)
    right_panel = (box[0] + poster.width // 2 + 6, box[1] + 10, box[2] - 10, box[3] - 10)
    dim_region(canvas, right_panel, 205)
    draw_focus_outline(draw, left_panel, SAFE, width=8)
    pulse = 0.75 + 0.25 * math.sin(t * math.pi * 4)
    arrow_alpha = int(110 + 110 * pulse)
    arrow_x = box[0] + 200
    y_top = box[1] + 456
    y_bottom = box[1] + 604
    arrow_overlay = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    arrow_draw = ImageDraw.Draw(arrow_overlay)
    draw_arrow(arrow_draw, arrow_x, y_top, y_bottom, (arrow_alpha, 53, 40))
    bubble_box = (box[0] + 30, box[1] + 520, box[0] + 340, box[1] + 620)
    arrow_draw.rounded_rectangle(bubble_box, 20, fill=(255, 253, 248, 240), outline=(207, 61, 46, 255), width=3)
    arrow_draw.text((bubble_box[0] + 30, bubble_box[1] + 26), "0,2-0,3 м", font=TITLE_FONT, fill=ACCENT)
    canvas.alpha_composite(arrow_overlay)
    add_footer(draw, "Даже простая пульсация стрелки делает число заметным и убирает ощущение полной статичности.")


def render_hook_highlight(canvas: Image.Image, poster: Image.Image, t: float) -> None:
    draw = ImageDraw.Draw(canvas)
    add_header(draw, "Шаг 3. Подсветить точку контроля: крюк, стропы, вертикальность")
    box = paste_poster(canvas, poster)
    right_panel = (box[0] + poster.width // 2 + 6, box[1] + 10, box[2] - 10, box[3] - 10)
    dim_region(canvas, right_panel, 205)
    glow_alpha = int(180 * ease_in_out(t))
    draw_glow_circle(canvas, (box[0] + 151, box[1] + 368), 28, (207, 61, 46), glow_alpha)
    draw_glow_circle(canvas, (box[0] + 190, box[1] + 430), 24, (13, 139, 97), glow_alpha)
    rope_overlay = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    rope_draw = ImageDraw.Draw(rope_overlay)
    rope_draw.line((box[0] + 180, box[1] + 170, box[0] + 180, box[1] + 350), fill=(255, 196, 90, glow_alpha), width=12)
    rope_draw.line((box[0] + 208, box[1] + 170, box[0] + 208, box[1] + 350), fill=(255, 196, 90, glow_alpha), width=12)
    canvas.alpha_composite(rope_overlay.filter(ImageFilter.GaussianBlur(8)))
    draw_callout(
        draw,
        (960, 232, 1480, 394),
        "Что оживает",
        "Не человек целиком, а только учебная точка внимания:\nкрюк, стропы, направление нагрузки.",
    )
    add_footer(draw, "Так плакат остается строгим, но взгляд зрителя идет туда, где нужна проверка.")


def render_right_response(canvas: Image.Image, poster: Image.Image, t: float) -> None:
    draw = ImageDraw.Draw(canvas)
    add_header(draw, "Шаг 4. При аварийной ситуации раскрывать действия по одному")
    box = paste_poster(canvas, poster)
    left_panel = (box[0] + 10, box[1] + 10, box[0] + poster.width // 2 - 6, box[3] - 10)
    right_panel = (box[0] + poster.width // 2 + 6, box[1] + 10, box[2] - 10, box[3] - 10)
    dim_region(canvas, left_panel, 205)
    draw_focus_outline(draw, right_panel, ACCENT, width=8)

    steps = [
        "Подать сигнал о немедленном опускании груза",
        "Освободить крюк",
        "Не продолжать работы до устранения неисправности",
    ]
    visible = min(len(steps), int(t * 3.4) + 1)
    for index, step in enumerate(steps[:visible]):
        y = 242 + index * 118
        step_box = (925, y, 1480, y + 84)
        draw.rounded_rectangle(step_box, 18, fill=PANEL, outline=ACCENT_SOFT, width=3)
        draw.ellipse((944, y + 24, 972, y + 52), fill=ACCENT, outline=ACCENT)
        draw.multiline_text((995, y + 18), step, font=SMALL_FONT, fill=TEXT, spacing=5)
    add_footer(draw, "Инструкции появляются последовательно. Так зритель не читает весь правый блок одновременно.")


def render_final(canvas: Image.Image, poster: Image.Image, t: float) -> None:
    draw = ImageDraw.Draw(canvas)
    add_header(draw, "Итог. Плакат оживает за счет ритма, акцентов и последовательности")
    paste_poster(canvas, poster)
    panel = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    panel_draw = ImageDraw.Draw(panel)
    alpha = int(lerp(0, 245, ease_out_cubic(t)))
    panel_draw.rounded_rectangle((905, 180, 1495, 470), 22, fill=(255, 253, 248, alpha), outline=(207, 61, 46, alpha), width=3)
    panel_draw.text((938, 210), "Что можно тиражировать", font=BODY_FONT, fill=TEXT)
    points = [
        "1. Приглушать второстепенную часть.",
        "2. Пульсировать важную цифру или знак.",
        "3. Подсвечивать крюк, строп, опасную зону.",
        "4. Показывать действия по шагам.",
    ]
    for idx, line in enumerate(points):
        panel_draw.text((940, 258 + idx * 44), line, font=SMALL_FONT, fill=MUTED)
    canvas.alpha_composite(panel)
    add_footer(draw, "Это уже можно повторять на других плакатах без сложной покадровой анимации.")


def build_frames(source_path: Path) -> list[Image.Image]:
    source = Image.open(source_path).convert("RGBA")
    poster = fit_poster(source)
    segments = [
        Segment("intro", 18, render_intro),
        Segment("left_focus", 18, render_left_focus),
        Segment("height_emphasis", 18, render_height_emphasis),
        Segment("hook_highlight", 18, render_hook_highlight),
        Segment("right_response", 24, render_right_response),
        Segment("final", 24, render_final),
    ]
    frames: list[Image.Image] = []
    for segment in segments:
        for frame_idx in range(segment.frames):
            t = 0.0 if segment.frames == 1 else frame_idx / (segment.frames - 1)
            canvas = create_base_canvas()
            segment.renderer(canvas, poster, t)
            frames.append(canvas.convert("P", palette=Image.Palette.ADAPTIVE))
    return frames


def main() -> None:
    if not SOURCE_IMAGE.exists():
        raise FileNotFoundError(f"Source image not found: {SOURCE_IMAGE}")

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    frames = build_frames(SOURCE_IMAGE)

    gif_path = OUTPUT_DIR / "poster_animation_preview.gif"
    preview_path = OUTPUT_DIR / "poster_animation_preview_last_frame.png"

    frames[0].save(
        gif_path,
        save_all=True,
        append_images=frames[1:],
        duration=FRAME_MS,
        loop=0,
        disposal=2,
    )
    frames[-1].convert("RGBA").save(preview_path)

    print(gif_path)
    print(preview_path)


if __name__ == "__main__":
    main()
