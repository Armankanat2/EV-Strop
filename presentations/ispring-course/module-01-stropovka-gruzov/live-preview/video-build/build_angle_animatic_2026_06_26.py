from __future__ import annotations

import math
import random
import re
import struct
import subprocess
import wave
from dataclasses import dataclass
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont, ImageFilter


ROOT = Path(__file__).resolve().parents[5]
BASE_DIR = Path(__file__).resolve().parent
OUT_DIR = BASE_DIR / "angle-animatic-2026-06-26"
FRAMES_DIR = OUT_DIR / "frames"

VOICE_MP3 = ROOT / "assets" / "course-media" / "module-01-stropovka-gruzov" / "audio" / "angle-animation-voiceover-short-male_2026-06-26_dmitry.mp3"
VTT_FILE = ROOT / "assets" / "course-media" / "module-01-stropovka-gruzov" / "audio" / "angle-animation-voiceover-short-male_2026-06-26_dmitry.vtt"
FX_WAV = OUT_DIR / "angle_animatic_fx.wav"
FULL_WAV = OUT_DIR / "angle_animatic_audio.wav"
VIDEO_FILE = OUT_DIR / "angle_animatic_2026-06-26.mp4"
POSTER_FILE = OUT_DIR / "angle_animatic_poster_2026-06-26.png"
PLAN_FILE = OUT_DIR / "angle_animatic_plan_2026-06-26.txt"
PANEL_BOX = (54, 132, 1226, 574)

WIDTH = 1280
HEIGHT = 720
FPS = 15
SR = 24000
BODY_W = 520
CAP_R = 68
BODY_H = 132
ANCHOR_INSET_RATIO = 0.25

VOICE_END = 37.596
PAUSE_AFTER_VOICE = 2.2
VIBRATION_SECONDS = 1.8
CLAP_SECONDS = 0.12
TAIL_SECONDS = 0.88
TOTAL_SECONDS = VOICE_END + PAUSE_AFTER_VOICE + VIBRATION_SECONDS + CLAP_SECONDS + TAIL_SECONDS

BG = (244, 240, 231, 255)
BG2 = (230, 224, 212, 255)
TEXT = (48, 56, 67, 255)
TEXT_SOFT = (92, 101, 114, 255)
WHITE = (255, 255, 255, 255)
GREEN = (45, 152, 84, 255)
GREEN_SOFT = (94, 176, 118, 255)
AMBER = (221, 161, 52, 255)
ORANGE = (228, 126, 49, 255)
RED = (205, 72, 58, 255)
DARK_RED = (145, 38, 31, 255)
STEEL = (86, 118, 148, 255)
SLING = (163, 101, 41, 255)
SLING_HI = (198, 147, 90, 255)
HOOK = (77, 89, 104, 255)
LOAD = (58, 60, 64, 255)
LOAD_EDGE = (31, 33, 36, 255)
PANEL = (255, 251, 244, 242)
PANEL_BORDER = (223, 212, 194, 255)
SHADOW = (35, 40, 45, 48)

FONT_UI = Path(r"C:\Windows\Fonts\segoeui.ttf")
FONT_BOLD = Path(r"C:\Windows\Fonts\arialbd.ttf")
FONT_REG = Path(r"C:\Windows\Fonts\arial.ttf")


@dataclass(frozen=True)
class Cue:
    start: float
    end: float
    text: str


@dataclass(frozen=True)
class Stage:
    angle: int
    loss: int
    load_pct: int
    critical: bool
    start: float
    hold_start: float
    hold_end: float
    center_y: float
    anchor_half: float
    wrap_half: float
    tension: float


@dataclass(frozen=True)
class HookGeometry:
    center_x: int
    top_y: int
    apex_y: int
    shank_bottom: int
    seat_left: tuple[int, int]
    seat_mid: tuple[int, int]
    seat_right: tuple[int, int]


STAGES = [
    Stage(angle=30, loss=0, load_pct=100, critical=False, start=0.0, hold_start=0.0, hold_end=6.334, center_y=585, anchor_half=108, wrap_half=144, tension=0.12),
    Stage(angle=45, loss=10, load_pct=90, critical=False, start=6.334, hold_start=7.8, hold_end=13.799, center_y=555, anchor_half=145, wrap_half=128, tension=0.28),
    Stage(angle=60, loss=15, load_pct=85, critical=False, start=13.799, hold_start=15.2, hold_end=21.172, center_y=522, anchor_half=182, wrap_half=108, tension=0.46),
    Stage(angle=90, loss=30, load_pct=70, critical=False, start=21.172, hold_start=23.0, hold_end=28.001, center_y=478, anchor_half=252, wrap_half=78, tension=0.74),
    Stage(angle=100, loss=40, load_pct=60, critical=True, start=28.001, hold_start=29.4, hold_end=VOICE_END, center_y=458, anchor_half=275, wrap_half=58, tension=1.0),
]


def get_font(size: int, bold: bool = False):
    path = FONT_BOLD if bold and FONT_BOLD.exists() else FONT_UI if FONT_UI.exists() else FONT_REG
    return ImageFont.truetype(str(path), size)


def ease(t: float) -> float:
    t = max(0.0, min(1.0, t))
    return 0.5 - 0.5 * math.cos(math.pi * t)


def ease_out(t: float) -> float:
    t = max(0.0, min(1.0, t))
    return 1.0 - (1.0 - t) ** 3


def mix(a: float, b: float, t: float) -> float:
    return a + (b - a) * t


def mix_color(a: tuple[int, int, int, int], b: tuple[int, int, int, int], t: float) -> tuple[int, int, int, int]:
    return tuple(int(round(mix(x, y, t))) for x, y in zip(a, b))


def parse_vtt(path: Path) -> list[Cue]:
    text = path.read_text(encoding="utf-8")
    pattern = re.compile(r"(\d\d:\d\d:\d\d,\d\d\d)\s+-->\s+(\d\d:\d\d:\d\d,\d\d\d)\s+(.+?)(?=\n\n|\Z)", re.S)
    cues: list[Cue] = []
    for start, end, raw_text in pattern.findall(text):
        cues.append(Cue(to_seconds(start), to_seconds(end), raw_text.replace("\n", " ").strip()))
    return cues


def to_seconds(ts: str) -> float:
    hh, mm, rest = ts.split(":")
    ss, ms = rest.split(",")
    return int(hh) * 3600 + int(mm) * 60 + int(ss) + int(ms) / 1000


def cue_at(cues: list[Cue], t: float) -> Cue | None:
    for cue in cues:
        if cue.start <= t <= cue.end:
            return cue
    return None


def stage_state(t: float) -> Stage:
    current = STAGES[-1]
    for idx, stage in enumerate(STAGES):
        current = stage
        if stage.start <= t <= stage.hold_end:
            if idx == 0 or t <= stage.hold_start:
                if idx == 0:
                    return stage
                prev = STAGES[idx - 1]
                local = ease_out((t - stage.start) / max(stage.hold_start - stage.start, 0.001))
                return Stage(
                    angle=int(round(mix(prev.angle, stage.angle, local))),
                    loss=int(round(mix(prev.loss, stage.loss, local))),
                    load_pct=int(round(mix(prev.load_pct, stage.load_pct, local))),
                    critical=stage.critical,
                    start=stage.start,
                    hold_start=stage.hold_start,
                    hold_end=stage.hold_end,
                    center_y=mix(prev.center_y, stage.center_y, local),
                    anchor_half=mix(prev.anchor_half, stage.anchor_half, local),
                    wrap_half=mix(prev.wrap_half, stage.wrap_half, local),
                    tension=mix(prev.tension, stage.tension, local),
                )
            return stage
    return current


def build_fx_wav(path: Path) -> None:
    total = PAUSE_AFTER_VOICE + VIBRATION_SECONDS + CLAP_SECONDS + TAIL_SECONDS
    total_samples = int(total * SR)
    pause_end = int(PAUSE_AFTER_VOICE * SR)
    vib_end = int((PAUSE_AFTER_VOICE + VIBRATION_SECONDS) * SR)
    clap_end = int((PAUSE_AFTER_VOICE + VIBRATION_SECONDS + CLAP_SECONDS) * SR)
    rng = random.Random(26)

    samples: list[int] = []
    for i in range(total_samples):
        value = 0.0
        if pause_end <= i < vib_end:
            p = (i - pause_end) / max(1, vib_end - pause_end)
            amp = 0.05 + 0.17 * p
            trem = 0.5 + 0.5 * math.sin(2 * math.pi * 8 * p)
            hum = math.sin(2 * math.pi * 74 * i / SR) * 0.6 + math.sin(2 * math.pi * 132 * i / SR) * 0.3
            noise = (rng.random() * 2 - 1) * 0.25
            value = amp * trem * (hum + noise)
        elif vib_end <= i < clap_end:
            p = (i - vib_end) / max(1, clap_end - vib_end)
            env = math.exp(-10 * p)
            noise = (rng.random() * 2 - 1)
            value = 0.72 * env * noise
        samples.append(max(-32767, min(32767, int(value * 32767))))

    with wave.open(str(path), "wb") as wav_file:
        wav_file.setnchannels(1)
        wav_file.setsampwidth(2)
        wav_file.setframerate(SR)
        wav_file.writeframes(b"".join(struct.pack("<h", sample) for sample in samples))


def combine_audio() -> None:
    build_fx_wav(FX_WAV)
    cmd = [
        "ffmpeg",
        "-y",
        "-i",
        str(VOICE_MP3),
        "-i",
        str(FX_WAV),
        "-filter_complex",
        "[0:a][1:a]concat=n=2:v=0:a=1[a]",
        "-map",
        "[a]",
        str(FULL_WAV),
    ]
    subprocess.run(cmd, check=True)


def background() -> Image.Image:
    img = Image.new("RGBA", (WIDTH, HEIGHT), BG)
    draw = ImageDraw.Draw(img, "RGBA")
    for y in range(HEIGHT):
        t = y / max(1, HEIGHT - 1)
        draw.line((0, y, WIDTH, y), fill=mix_color(BG, BG2, t))
    draw.ellipse((-140, -110, 440, 310), fill=(255, 255, 255, 94))
    draw.ellipse((900, 10, 1420, 360), fill=(220, 174, 88, 42))
    draw.ellipse((860, 490, 1400, 980), fill=(73, 132, 190, 26))
    return img


def draw_header(draw: ImageDraw.ImageDraw) -> None:
    draw.text((64, 42), "Угол между ветвями стропа", font=get_font(34, bold=True), fill=TEXT)
    draw.text((66, 86), "Черновой аниматик: рост нагрузки по мере увеличения угла", font=get_font(17), fill=TEXT_SOFT)


def draw_shadowed_panel(base: Image.Image, box: tuple[int, int, int, int]) -> None:
    x0, y0, x1, y1 = box
    shadow = Image.new("RGBA", (x1 - x0 + 34, y1 - y0 + 34), (0, 0, 0, 0))
    sdraw = ImageDraw.Draw(shadow, "RGBA")
    sdraw.rounded_rectangle((16, 18, shadow.width - 2, shadow.height - 2), radius=30, fill=SHADOW)
    shadow = shadow.filter(ImageFilter.GaussianBlur(14))
    base.alpha_composite(shadow, (x0 - 16, y0 - 10))
    draw = ImageDraw.Draw(base, "RGBA")
    draw.rounded_rectangle(box, radius=28, fill=PANEL, outline=PANEL_BORDER, width=2)


def make_hook_geometry(center_x: int, top_y: int, apex_y: int) -> HookGeometry:
    scale = 1.0
    shank_top = top_y + 8
    shank_bottom = apex_y + 2
    seat_left = (center_x - int(20 * scale), apex_y + int(4 * scale))
    seat_mid = (center_x, apex_y + int(2 * scale))
    seat_right = (center_x + int(20 * scale), apex_y + int(4 * scale))
    return HookGeometry(
        center_x=center_x,
        top_y=top_y,
        apex_y=apex_y,
        shank_bottom=shank_bottom,
        seat_left=seat_left,
        seat_mid=seat_mid,
        seat_right=seat_right,
    )


def draw_hook(draw: ImageDraw.ImageDraw, hook: HookGeometry) -> None:
    scale = 1.0
    stem_w = max(6, int(10 * scale))
    center_x = hook.center_x
    shank_top = hook.top_y + 8

    # Only the vertical suspension remains; the hook is removed.
    draw.line((center_x, shank_top, center_x, hook.shank_bottom), fill=HOOK, width=stem_w)


def draw_saddle_contact(draw: ImageDraw.ImageDraw, hook: HookGeometry, tension: float) -> None:
    width = 8 + int(6 * tension)
    color = mix_color(SLING, RED, max(0.0, tension - 0.45) / 0.55)
    contact_left = (hook.seat_left[0], hook.seat_left[1])
    contact_mid = (hook.seat_mid[0], hook.seat_mid[1])
    contact_right = (hook.seat_right[0], hook.seat_right[1])
    draw.line((contact_left, contact_mid, contact_right), fill=color, width=width)
    draw.line((contact_left[0] + 2, contact_left[1], contact_mid[0], contact_mid[1] - 1, contact_right[0] - 2, contact_right[1]), fill=SLING_HI, width=max(2, width // 3))


def draw_load(draw: ImageDraw.ImageDraw, center_x: int, center_y: float, t_vibe: float, t_drop: float) -> tuple[int, int, int, int]:
    cy = center_y + t_drop * 220
    if t_vibe > 0:
        cy += math.sin(t_vibe * math.pi * 14) * (2 + 9 * t_vibe)
    x0 = int(center_x - BODY_W / 2)
    x1 = int(center_x + BODY_W / 2)
    y0 = int(cy - BODY_H / 2)
    y1 = int(cy + BODY_H / 2)

    # Long dark cylinder based on the reference: matte graphite body, soft top sheen, no pipe opening.
    draw.rectangle((x0, y0, x1, y1), fill=LOAD, outline=LOAD_EDGE, width=4)

    left_outer = (x0 - CAP_R, y0, x0 + CAP_R, y1)
    draw.ellipse(left_outer, fill=(52, 54, 58, 255), outline=LOAD_EDGE, width=4)
    draw.ellipse((x0 - 42, y0 + 18, x0 + 42, y1 - 18), fill=(45, 47, 51, 255))

    right_outer = (x1 - CAP_R, y0, x1 + CAP_R, y1)
    draw.ellipse(right_outer, fill=(61, 63, 67, 255))
    draw.arc(right_outer, start=270, end=90, fill=LOAD_EDGE, width=4)

    # Soft cylindrical shading and sheen similar to the black studio render.
    draw.line((x0 + 16, y0 + 20, x1 - 16, y0 + 20), fill=(132, 136, 142, 78), width=6)
    draw.line((x0 + 22, y0 + 34, x1 - 24, y0 + 34), fill=(188, 192, 198, 54), width=3)
    draw.line((x0 + 24, int(cy) - 2, x1 - 18, int(cy) - 2), fill=(96, 100, 106, 54), width=3)
    draw.line((x0 + 30, y1 - 18, x1 - 24, y1 - 18), fill=(23, 24, 27, 84), width=5)

    # Gentle front-face modeling keeps the left cap solid, not hollow.
    draw.arc((x0 - 30, y0 + 24, x0 + 28, y1 - 24), start=78, end=282, fill=(124, 128, 134, 88), width=2)
    return x0 - CAP_R, y0, x1 + CAP_R, y1


def anchor_points(center_x: int, center_y: float) -> tuple[tuple[int, int], tuple[int, int]]:
    total_half = (BODY_W + CAP_R * 2) / 2
    total_w = total_half * 2
    left_edge = center_x - total_half
    left_x = int(left_edge + total_w * ANCHOR_INSET_RATIO)
    right_x = int(left_edge + total_w * (1 - ANCHOR_INSET_RATIO))
    y = int(center_y - BODY_H / 2)
    return (left_x, y), (right_x, y)


def wrap_side_count(tension: float) -> int:
    if tension < 0.2:
        return 1
    if tension < 0.38:
        return 3
    if tension < 0.6:
        return 4
    return 5


def draw_wraps(draw: ImageDraw.ImageDraw, center_x: int, center_y: float, wrap_half: float, tension: float, vibe: float, t: float) -> None:
    y0 = int(center_y - BODY_H / 2)
    y1 = int(center_y + BODY_H / 2)
    left_anchor, right_anchor = anchor_points(center_x, center_y)
    center_gap = max(26, int(wrap_half * 0.55))
    left_limit = center_x - center_gap / 2
    right_limit = center_x + center_gap / 2
    side_count = wrap_side_count(tension)
    band_w = 16 + int(7 * tension)
    left_span = max(24, left_limit - left_anchor[0])
    right_span = max(24, right_anchor[0] - right_limit)
    shift = math.sin(vibe * math.pi * 12) * 4 if vibe > 0 else 0
    flow = math.sin(t * math.pi * 0.42)
    span_power = mix(1.05, 1.55, tension)
    positions: list[int] = []

    for idx in range(side_count):
        local = idx / max(1, side_count - 1) if side_count > 1 else 0.0
        inward = 1.0 - (1.0 - local) ** span_power
        px = int(left_anchor[0] + left_span * inward + shift * (0.22 + idx * 0.08))
        positions.append(px)
    for idx in range(side_count):
        local = idx / max(1, side_count - 1) if side_count > 1 else 0.0
        inward = 1.0 - (1.0 - local) ** span_power
        px = int(right_anchor[0] - right_span * inward + shift * (0.22 + idx * 0.08))
        positions.append(px)

    positions = sorted(set(positions))
    center_span = max(1.0, float(max(abs(px - center_x) for px in positions))) if positions else 1.0
    for idx, px in enumerate(positions):
        side_sign = -1 if px < center_x else 1
        inward = 1.0 - min(1.0, abs(px - center_x) / center_span)
        phase = flow + idx * 0.52
        alive = math.sin(phase) * (0.7 + 0.6 * tension)
        skew = int(round(side_sign * (3 + 4.5 * tension) * (0.4 + inward * 0.46) + alive))
        top_x0 = px - band_w // 2 + skew
        top_x1 = px + band_w // 2 + skew
        bot_x0 = px - band_w // 2 - skew
        bot_x1 = px + band_w // 2 - skew
        poly = [(top_x0, y0 - 3), (top_x1, y0 - 3), (bot_x1, y1 + 3), (bot_x0, y1 + 3)]
        draw.polygon(poly, fill=SLING, outline=(104, 62, 24, 255))
        draw.line((top_x0 + 4, y0 + 10, bot_x0 + 8, y1 - 12), fill=SLING_HI, width=2)
        draw.line((top_x0 + band_w - 6, y0 + 12, bot_x0 + band_w - 2, y1 - 10), fill=(120, 71, 28, 210), width=1)


def draw_branches(draw: ImageDraw.ImageDraw, center_x: int, hook: HookGeometry, center_y: float, anchor_half: float, tension: float, vibe: float, broken: bool) -> tuple[tuple[int, int], tuple[int, int]]:
    left, right = anchor_points(center_x, center_y)
    width = 8 + int(7 * tension)
    color = mix_color(SLING, RED, max(0.0, tension - 0.45) / 0.55)
    left_top = (hook.seat_left[0] + 1, hook.seat_left[1] + 1)
    right_top = (hook.seat_right[0] - 1, hook.seat_right[1] + 1)
    seat_top = (hook.seat_mid[0], hook.seat_mid[1] + 2)

    if broken:
        draw.line((left_top, (center_x - 24, hook.seat_mid[1] + 86)), fill=DARK_RED, width=width)
        draw.line((right_top, (center_x + 24, hook.seat_mid[1] + 86)), fill=DARK_RED, width=width)
        for px, py in [left, right]:
            draw.line(((px, py - 42), (px - 14, py + 10)), fill=DARK_RED, width=max(4, width - 2))
        return left, right

    jitter = 0.0
    if vibe > 0:
        jitter = math.sin(vibe * math.pi * 12) * (3 + 10 * vibe)

    left_apex = (int(left_top[0] - jitter * 0.24), left_top[1])
    right_apex = (int(right_top[0] + jitter * 0.24), right_top[1])
    seat_apex = (hook.seat_mid[0], int(seat_top[1] + abs(jitter) * 0.04))
    glow = mix_color((224, 180, 120, 88), (255, 122, 98, 128), tension)
    draw.line((left_apex, seat_apex, right_apex), fill=glow, width=width + 6)
    draw.line((left_apex, seat_apex, right_apex), fill=color, width=width)
    draw.line((left_apex, left), fill=glow, width=width + 6)
    draw.line((right_apex, right), fill=glow, width=width + 6)
    draw.line((left_apex, left), fill=color, width=width)
    draw.line((right_apex, right), fill=color, width=width)
    draw.line((left_apex[0] + 2, left_apex[1] - 1, seat_apex[0] + 1, seat_apex[1] - 1, right_apex[0] - 2, right_apex[1]), fill=SLING_HI, width=max(2, width // 3))
    draw.line((left_apex[0] + 2, left_apex[1] + 4, left[0] + 4, left[1]), fill=SLING_HI, width=max(2, width // 3))
    draw.line((right_apex[0] - 2, right_apex[1] + 4, right[0] - 4, right[1]), fill=SLING_HI, width=max(2, width // 3))
    if tension >= 0.55:
        marker_color = mix_color(AMBER, RED, min(1.0, (tension - 0.55) / 0.45))
        for shift_x, shift_y, direction in [(-34, 62, -1), (34, 62, 1)]:
            px = hook.seat_mid[0] + shift_x
            py = hook.seat_mid[1] + shift_y
            draw.line((px, py, px + 16 * direction, py + 10), fill=marker_color, width=3)
            draw.line((px, py, px + 18 * direction, py - 10), fill=marker_color, width=3)
    return left, right


def draw_angle_arc(draw: ImageDraw.ImageDraw, center_x: int, apex_y: int, stage: Stage, vibe: float) -> None:
    r = 100
    box = (center_x - r, apex_y + 18 - r, center_x + r, apex_y + 18 + r)
    half = stage.angle / 2
    start = 90 - half
    end = 90 + half
    color = mix_color(GREEN, RED, stage.tension)
    if vibe > 0:
        color = mix_color(color, WHITE, 0.25 * math.sin(vibe * math.pi * 10) ** 2)
    draw.arc(box, start=start, end=end, fill=color, width=5)


def draw_centerline(draw: ImageDraw.ImageDraw, center_x: int, top_y: int, bottom_y: int) -> None:
    y = top_y
    while y < bottom_y:
        draw.line((center_x, y, center_x, min(y + 10, bottom_y)), fill=(132, 145, 158, 88), width=2)
        y += 18


def draw_angle_label(draw: ImageDraw.ImageDraw, center_x: int, apex_y: int) -> None:
    text = "Угол между ветвями"
    font = get_font(18, bold=True)
    w = draw.textbbox((0, 0), text, font=font)[2]
    draw.rounded_rectangle((center_x - w // 2 - 12, apex_y + 108, center_x + w // 2 + 12, apex_y + 140), radius=15, fill=(255, 251, 244, 214), outline=(214, 203, 186, 255), width=2)
    draw.text((center_x - w / 2, apex_y + 115), text, font=font, fill=TEXT_SOFT)


def camera_state(t: float) -> tuple[float, float, float]:
    # Early stops should show the whole pipe clearly: far -> medium -> closer.
    # Final frames keep the already approved tighter composition.
    zoom_stops = [0.68, 0.76, 0.86, 1.02, 1.08]

    if t <= STAGES[0].hold_end:
        zoom = zoom_stops[0]
    elif t <= STAGES[1].hold_end:
        if t < STAGES[1].hold_start:
            local = ease_out((t - STAGES[1].start) / max(STAGES[1].hold_start - STAGES[1].start, 0.001))
            zoom = mix(zoom_stops[0], zoom_stops[1], local)
        else:
            zoom = zoom_stops[1]
    elif t <= STAGES[2].hold_end:
        if t < STAGES[2].hold_start:
            local = ease_out((t - STAGES[2].start) / max(STAGES[2].hold_start - STAGES[2].start, 0.001))
            zoom = mix(zoom_stops[1], zoom_stops[2], local)
        else:
            zoom = zoom_stops[2]
    elif t <= STAGES[3].hold_end:
        if t < STAGES[3].hold_start:
            local = ease_out((t - STAGES[3].start) / max(STAGES[3].hold_start - STAGES[3].start, 0.001))
            zoom = mix(zoom_stops[2], zoom_stops[3], local)
        else:
            zoom = zoom_stops[3]
    elif t <= VOICE_END:
        if t < STAGES[4].hold_start:
            local = ease_out((t - STAGES[4].start) / max(STAGES[4].hold_start - STAGES[4].start, 0.001))
            zoom = mix(zoom_stops[3], zoom_stops[4], local)
        else:
            zoom = zoom_stops[4]
    else:
        zoom = zoom_stops[4]

    shift_x = 0.0
    shift_y = 0.0
    return zoom, shift_x, shift_y


def compose_scene_with_camera(base: Image.Image, scene: Image.Image, t: float) -> None:
    panel_x0, panel_y0, panel_x1, panel_y1 = PANEL_BOX
    panel_w = panel_x1 - panel_x0
    panel_h = panel_y1 - panel_y0
    zoom, shift_x, shift_y = camera_state(t)
    crop_w = panel_w / zoom
    crop_h = panel_h / zoom
    center_x = (panel_x0 + panel_x1) / 2 + shift_x
    center_y = (panel_y0 + panel_y1) / 2 + shift_y
    crop_box = (
        int(round(center_x - crop_w / 2)),
        int(round(center_y - crop_h / 2)),
        int(round(center_x + crop_w / 2)),
        int(round(center_y + crop_h / 2)),
    )
    cropped = scene.crop(crop_box).resize((panel_w, panel_h), Image.Resampling.LANCZOS)
    base.alpha_composite(cropped, (panel_x0, panel_y0))


def draw_badges(draw: ImageDraw.ImageDraw, stage: Stage) -> None:
    badge_color = GREEN if stage.angle <= 30 else GREEN_SOFT if stage.angle <= 45 else AMBER if stage.angle <= 60 else ORANGE if stage.angle < 100 else RED
    draw.rounded_rectangle((1040, 36, 1200, 102), radius=22, fill=badge_color)
    angle_text = f"{stage.angle}°"
    font = get_font(34, bold=True)
    bbox = draw.textbbox((0, 0), angle_text, font=font)
    draw.text((1120 - (bbox[2] - bbox[0]) / 2, 50), angle_text, font=font, fill=WHITE)

    draw.rounded_rectangle((954, 122, 1216, 212), radius=20, fill=(255, 255, 255, 220), outline=badge_color, width=3)
    draw.text((972, 142), f"Потеря: {stage.loss}%", font=get_font(22, bold=True), fill=TEXT)
    draw.text((972, 172), f"Нагрузка: {stage.load_pct}%", font=get_font(22, bold=True), fill=TEXT)
    if stage.critical:
        draw.rounded_rectangle((980, 224, 1190, 264), radius=16, fill=RED)
        draw.text((1044, 233), "КРИТИЧНО", font=get_font(20, bold=True), fill=WHITE)


def draw_meter(draw: ImageDraw.ImageDraw, stage: Stage) -> None:
    x0, y0, x1, y1 = 64, 618, 468, 640
    draw.text((64, 582), "Нагрузка на ветви", font=get_font(22, bold=True), fill=TEXT)
    draw.rounded_rectangle((x0, y0, x1, y1), radius=12, fill=(221, 214, 204, 255))
    fill = int((x1 - x0) * (1 - stage.load_pct / 100))
    color = mix_color(GREEN, RED, stage.tension)
    if fill > 0:
        draw.rounded_rectangle((x0, y0, x0 + fill, y1), radius=12, fill=color)
    label = "минимум" if stage.angle == 30 else "растет" if stage.angle < 90 else "высоко" if stage.angle < 100 else "предел"
    draw.text((486, 611), label, font=get_font(24, bold=True), fill=color)


def draw_subtitle(draw: ImageDraw.ImageDraw, cue: Cue | None) -> None:
    if cue is None:
        return
    box = (70, 654, 1210, 708)
    draw.rounded_rectangle(box, radius=22, fill=(17, 26, 37, 214))
    font = get_font(24, bold=True)
    text = cue.text
    max_width = box[2] - box[0] - 30
    words = text.split()
    lines = []
    current = ""
    for word in words:
        probe = word if not current else f"{current} {word}"
        width = draw.textbbox((0, 0), probe, font=font)[2]
        if width <= max_width:
            current = probe
        else:
            lines.append(current)
            current = word
    if current:
        lines.append(current)
    total_h = len(lines) * 26 + max(0, len(lines) - 1) * 4
    y = box[1] + ((box[3] - box[1]) - total_h) / 2 - 1
    for line in lines:
        width = draw.textbbox((0, 0), line, font=font)[2]
        draw.text((box[0] + ((box[2] - box[0]) - width) / 2, y), line, font=font, fill=WHITE)
        y += 30


def draw_scene_note(draw: ImageDraw.ImageDraw, t: float) -> None:
    if t < VOICE_END:
        return
    x0, y0, x1, y1 = 784, 596, 1210, 650
    if t < VOICE_END + PAUSE_AFTER_VOICE:
        text = "Критическая нагрузка"
        color = RED
    elif t < VOICE_END + PAUSE_AFTER_VOICE + VIBRATION_SECONDS:
        text = "Вибрация стропа"
        color = ORANGE
    else:
        text = "Разрыв"
        color = DARK_RED
    draw.rounded_rectangle((x0, y0, x1, y1), radius=18, fill=(255, 252, 246, 236), outline=color, width=3)
    font = get_font(24, bold=True)
    w = draw.textbbox((0, 0), text, font=font)[2]
    draw.text((x0 + ((x1 - x0) - w) / 2, 611), text, font=font, fill=TEXT)


def render_frame(t: float, cues: list[Cue]) -> Image.Image:
    base = background()
    draw_shadowed_panel(base, PANEL_BOX)
    draw = ImageDraw.Draw(base, "RGBA")
    draw_header(draw)

    center_x = WIDTH // 2
    hook_top = 68
    apex_y = 244

    stage = stage_state(min(t, VOICE_END))
    vibe_start = VOICE_END + PAUSE_AFTER_VOICE
    break_start = vibe_start + VIBRATION_SECONDS
    break_end = break_start + CLAP_SECONDS
    tail_end = break_end + TAIL_SECONDS

    vibe = 0.0
    if vibe_start <= t < break_start:
        vibe = (t - vibe_start) / VIBRATION_SECONDS
    drop = 0.0
    broken = False
    if t >= break_start:
        broken = True
        drop = min(1.0, (t - break_start) / max(tail_end - break_start, 0.001))

    scene = Image.new("RGBA", (WIDTH, HEIGHT), (0, 0, 0, 0))
    sdraw = ImageDraw.Draw(scene, "RGBA")
    draw_centerline(sdraw, center_x, hook_top + 58, int(stage.center_y + drop * 220))
    hook = make_hook_geometry(center_x, hook_top, apex_y)
    draw_angle_arc(sdraw, center_x, apex_y, stage, vibe)
    load_box = draw_load(sdraw, center_x, stage.center_y, vibe, drop)
    draw_wraps(sdraw, center_x, stage.center_y + drop * 220, stage.wrap_half, stage.tension, vibe, t)
    draw_branches(sdraw, center_x, hook, stage.center_y + drop * 220, stage.anchor_half, stage.tension, vibe, broken)
    draw_hook(sdraw, hook)
    draw_saddle_contact(sdraw, hook, stage.tension)
    draw_angle_label(sdraw, center_x, apex_y)
    compose_scene_with_camera(base, scene, t)

    draw = ImageDraw.Draw(base, "RGBA")
    draw_badges(draw, stage)
    draw_meter(draw, stage)
    draw_subtitle(draw, cue_at(cues, t))
    draw_scene_note(draw, t)

    if broken:
        flash = max(0.0, 1.0 - (t - break_start) / 0.18)
        if flash > 0:
            overlay = Image.new("RGBA", (WIDTH, HEIGHT), (255, 255, 255, int(140 * flash)))
            base.alpha_composite(overlay)
            draw = ImageDraw.Draw(base, "RGBA")
            lx = load_box[0] + 250
            ly = load_box[1] - 14
            draw.line((hook.seat_left[0], hook.seat_left[1] + 52, lx, ly), fill=(255, 245, 230, int(220 * flash)), width=5)
            draw.line((hook.seat_right[0], hook.seat_right[1] + 52, lx + 34, ly + 8), fill=(255, 245, 230, int(220 * flash)), width=5)

    if t < 1.4:
        alpha = int(140 * (1 - t / 1.4))
        draw.rounded_rectangle((92, 118, 446, 176), radius=18, fill=(255, 255, 255, alpha))
        draw.text((116, 133), "Влияние угла на строп", font=get_font(24, bold=True), fill=(TEXT[0], TEXT[1], TEXT[2], min(255, alpha + 70)))

    return base


def render_frames() -> list[Path]:
    cues = parse_vtt(VTT_FILE)
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    FRAMES_DIR.mkdir(parents=True, exist_ok=True)
    frame_paths: list[Path] = []
    frame_count = math.ceil(TOTAL_SECONDS * FPS)
    for idx in range(frame_count):
        t = idx / FPS
        image = render_frame(t, cues)
        path = FRAMES_DIR / f"frame_{idx:04d}.png"
        image.save(path)
        frame_paths.append(path)
    poster = render_frame(29.8, cues)
    poster.save(POSTER_FILE)
    return frame_paths


def mux_video() -> None:
    cmd = [
        "ffmpeg",
        "-y",
        "-framerate",
        str(FPS),
        "-i",
        str(FRAMES_DIR / "frame_%04d.png"),
        "-i",
        str(FULL_WAV),
        "-c:v",
        "libx264",
        "-pix_fmt",
        "yuv420p",
        "-c:a",
        "aac",
        "-b:a",
        "192k",
        "-movflags",
        "+faststart",
        str(VIDEO_FILE),
    ]
    subprocess.run(cmd, check=True)


def write_plan() -> None:
    lines = [
        f"voice={VOICE_MP3}",
        f"subtitles={VTT_FILE}",
        f"fx={FX_WAV}",
        f"full_audio={FULL_WAV}",
        f"video={VIDEO_FILE}",
        f"poster={POSTER_FILE}",
        f"fps={FPS}",
        f"voice_end={VOICE_END}",
        f"pause_after_voice={PAUSE_AFTER_VOICE}",
        f"vibration_seconds={VIBRATION_SECONDS}",
        f"clap_seconds={CLAP_SECONDS}",
        f"tail_seconds={TAIL_SECONDS}",
        f"total_seconds={TOTAL_SECONDS}",
        "",
        "stage_ranges:",
        "30deg=0.000-6.334",
        "45deg=6.334-13.799",
        "60deg=13.799-21.172",
        "90deg=21.172-28.001",
        "100deg=28.001-37.596",
        "silent_hold=37.596-39.796",
        "vibration=39.796-41.596",
        "break=41.596-41.716",
        "tail=41.716-42.596",
    ]
    PLAN_FILE.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> None:
    render_frames()
    combine_audio()
    mux_video()
    write_plan()
    print(VIDEO_FILE)
    print(POSTER_FILE)
    print(PLAN_FILE)


if __name__ == "__main__":
    main()
