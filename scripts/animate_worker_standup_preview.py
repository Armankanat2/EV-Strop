from __future__ import annotations

from pathlib import Path
import math

from PIL import Image, ImageChops, ImageDraw, ImageFilter


ROOT = Path(__file__).resolve().parent.parent
SOURCE = ROOT / "temp" / "stropal-20sm-source.png"
OUT_DIR = (
    ROOT
    / "presentations"
    / "ispring-course"
    / "module-01-stropovka-gruzov"
    / "live-preview"
    / "video-build"
    / "worker_standup_preview_2026-06-21"
)
GIF_PATH = OUT_DIR / "worker_standup_preview.gif"
LAST_FRAME_PATH = OUT_DIR / "worker_standup_preview_last_frame.png"
STAND_FRAME_PATH = OUT_DIR / "worker_standup_preview_stand_frame.png"
FRAMES_DIR = OUT_DIR / "frames"

FRAME_MS = 90


def ease_in_out(t: float) -> float:
    t = max(0.0, min(1.0, t))
    return 0.5 - 0.5 * math.cos(math.pi * t)


def transform_layer(
    image: Image.Image,
    scale: float = 1.0,
    rotate_deg: float = 0.0,
    opacity: float = 1.0,
) -> Image.Image:
    rgba = image.convert("RGBA")
    if scale != 1.0:
        new_size = (
            max(1, int(rgba.width * scale)),
            max(1, int(rgba.height * scale)),
        )
        rgba = rgba.resize(new_size, Image.Resampling.LANCZOS)
    if rotate_deg:
        rgba = rgba.rotate(rotate_deg, resample=Image.Resampling.BICUBIC, expand=True)
    if opacity < 1.0:
        alpha = rgba.getchannel("A")
        alpha = alpha.point(lambda px: int(px * opacity))
        rgba.putalpha(alpha)
    return rgba


def alpha_box_mask(size: tuple[int, int], box: tuple[int, int, int, int], blur: int = 6) -> Image.Image:
    mask = Image.new("L", size, 0)
    draw = ImageDraw.Draw(mask)
    draw.rounded_rectangle(box, radius=18, fill=255)
    if blur:
        mask = mask.filter(ImageFilter.GaussianBlur(blur))
    return mask


def polygon_mask(size: tuple[int, int], points: list[tuple[int, int]], blur: int = 0) -> Image.Image:
    mask = Image.new("L", size, 0)
    draw = ImageDraw.Draw(mask)
    draw.polygon(points, fill=255)
    if blur:
        mask = mask.filter(ImageFilter.GaussianBlur(blur))
    return mask


def extract_figure(
    region: Image.Image,
    white_threshold: int = 242,
    silhouette_points: list[tuple[int, int]] | None = None,
) -> Image.Image:
    rgba = region.convert("RGBA")
    mask = Image.new("L", rgba.size, 0)
    src = rgba.load()
    dst = mask.load()
    for y in range(rgba.height):
        for x in range(rgba.width):
            r, g, b, _ = src[x, y]
            if r < white_threshold or g < white_threshold or b < white_threshold:
                dst[x, y] = 255
    if silhouette_points:
        shape_mask = polygon_mask(rgba.size, silhouette_points, blur=1)
        mask = ImageChops.multiply(mask, shape_mask)
    mask = mask.filter(ImageFilter.GaussianBlur(1))
    rgba.putalpha(mask)
    return rgba


def paste_alpha(base: Image.Image, layer: Image.Image, position: tuple[int, int]) -> None:
    base.alpha_composite(layer, position)


def build_base_clean(source: Image.Image) -> Image.Image:
    clean = source.convert("RGBA")

    # Softly remove the bent worker while keeping the crate and left sling area readable.
    worker_mask = Image.new("L", clean.size, 0)
    draw = ImageDraw.Draw(worker_mask)
    draw.polygon(
        [(148, 196), (235, 196), (322, 254), (344, 464), (202, 480), (138, 320)],
        fill=255,
    )
    worker_mask = worker_mask.filter(ImageFilter.GaussianBlur(5))
    white_fill = Image.new("RGBA", clean.size, (255, 255, 255, 255))
    clean = Image.composite(white_fill, clean, worker_mask)

    # Restore the floor line where the original figure was removed.
    floor = ImageDraw.Draw(clean)
    # Clean residual fragments from the removed bent pose.
    floor.ellipse((142, 248, 171, 287), fill=(255, 255, 255, 255))
    floor.polygon(
        [(132, 332), (170, 317), (176, 347), (166, 380), (132, 385), (120, 358)],
        fill=(255, 255, 255, 255),
    )
    floor.rectangle((206, 462, 309, 490), fill=(255, 255, 255, 255))

    floor.line((34, 484, 328, 484), fill=(40, 40, 40, 255), width=2)
    floor.line((83, 188, 153, 387), fill=(63, 109, 120, 255), width=3)
    floor.line((82, 188, 152, 387), fill=(104, 156, 166, 255), width=1)
    floor.line((34, 321, 157, 321), fill=(56, 56, 56, 255), width=2)
    floor.line((34, 387, 157, 387), fill=(56, 56, 56, 255), width=2)
    floor.line((157, 321, 157, 387), fill=(56, 56, 56, 255), width=2)
    return clean


def build_target_frame(source: Image.Image, base_clean: Image.Image, stand: Image.Image) -> Image.Image:
    frame = base_clean.copy()
    stand_layer = transform_layer(stand, scale=0.87, rotate_deg=0.0, opacity=1.0)
    frame.alpha_composite(stand_layer, (152, 212))
    draw = ImageDraw.Draw(frame)
    draw.ellipse((196, 458, 276, 480), fill=(0, 0, 0, 45))
    return frame


def build_frames() -> list[Image.Image]:
    source = Image.open(SOURCE).convert("RGBA")
    base_clean = build_base_clean(source)

    left_bent_box = (130, 198, 342, 480)
    right_stand_box = (500, 178, 648, 482)

    bent = extract_figure(source.crop(left_bent_box))
    stand = extract_figure(
        source.crop(right_stand_box),
        silhouette_points=[
            (40, 36),
            (69, 34),
            (94, 43),
            (118, 78),
            (128, 132),
            (126, 206),
            (118, 284),
            (88, 302),
            (55, 300),
            (46, 258),
            (40, 208),
            (8, 176),
            (0, 156),
            (8, 141),
            (24, 142),
            (44, 132),
            (49, 107),
            (50, 82),
        ],
    )

    stand_mask_box = alpha_box_mask(source.size, (146, 186, 286, 486), blur=8)
    bent_mask_box = alpha_box_mask(source.size, (126, 194, 344, 485), blur=8)
    target = build_target_frame(source, base_clean, stand)

    frames: list[Image.Image] = []

    def compose(progress: float) -> Image.Image:
        p = ease_in_out(progress)
        frame = Image.blend(source, target, p)

        # Fade the original bent worker out while lifting it slightly.
        bent_layer = transform_layer(
            bent,
            scale=1.0 - 0.06 * p,
            rotate_deg=-6 * p,
            opacity=max(0.0, 1.0 - p * 1.2),
        )
        bent_pos = (
            int(left_bent_box[0] + 6 * p),
            int(left_bent_box[1] - 18 * p),
        )

        # Bring the standing pose in with only a small positional drift.
        stand_layer = transform_layer(
            stand,
            scale=0.83 + 0.04 * p,
            rotate_deg=0.8 * (1.0 - p),
            opacity=max(0.0, min(1.0, (p - 0.18) / 0.82)),
        )
        stand_pos = (
            int(156 - 10 * (1.0 - p)),
            int(216 + 10 * (1.0 - p)),
        )

        bent_canvas = Image.new("RGBA", source.size, (0, 0, 0, 0))
        stand_canvas = Image.new("RGBA", source.size, (0, 0, 0, 0))
        paste_alpha(bent_canvas, bent_layer, bent_pos)
        paste_alpha(stand_canvas, stand_layer, stand_pos)

        bent_canvas.putalpha(ImageChops.multiply(bent_canvas.getchannel("A"), bent_mask_box))
        stand_canvas.putalpha(ImageChops.multiply(stand_canvas.getchannel("A"), stand_mask_box))

        frame.alpha_composite(bent_canvas)
        frame.alpha_composite(stand_canvas)

        shadow = ImageDraw.Draw(frame)
        shadow_alpha = int(45 * p)
        shadow.ellipse((196, 458, 276, 480), fill=(0, 0, 0, shadow_alpha))
        return frame

    for idx in range(14):
        frames.append(compose(idx / 13))
    for _ in range(8):
        frames.append(compose(1.0))
    for idx in range(14):
        frames.append(compose(1.0 - (idx / 13)))
    for _ in range(5):
        frames.append(compose(0.0))

    return frames


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    FRAMES_DIR.mkdir(parents=True, exist_ok=True)
    frames = build_frames()
    for old_frame in FRAMES_DIR.glob("frame_*.png"):
        old_frame.unlink()
    for idx, frame in enumerate(frames):
        frame.save(FRAMES_DIR / f"frame_{idx:04d}.png")
    frames[0].save(
        GIF_PATH,
        save_all=True,
        append_images=frames[1:],
        duration=FRAME_MS,
        loop=0,
        disposal=2,
    )
    frames[-1].save(LAST_FRAME_PATH)
    frames[13].save(STAND_FRAME_PATH)
    print(GIF_PATH)
    print(LAST_FRAME_PATH)
    print(STAND_FRAME_PATH)
    print(FRAMES_DIR)


if __name__ == "__main__":
    main()
