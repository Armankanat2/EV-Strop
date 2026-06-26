from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw

import build_angle_animatic_2026_06_26 as build


ROOT = Path(__file__).resolve().parents[5]
OUT_DIR = ROOT / "assets" / "course-media" / "module-01-stropovka-gruzov" / "images" / "hook-saddle-preview-v1"


def stop_frame_time(stage: build.Stage) -> float:
    return stage.hold_start + (stage.hold_end - stage.hold_start) * 0.5


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    cues = build.parse_vtt(build.VTT_FILE)
    stage = next(item for item in build.STAGES if item.angle == 60)
    frame = build.render_frame(stop_frame_time(stage), cues)
    full_path = OUT_DIR / "hook-saddle-preview_full-ui_60deg_v1_2026-06-26.png"
    frame.save(full_path)

    crop = frame.crop((420, 120, 860, 420))
    crop_path = OUT_DIR / "hook-saddle-preview_crop_60deg_v1_2026-06-26.png"
    crop.save(crop_path)

    board = Image.new("RGBA", (980, 430), (244, 240, 231, 255))
    draw = ImageDraw.Draw(board, "RGBA")
    draw.text((28, 22), "Hook + Saddle Preview", font=build.get_font(28, bold=True), fill=(48, 56, 67, 255))
    draw.text((30, 58), "Reference check of the sling seating on the hook saddle", font=build.get_font(16), fill=(92, 101, 114, 255))
    board.alpha_composite(crop.resize((620, 422), Image.Resampling.LANCZOS), (330, 4))
    draw.rounded_rectangle((330, 4, 950, 426), radius=18, outline=(223, 212, 194, 255), width=2)
    draw.rounded_rectangle((24, 106, 278, 340), radius=18, fill=(255, 251, 244, 255), outline=(223, 212, 194, 255), width=2)
    notes = [
        "1. Строп лежит",
        "   на седле крюка.",
        "2. Ветви выходят",
        "   из зоны контакта.",
        "3. Узел не висит",
        "   отдельно от крюка.",
    ]
    y = 136
    for line in notes:
        draw.text((46, y), line, font=build.get_font(18), fill=(48, 56, 67, 255))
        y += 30

    board_path = OUT_DIR / "hook-saddle-preview_board_v1_2026-06-26.png"
    board.save(board_path)

    print(full_path)
    print(crop_path)
    print(board_path)


if __name__ == "__main__":
    main()
