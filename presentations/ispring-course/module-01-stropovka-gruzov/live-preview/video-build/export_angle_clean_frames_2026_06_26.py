from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw

import build_angle_animatic_2026_06_26 as build


ROOT = Path(__file__).resolve().parents[5]
OUT_DIR = ROOT / "assets" / "course-media" / "module-01-stropovka-gruzov" / "images" / "angle-animation-clean-frames-v9"
PANEL_BOX = build.PANEL_BOX


def get_font(size: int, bold: bool = False):
    return build.get_font(size, bold=bold)


def stop_frame_time(stage: build.Stage) -> float:
    return stage.hold_start + (stage.hold_end - stage.hold_start) * 0.5


def save_ui_and_panel_frames(cues: list[build.Cue]) -> list[tuple[str, Path, Path]]:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    exported: list[tuple[str, Path, Path]] = []

    for stage in build.STAGES:
        t = stop_frame_time(stage)
        frame = build.render_frame(t, cues)
        ui_path = OUT_DIR / f"angle-{stage.angle:03d}_full-ui_v9_2026-06-26.png"
        panel_path = OUT_DIR / f"angle-{stage.angle:03d}_panel-only_v9_2026-06-26.png"
        frame.save(ui_path)
        panel = frame.crop(PANEL_BOX)
        panel.save(panel_path)
        exported.append((f"{stage.angle}°", ui_path, panel_path))

    return exported


def make_contact_sheet(items: list[tuple[str, Path, Path]]) -> Path:
    cards = []
    for label, ui_path, _panel_path in items:
        image = Image.open(ui_path).convert("RGBA")
        thumb = image.resize((480, 270), Image.Resampling.LANCZOS)
        cards.append((label, thumb))

    cols = 2
    rows = 3
    pad = 28
    title_h = 90
    card_w = 480
    card_h = 320
    width = pad + cols * card_w + (cols - 1) * pad + pad
    height = title_h + pad + rows * card_h + (rows - 1) * pad + pad

    sheet = Image.new("RGBA", (width, height), (244, 240, 231, 255))
    draw = ImageDraw.Draw(sheet, "RGBA")
    draw.text((28, 24), "Clean Stop Frames V9", font=get_font(34, bold=True), fill=(48, 56, 67, 255))
    draw.text((30, 58), "30°, 45°, 60°, 90°, 100°", font=get_font(18), fill=(92, 101, 114, 255))

    for idx, (label, thumb) in enumerate(cards):
        row = idx // cols
        col = idx % cols
        x = pad + col * (card_w + pad)
        y = title_h + pad + row * (card_h + pad)
        draw.rounded_rectangle((x, y, x + card_w, y + card_h), radius=22, fill=(255, 251, 244, 255), outline=(223, 212, 194, 255), width=2)
        sheet.alpha_composite(thumb, (x, y))
        label_w = draw.textbbox((0, 0), label, font=get_font(24, bold=True))[2]
        draw.rounded_rectangle((x + 18, y + 18, x + 18 + label_w + 22, y + 58), radius=18, fill=(17, 26, 37, 215))
        draw.text((x + 30, y + 26), label, font=get_font(24, bold=True), fill=(255, 255, 255, 255))

    out_path = OUT_DIR / "angle-clean-frames-contact-sheet_v9_2026-06-26.png"
    sheet.save(out_path)
    return out_path


def write_index(items: list[tuple[str, Path, Path]], contact_sheet: Path) -> Path:
    lines = [
        "# Clean Frames V9",
        "",
        f"contact_sheet={contact_sheet}",
        "",
    ]
    for label, ui_path, panel_path in items:
        lines.append(f"{label}_ui={ui_path}")
        lines.append(f"{label}_panel={panel_path}")
    out_path = OUT_DIR / "angle-clean-frames-index_v9_2026-06-26.txt"
    out_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return out_path


def main() -> None:
    cues = build.parse_vtt(build.VTT_FILE)
    items = save_ui_and_panel_frames(cues)
    contact_sheet = make_contact_sheet(items)
    index_path = write_index(items, contact_sheet)
    print(OUT_DIR)
    print(contact_sheet)
    print(index_path)


if __name__ == "__main__":
    main()
