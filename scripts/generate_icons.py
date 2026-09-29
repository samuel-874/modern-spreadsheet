import os
import subprocess
from PIL import Image, ImageDraw, ImageFilter

def create_gridbook_icon(size=1024):
    scale = size / 1024.0

    # Create canvas
    canvas = Image.new("RGBA", (size, size), (0, 0, 0, 0))

    # Squircle geometry
    margin = int(92 * scale)
    card_w = size - 2 * margin
    card_h = size - 2 * margin
    corner_radius = int(188 * scale)
    squircle_box = [margin, margin, margin + card_w, margin + card_h]

    # Subtle ambient drop shadow for macOS dock depth
    shadow_offset = int(10 * scale)
    shadow_blur = int(18 * scale)
    shadow_mask = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    sdraw = ImageDraw.Draw(shadow_mask)
    shadow_box = [
        margin,
        margin + shadow_offset,
        margin + card_w,
        margin + card_h + shadow_offset
    ]
    sdraw.rounded_rectangle(shadow_box, radius=corner_radius, fill=(150, 48, 18, 90))
    shadow = shadow_mask.filter(ImageFilter.GaussianBlur(shadow_blur))
    canvas = Image.alpha_composite(canvas, shadow)

    # Base gradient image for squircle
    grad = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    gdraw = ImageDraw.Draw(grad)

    for y in range(size):
        ratio = y / size
        # #eb714c (235, 113, 76) -> #d65730 (214, 87, 48)
        r = int(235 * (1 - ratio) + 214 * ratio)
        g = int(113 * (1 - ratio) + 87 * ratio)
        b = int(76 * (1 - ratio) + 48 * ratio)
        gdraw.line([(0, y), (size, y)], fill=(r, g, b, 255))

    # Mask for squircle
    squircle_mask = Image.new("L", (size, size), 0)
    sq_draw = ImageDraw.Draw(squircle_mask)
    sq_draw.rounded_rectangle(squircle_box, radius=corner_radius, fill=255)

    # Composite gradient with squircle mask
    squircle_layer = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    squircle_layer.paste(grad, (0, 0), squircle_mask)

    # Inner subtle border highlight
    border_mask = Image.new("L", (size, size), 0)
    bdraw = ImageDraw.Draw(border_mask)
    bdraw.rounded_rectangle(
        squircle_box,
        radius=corner_radius,
        outline=255,
        width=int(3.5 * scale)
    )
    white_overlay = Image.new("RGBA", (size, size), (255, 255, 255, 50))
    squircle_layer.paste(white_overlay, (0, 0), border_mask)

    canvas = Image.alpha_composite(canvas, squircle_layer)

    # Vector Drawing Layer for Gridbook
    vector_layer = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    vdraw = ImageDraw.Draw(vector_layer)

    # EXACT MATHEMATICAL CENTERING:
    # Icon spans x from 2.5 to 20.5 (width 18.0 units, center X = 11.5)
    # Icon spans y from 2.5 to 21.5 (height 19.0 units, center Y = 12.0)
    # Target height: 535px at 1024x1024
    target_height = 535.0 * scale
    factor = target_height / 19.0
    ox = (size / 2.0) - (11.5 * factor) + (13.0 * scale)
    oy = (size / 2.0) - (12.0 * factor)

    def tx(x): return int(ox + x * factor)
    def ty(y): return int(oy + y * factor)

    sw_main = max(2, int(1.8 * factor))
    sw_thin = max(2, int(1.2 * factor))
    cover_r = int(2.5 * factor)

    # 1. Notebook outer cover
    cover_box = [tx(5.0), ty(2.5), tx(20.5), ty(21.5)]
    vdraw.rounded_rectangle(
        cover_box,
        radius=cover_r,
        outline=(255, 255, 255, 255),
        width=sw_main
    )

    # 2. Left spine margin divider
    vdraw.line(
        [(tx(8.5), ty(2.5)), (tx(8.5), ty(21.5))],
        fill=(255, 255, 255, 120),
        width=sw_thin
    )

    # 3. Spiral binding rings on left edge
    ring_ys = [6.5, 10.5, 14.5, 18.5]
    for ry in ring_ys:
        y_coord = ty(ry)
        x_start = tx(2.5)
        x_end = tx(6.25)
        vdraw.line(
            [(x_start, y_coord), (x_end, y_coord)],
            fill=(255, 255, 255, 255),
            width=sw_main
        )
        # Smooth rounded end caps
        cap_r = sw_main // 2
        vdraw.ellipse([x_start - cap_r, y_coord - cap_r, x_start + cap_r, y_coord + cap_r], fill=(255, 255, 255, 255))
        vdraw.ellipse([x_end - cap_r, y_coord - cap_r, x_end + cap_r, y_coord + cap_r], fill=(255, 255, 255, 255))

    # 4. Spreadsheet table header row highlight
    hdr_box = [tx(10.0), ty(6.0), tx(18.0), ty(9.5)]
    vdraw.rounded_rectangle(hdr_box, radius=max(2, int(0.6 * factor)), fill=(255, 255, 255, 60))

    # 5. Active spreadsheet cell (Cell A2) highlight
    act_box = [tx(10.0), ty(9.5), tx(14.0), ty(13.5)]
    vdraw.rounded_rectangle(act_box, radius=max(2, int(0.6 * factor)), fill=(255, 255, 255, 125))

    # 6. Grid columns line
    vdraw.line(
        [(tx(14.0), ty(6.0)), (tx(14.0), ty(17.5))],
        fill=(255, 255, 255, 220),
        width=sw_thin
    )

    # 7. Grid rows lines
    vdraw.line(
        [(tx(10.0), ty(9.5)), (tx(18.0), ty(9.5))],
        fill=(255, 255, 255, 220),
        width=sw_thin
    )
    vdraw.line(
        [(tx(10.0), ty(13.5)), (tx(18.0), ty(13.5))],
        fill=(255, 255, 255, 220),
        width=sw_thin
    )

    # Composite vector layer over squircle
    final_img = Image.alpha_composite(canvas, vector_layer)
    return final_img

def main():
    os.makedirs("build", exist_ok=True)
    os.makedirs("public", exist_ok=True)

    print("Generating 1024x1024 master icon with exact centering...")
    master = create_gridbook_icon(1024)

    # Save PNGs
    master.save("build/icon.png", format="PNG")
    master.resize((512, 512), Image.Resampling.LANCZOS).save("public/icon.png", format="PNG")
    master.resize((256, 256), Image.Resampling.LANCZOS).save("public/icon-256.png", format="PNG")

    # Generate Windows .ico with multi-resolutions
    print("Generating Windows icon.ico...")
    ico_sizes = [(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
    master.save("build/icon.ico", format="ICO", sizes=ico_sizes)
    master.save("public/icon.ico", format="ICO", sizes=ico_sizes)

    # Generate macOS .icns with iconutil
    print("Generating macOS icon.icns...")
    iconset_dir = "build/icon.iconset"
    os.makedirs(iconset_dir, exist_ok=True)

    mac_specs = [
        ("icon_16x16.png", 16),
        ("icon_16x16@2x.png", 32),
        ("icon_32x32.png", 32),
        ("icon_32x32@2x.png", 64),
        ("icon_128x128.png", 128),
        ("icon_128x128@2x.png", 256),
        ("icon_256x256.png", 256),
        ("icon_256x256@2x.png", 512),
        ("icon_512x512.png", 512),
        ("icon_512x512@2x.png", 1024),
    ]

    for fname, sz in mac_specs:
        resized = master.resize((sz, sz), Image.Resampling.LANCZOS)
        resized.save(os.path.join(iconset_dir, fname), format="PNG")

    try:
        subprocess.run(["iconutil", "-c", "icns", iconset_dir, "-o", "build/icon.icns"], check=True)
        print("Successfully created build/icon.icns!")
    except Exception as e:
        print("iconutil error:", e)

    # Clean up temporary iconset directory
    try:
        subprocess.run(["rm", "-rf", iconset_dir], check=True)
    except Exception:
        pass

    print("Icons successfully generated with exact centering!")

if __name__ == "__main__":
    main()
