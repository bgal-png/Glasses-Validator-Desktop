"""Generate app_icon.ico — a pair of glasses with a green check mark.

Drawn at high resolution and downsampled so the small sizes stay smooth.
Re-run only when the icon should change:  python make_icon.py
"""
from PIL import Image, ImageDraw

S = 1024                      # working canvas
FRAME = (43, 90, 158, 255)    # slate blue
CHECK = (34, 160, 74, 255)    # green
LENS = (120, 170, 225, 70)    # faint glass tint


def rounded(draw, box, r, **kw):
    draw.rounded_rectangle(box, radius=r, **kw)


def build(size=S):
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    u = size / 1024.0                       # scale helper
    w = int(58 * u)                         # stroke width

    # ---- glasses ----
    top, h = int(340 * u), int(310 * u)
    lens_w = int(350 * u)
    inset = int(96 * u)                     # room for the temple stubs
    left = (inset, top, inset + lens_w, top + h)
    right = (size - inset - lens_w, top, size - inset, top + h)
    for box in (left, right):
        rounded(d, box, int(85 * u), fill=LENS, outline=FRAME, width=w)

    # bridge
    d.arc((left[2] - int(20 * u), top + int(30 * u),
           right[0] + int(20 * u), top + int(180 * u)),
          start=200, end=340, fill=FRAME, width=w)

    # temple stubs — short, near-horizontal, hugging the frame
    ty = top + int(70 * u)
    d.line((left[0], ty, int(10 * u), ty - int(34 * u)), fill=FRAME, width=w)
    d.line((right[2], ty, size - int(10 * u), ty - int(34 * u)), fill=FRAME, width=w)

    # ---- check mark (bottom-right, on its own disc so it reads at 16px) ----
    cx, cy, r = int(720 * u), int(730 * u), int(250 * u)
    d.ellipse((cx - r, cy - r, cx + r, cy + r), fill=CHECK)
    cw = int(78 * u)
    d.line([(cx - int(120 * u), cy + int(10 * u)),
            (cx - int(30 * u), cy + int(105 * u)),
            (cx + int(130 * u), cy - int(105 * u))],
           fill=(255, 255, 255, 255), width=cw, joint="curve")
    # round the stroke ends
    for pt in ((cx - int(120 * u), cy + int(10 * u)),
               (cx + int(130 * u), cy - int(105 * u))):
        d.ellipse((pt[0] - cw // 2, pt[1] - cw // 2,
                   pt[0] + cw // 2, pt[1] + cw // 2), fill=(255, 255, 255, 255))
    return img


def main():
    base = build()
    sizes = [16, 24, 32, 48, 64, 128, 256]
    frames = [base.resize((s, s), Image.LANCZOS) for s in sizes]
    frames[-1].save("app_icon.ico", format="ICO",
                    sizes=[(s, s) for s in sizes])
    base.resize((256, 256), Image.LANCZOS).save("app_icon.png")
    print("wrote app_icon.ico and app_icon.png")


if __name__ == "__main__":
    main()
