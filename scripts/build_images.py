"""Generate deck illustrations for one language into assets/<lang>/.
Usage: python scripts/build_images.py [en|hi|gu]
Shapes are drawn with matplotlib; all words are drawn afterwards with Pillow + raqm,
so Hindi and Gujarati letters are shaped correctly."""
import os, random, sys
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle, Circle, FancyBboxPatch
from PIL import Image, ImageDraw, ImageFont
sys.path.insert(0, os.path.dirname(__file__))
from strings import tr

LANG = sys.argv[1] if len(sys.argv) > 1 else "en"
T = lambda s: tr(s, LANG)
OUT = f"assets/{LANG}/"; os.makedirs(OUT, exist_ok=True)
G_DARK, G_MID, GOLD, SAND, ROAD = "#0F3D2E", "#2E7D5B", "#D4A017", "#F3E9D2", "#5A5A5A"
FONTDIR = "/usr/share/fonts/truetype/noto/"
FONTS = {"en": "NotoSans", "hi": "NotoSansDevanagari", "gu": "NotoSansGujarati"}
plt.rcParams["font.family"] = "DejaVu Sans"
random.seed(7)
SCALE = 1.4  # labels a bit larger than matplotlib defaults, so they read well on slides

def save(fig, ax, labels, name):
    """labels: (x, y, text, size_pt, color, bold, anchor) in data coords. anchor is a Pillow anchor like 'mm'."""
    fig.canvas.draw()
    path = OUT + name
    fig.savefig(path, dpi=fig.dpi, transparent=True)
    H = fig.bbox.height
    pts = [(ax.transData.transform((x, y)), t, sz, c, b, a) for (x, y, t, sz, c, b, a) in labels]
    plt.close(fig)
    im = Image.open(path).convert("RGBA"); d = ImageDraw.Draw(im)
    for (px, py), t, sz, c, b, a in pts:
        f = ImageFont.truetype(FONTDIR + FONTS[LANG] + ("-Bold" if b else "-Regular") + ".ttf",
                               int(sz * SCALE * fig.dpi / 72), layout_engine=ImageFont.Layout.RAQM)
        lines = t.split("\n"); lh = int(sz * SCALE * fig.dpi / 72 * 1.25)
        y0 = H - py - (len(lines) - 1) * lh / 2
        for i, ln in enumerate(lines):
            d.text((px, y0 + i * lh), ln, font=f, fill=c, anchor=a)
    bb = im.getbbox(); im = im.crop((max(bb[0]-10, 0), max(bb[1]-10, 0), min(bb[2]+10, im.width), min(bb[3]+10, im.height)))
    im.save(path)

def tree(ax, x, y, r):
    ax.add_patch(Circle((x, y), r, color=random.choice(["#2E7D32", "#388E3C", "#43A047", "#1B5E20"]), zorder=5))
    ax.add_patch(Circle((x - r*0.3, y + r*0.3), r*0.45, color="#66BB6A", alpha=0.6, zorder=6))

# 1. Aerial site layout (no words; same for all languages)
fig = plt.figure(figsize=(16, 9), dpi=120); ax = fig.add_axes([0, 0, 1, 1])
ax.set_xlim(0, 160); ax.set_ylim(0, 90); ax.axis("off")
ax.add_patch(Rectangle((0, 0), 160, 90, color="#8BC34A")); ax.add_patch(Rectangle((4, 4), 152, 82, color="#DCE8C8"))
for y in (4, 30, 56, 82):
    ax.add_patch(Rectangle((4, y), 152, 4, color=ROAD, zorder=2))
    for x in range(6, 156, 6): ax.add_patch(Rectangle((x, y + 1.85), 3, 0.3, color="white", zorder=3))
for x in (4, 80, 152): ax.add_patch(Rectangle((x, 4), 4, 82, color=ROAD, zorder=2))
cols = [SAND, "#EFE0BF", "#F7EEDB"]
for (y0, y1) in ((8, 30), (34, 56), (60, 82)):
    for (x0, x1) in ((8, 80), (84, 152)):
        w = (x1 - x0) / 8; h = (y1 - y0) / 2
        for i in range(8):
            for j in range(2):
                cx = x0 + i*w + w/2
                if y0 == 34 and 56 < cx < 104: continue
                ax.add_patch(Rectangle((x0 + i*w + 0.3, y0 + j*h + 0.3), w - 0.6, h - 0.6, facecolor=cols[(i + j) % 3], edgecolor="#B08D57", lw=1.2, zorder=3))
                if random.random() < 0.35:
                    ax.add_patch(Rectangle((x0 + i*w + w*0.25, y0 + j*h + h*0.25), w*0.5, h*0.5, facecolor="#B71C1C" if random.random() < .5 else "#6D4C41", zorder=4, alpha=0.85))
ax.add_patch(FancyBboxPatch((58, 35), 44, 20, boxstyle="round,pad=0.5", color="#7CB342", zorder=4))
for _ in range(22): tree(ax, random.uniform(61, 99), random.uniform(37.5, 52.5), random.uniform(1.2, 2.0))
for x in range(2, 160, 5): tree(ax, x, 1.5, 1.4); tree(ax, x + 2, 88.5, 1.4)
fig.savefig(OUT + "site_layout.png", dpi=120); plt.close(fig)

# 2. Single plot
fig = plt.figure(figsize=(8, 9), dpi=150); ax = fig.add_axes([0, 0, 1, 1])
ax.set_xlim(-12, 52); ax.set_ylim(-14, 58); ax.set_aspect("equal"); ax.axis("off")
ax.add_patch(Rectangle((-12, -14), 64, 8, color=ROAD))
ax.add_patch(Rectangle((0, 0), 30, 45, facecolor=SAND, edgecolor=G_DARK, lw=4))
ax.add_patch(Rectangle((5, 12), 20, 26, facecolor="#E8D3A8", edgecolor="#B08D57", lw=2, ls="--"))
ax.annotate("", xy=(0, 49), xytext=(30, 49), arrowprops=dict(arrowstyle="<->", lw=2.5, color=GOLD))
ax.annotate("", xy=(35, 0), xytext=(35, 45), arrowprops=dict(arrowstyle="<->", lw=2.5, color=GOLD))
tree(ax, -6, 10, 3); tree(ax, -6, 30, 3); tree(ax, 46, 15, 3); tree(ax, 46, 38, 3)
save(fig, ax, [(20, -10, T("ROAD"), 14, "white", True, "mm"), (15, 25, T("Your\nhome"), 18, "#6D4C41", True, "mm"),
               (15, 5.5, T("150 sq yd"), 20, G_DARK, True, "mm"), (15, 52, T("30 ft"), 16, G_DARK, True, "mm"),
               (36.5, 22.5, T("45 ft"), 16, G_DARK, True, "lm")], "single_plot.png")

# 3. R1 ladder (percent)
slabs = [("1 plot", 5.0), ("2 plots", 6.0), ("3-4 plots", 6.5), ("5 plots", 7.5), ("6+ plots", 8.0)]
fig = plt.figure(figsize=(10, 6), dpi=150); ax = fig.add_axes([0, 0, 1, 1])
colors = ["#A8D5BA", "#7CC49A", "#4FA97A", "#2E7D5B", GOLD]; L = []
for i, (lab, r) in enumerate(slabs):
    ax.add_patch(Rectangle((i, 0), 0.92, r, color=colors[i]))
    L += [(i + 0.46, r + 0.55, f"{r:g}%", 22, G_DARK, True, "mm"), (i + 0.46, -0.55, T(lab), 14, G_DARK, False, "mm")]
L.append((4.46, 4, T("MAX"), 16, "white", True, "mm"))
ax.set_xlim(-0.2, 5.1); ax.set_ylim(-1.2, 9.4); ax.axis("off")
save(fig, ax, L, "ladder.png")

# 4/5. Earnings charts (numbers only; axis titles go on the slide)
def earnings(rate, name, top, ymax):
    n = list(range(1, 11)); e = [k * 150 * rate(k) / 1e5 for k in n]
    fig, ax = plt.subplots(figsize=(10, 6), dpi=150); ax.set_facecolor("none")
    bars = ax.bar(n, e, width=0.7, color=[top(k) for k in n])
    L = [(b.get_x() + b.get_width()/2, v + ymax * 0.035, f"{v:.2f}".rstrip("0").rstrip("."), 13, G_DARK, True, "mm") for b, v in zip(bars, e)]
    ax.set_xticks(n)
    for s in ("top", "right"): ax.spines[s].set_visible(False)
    ax.tick_params(labelsize=13, colors=G_DARK); ax.set_ylim(0, ymax)
    save(fig, ax, L, name)
r1 = lambda k: (0.05 if k == 1 else 0.06 if k == 2 else 0.065 if k <= 4 else 0.075 if k == 5 else 0.08) * 20000
earnings(r1, "earnings.png", lambda k: GOLD if k >= 6 else G_MID, 27)
r2 = lambda k: 1700 if k == 1 else 1800 if k <= 4 else 2000
earnings(r2, "r2_earnings.png", lambda k: "#7CC49A" if k == 1 else G_MID if k <= 4 else GOLD, 33)

# 6. EMI timeline
fig = plt.figure(figsize=(12, 3.2), dpi=150); ax = fig.add_axes([0, 0, 1, 1])
ax.set_xlim(-0.8, 6.4); ax.set_ylim(-1.0, 0.9); ax.set_aspect("equal"); ax.axis("off")
ax.plot([0, 5.6], [0, 0], color=G_MID, lw=6, zorder=1); L = []
for i in range(5):
    x = i * 1.4
    ax.add_patch(Circle((x, 0), 0.28, color=GOLD if i == 0 else G_DARK, zorder=2))
    if i: L.append((x, 0, str(i), 16, "white", True, "mm"))
    L.append((x, 0.55, T("Book") if i == 0 else T("Month {n}").format(n=i), 15, G_DARK, True, "mm"))
    L.append((x, -0.6, T("Today") if i == 0 else "₹7,87,500" if LANG != "en" else "Rs 7,87,500", 14, G_DARK, False, "mm"))
save(fig, ax, L, "emi_timeline.png")

# 7. R2 ladder (INR per sq yd)
fig = plt.figure(figsize=(10, 6), dpi=150); ax = fig.add_axes([0, 0, 1, 1]); L = []
for i, (lab, v, c) in enumerate([("1 plot", 1700, "#7CC49A"), ("2 to 4 plots", 1800, G_MID), ("5+ plots", 2000, GOLD)]):
    h = 2 + i * 1.6
    ax.add_patch(Rectangle((i * 3, 0), 2.8, h, color=c))
    L += [(i * 3 + 1.4, h + 0.55, ("Rs " if LANG == "en" else "₹") + f"{v:,}", 26, G_DARK, True, "mm"),
          (i * 3 + 1.4, h - 0.55, T("per sq yd"), 15, "white", True, "mm"),
          (i * 3 + 1.4, -0.55, T(lab), 17, G_DARK, True, "mm")]
L.append((7.4, 1.6, T("TOP\nSLAB"), 15, "white", True, "mm"))
ax.set_xlim(-0.2, 9); ax.set_ylim(-1.1, 7.6); ax.axis("off")
save(fig, ax, L, "r2_ladder.png")
print("images done", LANG)
