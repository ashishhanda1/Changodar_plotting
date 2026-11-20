"""Generate illustrations for the marketing deck into assets/."""
import random
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle, Circle, FancyBboxPatch, Polygon

OUT = "assets/"
G_DARK, G_MID, G_LIGHT, GOLD, SAND, ROAD = "#0F3D2E", "#2E7D5B", "#A8D5BA", "#D4A017", "#F3E9D2", "#5A5A5A"
plt.rcParams["font.family"] = "DejaVu Sans"
random.seed(7)

def tree(ax, x, y, r):
    ax.add_patch(Circle((x, y), r, color=random.choice(["#2E7D32", "#388E3C", "#43A047", "#1B5E20"]), zorder=5))
    ax.add_patch(Circle((x - r*0.3, y + r*0.3), r*0.45, color="#66BB6A", alpha=0.6, zorder=6))

# 1. Aerial site layout (hero)
fig, ax = plt.subplots(figsize=(16, 9), dpi=120)
ax.set_xlim(0, 160); ax.set_ylim(0, 90); ax.axis("off")
ax.add_patch(Rectangle((0, 0), 160, 90, color="#8BC34A"))
ax.add_patch(Rectangle((4, 4), 152, 82, color="#DCE8C8"))
# roads
for y in (4, 30, 56, 82):
    ax.add_patch(Rectangle((4, y), 152, 4, color=ROAD, zorder=2))
    for x in range(6, 156, 6):
        ax.add_patch(Rectangle((x, y + 1.85), 3, 0.3, color="white", zorder=3))
for x in (4, 80, 152):
    ax.add_patch(Rectangle((x, 4), 4, 82, color=ROAD, zorder=2))
# plots
cols = [SAND, "#EFE0BF", "#F7EEDB"]
for (y0, y1) in ((8, 30), (34, 56), (60, 82)):
    for (x0, x1) in ((8, 80), (84, 152)):
        w = (x1 - x0) / 8; h = (y1 - y0) / 2
        for i in range(8):
            for j in range(2):
                cx = x0 + i*w + w/2
                if y0 == 34 and 56 < cx < 104:
                    continue  # park area
                ax.add_patch(Rectangle((x0 + i*w + 0.3, y0 + j*h + 0.3), w - 0.6, h - 0.6,
                                       facecolor=cols[(i + j) % 3], edgecolor="#B08D57", lw=1.2, zorder=3))
                if random.random() < 0.35:
                    ax.add_patch(Rectangle((x0 + i*w + w*0.25, y0 + j*h + h*0.25), w*0.5, h*0.5,
                                           facecolor="#B71C1C" if random.random() < .5 else "#6D4C41", zorder=4, alpha=0.85))
# park
ax.add_patch(FancyBboxPatch((58, 35), 44, 20, boxstyle="round,pad=0.5", color="#7CB342", zorder=4))
for _ in range(14):
    tree(ax, random.uniform(62, 98), random.uniform(39, 51), random.uniform(1.2, 2.0))
ax.text(80, 45, "PARK", ha="center", va="center", fontsize=22, color="white", fontweight="bold", zorder=7)
for x in range(2, 160, 5):
    tree(ax, x, 1.5, 1.4); tree(ax, x + 2, 88.5, 1.4)
fig.savefig(OUT + "site_layout.png", bbox_inches="tight", pad_inches=0); plt.close(fig)

# 2. Single plot 30 ft x 45 ft
fig, ax = plt.subplots(figsize=(8, 8), dpi=150)
ax.set_xlim(-12, 52); ax.set_ylim(-14, 58); ax.set_aspect("equal"); ax.axis("off")
ax.add_patch(Rectangle((-12, -14), 64, 8, color=ROAD))
ax.text(20, -10, "ROAD", ha="center", va="center", color="white", fontsize=14, fontweight="bold")
ax.add_patch(Rectangle((0, 0), 30, 45, facecolor=SAND, edgecolor=G_DARK, lw=4))
ax.add_patch(Rectangle((5, 12), 20, 26, facecolor="#E8D3A8", edgecolor="#B08D57", lw=2, ls="--"))
ax.text(15, 25, "Your\nhome", ha="center", va="center", fontsize=18, color="#6D4C41", fontweight="bold")
ax.text(15, 5.5, "150 sq yd", ha="center", va="center", fontsize=22, color=G_DARK, fontweight="bold")
ax.annotate("", xy=(0, 49), xytext=(30, 49), arrowprops=dict(arrowstyle="<->", lw=2.5, color=GOLD))
ax.text(15, 51.5, "30 ft", ha="center", fontsize=16, color=G_DARK, fontweight="bold")
ax.annotate("", xy=(35, 0), xytext=(35, 45), arrowprops=dict(arrowstyle="<->", lw=2.5, color=GOLD))
ax.text(37.5, 22.5, "45 ft", va="center", rotation=90, fontsize=16, color=G_DARK, fontweight="bold")
tree(ax, -6, 10, 3); tree(ax, -6, 30, 3); tree(ax, 46, 15, 3); tree(ax, 46, 38, 3)
fig.savefig(OUT + "single_plot.png", bbox_inches="tight", transparent=True); plt.close(fig)

# 3. Commission ladder staircase
slabs = [("1 plot", 5.0), ("2 plots", 6.0), ("3-4 plots", 6.5), ("5 plots", 7.5), ("6+ plots", 8.0)]
fig, ax = plt.subplots(figsize=(10, 6), dpi=150)
fig.patch.set_alpha(0); ax.set_facecolor("none")
colors = ["#A8D5BA", "#7CC49A", "#4FA97A", "#2E7D5B", GOLD]
for i, (lab, r) in enumerate(slabs):
    ax.add_patch(Rectangle((i, 0), 0.92, r, color=colors[i]))
    ax.text(i + 0.46, r + 0.25, f"{r:g}%", ha="center", fontsize=22, fontweight="bold", color=G_DARK)
    ax.text(i + 0.46, -0.6, lab, ha="center", fontsize=14, color=G_DARK)
ax.text(4.46, 4, "MAX", ha="center", fontsize=16, fontweight="bold", color="white", rotation=90)
ax.set_xlim(-0.2, 5.1); ax.set_ylim(-1.2, 9.4); ax.axis("off")
fig.savefig(OUT + "ladder.png", bbox_inches="tight", transparent=True); plt.close(fig)

# 4. Earnings by plots sold
n = list(range(1, 11))
rate = lambda k: 0.05 if k == 1 else 0.06 if k == 2 else 0.065 if k <= 4 else 0.075 if k == 5 else 0.08
earn = [k * 3000000 * rate(k) / 1e5 for k in n]
fig, ax = plt.subplots(figsize=(10, 6), dpi=150)
fig.patch.set_alpha(0); ax.set_facecolor("none")
bars = ax.bar(n, earn, color=[GOLD if k >= 6 else G_MID for k in n], width=0.7)
for b, e in zip(bars, earn):
    ax.text(b.get_x() + b.get_width()/2, e + 0.4, f"{e:.2f}".rstrip("0").rstrip("."), ha="center", fontsize=12, fontweight="bold", color=G_DARK)
ax.set_xticks(n); ax.set_xlabel("Plots you sell", fontsize=13, color=G_DARK)
ax.set_ylabel("You earn (Rs lakh)", fontsize=13, color=G_DARK)
for s in ("top", "right"): ax.spines[s].set_visible(False)
ax.tick_params(labelsize=12, colors=G_DARK); ax.set_ylim(0, 27)
fig.savefig(OUT + "earnings.png", bbox_inches="tight", transparent=True); plt.close(fig)

# 5. EMI timeline
fig, ax = plt.subplots(figsize=(12, 3.2), dpi=150)
fig.patch.set_alpha(0); ax.set_xlim(-0.8, 6.4); ax.set_ylim(-1.0, 0.9); ax.set_aspect("equal"); ax.axis("off")
ax.plot([0, 5.6], [0, 0], color=G_MID, lw=6, zorder=1)
labels = [("Book", "Today"), ("EMI 1", "Rs 7,87,500"), ("EMI 2", "Rs 7,87,500"), ("EMI 3", "Rs 7,87,500"), ("EMI 4", "Rs 7,87,500")]
for i, (a, b) in enumerate(labels):
    x = i * 1.4
    ax.add_patch(Circle((x, 0), 0.28, color=GOLD if i == 0 else G_DARK, zorder=2))
    ax.text(x, 0, str(i) if i else "", ha="center", va="center", color="white", fontsize=16, fontweight="bold", zorder=3)
    ax.text(x, 0.5, a if i == 0 else f"Month {i}", ha="center", fontsize=15, fontweight="bold", color=G_DARK)
    ax.text(x, -0.62, b, ha="center", fontsize=14, color=G_DARK)
fig.savefig(OUT + "emi_timeline.png", bbox_inches="tight", transparent=True); plt.close(fig)
print("images done")
