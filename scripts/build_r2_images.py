"""Images for the R2 sales deck."""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle, FancyArrowPatch
OUT = "assets/"
G_DARK, G_MID, GOLD = "#0F3D2E", "#2E7D5B", "#D4A017"
plt.rcParams["font.family"] = "DejaVu Sans"

# Staircase of R2 slabs
fig, ax = plt.subplots(figsize=(10, 6), dpi=150); fig.patch.set_alpha(0)
steps = [("1 plot", 1700, "#7CC49A"), ("2 to 4 plots", 1800, G_MID), ("5+ plots", 2000, GOLD)]
for i, (lab, v, c) in enumerate(steps):
    h = 2 + i * 1.6
    ax.add_patch(Rectangle((i * 3, 0), 2.8, h, color=c))
    ax.text(i * 3 + 1.4, h + 0.35, f"Rs {v:,}", ha="center", fontsize=26, fontweight="bold", color=G_DARK)
    ax.text(i * 3 + 1.4, h - 0.55, "per sq yd", ha="center", va="center", fontsize=15, color="white", fontweight="bold")
    ax.text(i * 3 + 1.4, -0.55, lab, ha="center", fontsize=17, color=G_DARK, fontweight="bold")
ax.text(7.4, 1.6, "TOP\nSLAB", ha="center", va="center", fontsize=15, fontweight="bold", color="white")
ax.set_xlim(-0.2, 9); ax.set_ylim(-1.1, 7.6); ax.axis("off")
fig.savefig(OUT + "r2_ladder.png", bbox_inches="tight", transparent=True); plt.close(fig)

# Earnings chart 1-10 plots
rate = lambda k: 1700 if k == 1 else 1800 if k <= 4 else 2000
n = list(range(1, 11)); e = [k * 150 * rate(k) / 1e5 for k in n]
fig, ax = plt.subplots(figsize=(10, 6), dpi=150); fig.patch.set_alpha(0); ax.set_facecolor("none")
bars = ax.bar(n, e, width=0.7, color=["#7CC49A" if k == 1 else G_MID if k <= 4 else GOLD for k in n])
for b, v in zip(bars, e):
    ax.text(b.get_x() + b.get_width()/2, v + 0.5, f"{v:g}", ha="center", fontsize=13, fontweight="bold", color=G_DARK)
ax.set_xticks(n); ax.set_xlabel("Plots you sell", fontsize=14, color=G_DARK); ax.set_ylabel("You earn (Rs lakh)", fontsize=14, color=G_DARK)
for s in ("top", "right"): ax.spines[s].set_visible(False)
ax.tick_params(labelsize=13, colors=G_DARK); ax.set_ylim(0, 33)
fig.savefig(OUT + "r2_earnings.png", bbox_inches="tight", transparent=True); plt.close(fig)
print("r2 images done")
