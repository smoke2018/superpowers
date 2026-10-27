# -*- coding: utf-8 -*-
"""生成 PPT 所需函数图象素材"""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
from matplotlib import font_manager

# 中文字体
plt.rcParams["font.sans-serif"] = ["Noto Sans CJK SC", "Noto Sans CJK JP", "WenQuanYi Zen Hei"]
plt.rcParams["axes.unicode_minus"] = False

NAVY = "#1F3864"
ORANGE = "#E87722"
GRAY = "#8A8F98"

OUT = "/workspace/试讲-函数的极值/assets"


def style_axes(ax, xlim, ylim):
    ax.spines["left"].set_position(("data", 0))
    ax.spines["bottom"].set_position(("data", 0))
    ax.spines["right"].set_color("none")
    ax.spines["top"].set_color("none")
    ax.spines["left"].set_color("#333333")
    ax.spines["bottom"].set_color("#333333")
    ax.spines["left"].set_linewidth(1.2)
    ax.spines["bottom"].set_linewidth(1.2)
    ax.set_xlim(*xlim)
    ax.set_ylim(*ylim)
    ax.grid(True, ls=":", lw=0.7, color="#C9CDD4", alpha=0.8)
    ax.plot(1, 0, ">k", transform=ax.get_yaxis_transform(), clip_on=False, markersize=5)
    ax.plot(0, 1, "^k", transform=ax.get_xaxis_transform(), clip_on=False, markersize=5)


def fig_main():
    fig, ax = plt.subplots(figsize=(6.4, 4.4), dpi=220)
    x = np.linspace(-2.6, 2.6, 600)
    y = x**3 - 3 * x
    ax.plot(x, y, color=NAVY, lw=2.6, zorder=3)
    style_axes(ax, (-2.6, 2.6), (-4.2, 4.2))

    # 极值点
    ax.plot([-1], [2], "o", color=ORANGE, ms=9, zorder=5)
    ax.plot([1], [-2], "o", color=ORANGE, ms=9, zorder=5)
    ax.annotate("极大值 f(-1)=2", xy=(-1, 2), xytext=(-2.45, 2.95),
                fontsize=11, color=ORANGE, weight="bold",
                arrowprops=dict(arrowstyle="-", color=ORANGE, lw=1.1))
    ax.annotate("极小值 f(1)=-2", xy=(1, -2), xytext=(0.15, -3.6),
                fontsize=11, color=ORANGE, weight="bold",
                arrowprops=dict(arrowstyle="-", color=ORANGE, lw=1.1))

    # 水平切线示意
    for x0, y0 in [(-1, 2), (1, -2)]:
        ax.plot([x0 - 0.75, x0 + 0.75], [y0, y0], ls="--", lw=1.2, color=GRAY, zorder=2)

    ax.text(1.75, 1.6, r"$y=f(x)$", fontsize=13, color=NAVY, weight="bold")
    ax.set_xticks([-2, -1, 1, 2])
    ax.set_yticks([-3, -2, 2, 3])
    ax.set_ylabel("")
    fig.tight_layout()
    fig.savefig(f"{OUT}/fig_main.png", transparent=False, facecolor="white")
    plt.close(fig)


def fig_counter():
    fig, ax = plt.subplots(figsize=(5.0, 4.0), dpi=220)
    x = np.linspace(-1.9, 1.9, 500)
    y = x**3
    ax.plot(x, y, color=NAVY, lw=2.6, zorder=3)
    style_axes(ax, (-1.9, 1.9), (-4.2, 4.2))
    ax.plot([0], [0], "o", color=GRAY, ms=9, zorder=5)
    ax.plot([-0.9, 0.9], [0, 0], ls="--", lw=1.2, color=GRAY, zorder=2)
    ax.annotate("f′(0)=0", xy=(0, 0), xytext=(0.28, 1.15), fontsize=11.5,
                color="#555A63", weight="bold",
                arrowprops=dict(arrowstyle="-", color=GRAY, lw=1.1))
    ax.text(0.12, -3.85, "但 x=0 不是极值点\n（f(x) 在 R 上单调递增）", fontsize=10.5,
            color="#555A63", va="bottom", ha="left",
            bbox=dict(boxstyle="round,pad=0.35", fc="white", ec="#D5D8DD", lw=0.8))
    ax.text(1.25, 3.0, r"$y=x^{3}$", fontsize=13, color=NAVY, weight="bold")
    ax.set_xticks([-1, 1])
    ax.set_yticks([-3, 3])
    fig.tight_layout()
    fig.savefig(f"{OUT}/fig_counter.png", transparent=False, facecolor="white")
    plt.close(fig)


if __name__ == "__main__":
    fig_main()
    fig_counter()
    print("figures done")
