# ============================================================
# Power BI AI Agent — Internal Python Script
# Run this INSIDE Power BI Desktop via the Python visual
# 
# HOW TO USE:
# 1. In Power BI Desktop, go to Insert → Python visual
# 2. Drag any fields you want analyzed into the visual
# 3. Paste this entire script into the Python script editor
# 4. Click Run
# ============================================================

import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
import warnings
warnings.filterwarnings('ignore')

# ── PREMIUM DARK THEME ────────────────────────────────────────────────────────
DARK_BG     = "#0d1b2a"
CARD_BG     = "#1a2940"
ACCENT      = "#17a98e"
ACCENT2     = "#2e86de"
GOLD        = "#c9a84c"
TEXT        = "#e8f4f8"
TEXT_MUTED  = "#7a9bb5"
POSITIVE    = "#22c55e"
NEGATIVE    = "#ef4444"
GRID_COLOR  = "#1e3a5f"

CHART_COLORS = [ACCENT, ACCENT2, GOLD, "#9b59b6", "#e67e22", "#e74c3c", "#1abc9c"]

plt.rcParams.update({
    'figure.facecolor':  DARK_BG,
    'axes.facecolor':    CARD_BG,
    'axes.edgecolor':    GRID_COLOR,
    'axes.labelcolor':   TEXT_MUTED,
    'axes.titlecolor':   TEXT,
    'xtick.color':       TEXT_MUTED,
    'ytick.color':       TEXT_MUTED,
    'text.color':        TEXT,
    'grid.color':        GRID_COLOR,
    'grid.alpha':        0.4,
    'font.family':       'sans-serif',
    'font.size':         10,
})

# ── READ DATA FROM POWER BI ───────────────────────────────────────────────────
# 'dataset' is automatically provided by Power BI
df = dataset.copy()

print(f"Dataset shape: {df.shape}")
print(f"Columns: {list(df.columns)}")

# ── AUTO-DETECT COLUMN TYPES ─────────────────────────────────────────────────
numeric_cols  = df.select_dtypes(include=[np.number]).columns.tolist()
date_cols     = df.select_dtypes(include=['datetime64']).columns.tolist()
category_cols = df.select_dtypes(include=['object', 'category']).columns.tolist()

print(f"Numeric : {numeric_cols}")
print(f"Dates   : {date_cols}")
print(f"Category: {category_cols}")

# ── BUILD DASHBOARD ───────────────────────────────────────────────────────────
n_numeric = len(numeric_cols)
has_date  = len(date_cols) > 0
has_cat   = len(category_cols) > 0

fig = plt.figure(figsize=(16, 10), facecolor=DARK_BG)
fig.patch.set_facecolor(DARK_BG)

# Title
fig.suptitle("AI-Generated Dashboard", fontsize=18, fontweight='bold',
             color=TEXT, y=0.98, x=0.5)

# ── KPI CARDS (top strip) ─────────────────────────────────────────────────────
n_kpis = min(4, n_numeric)
if n_kpis > 0:
    for i, col in enumerate(numeric_cols[:n_kpis]):
        ax_kpi = fig.add_axes([0.02 + i * 0.245, 0.82, 0.22, 0.13])
        ax_kpi.set_facecolor(CARD_BG)
        for spine in ax_kpi.spines.values():
            spine.set_edgecolor(ACCENT)
            spine.set_linewidth(1.5)
        ax_kpi.set_xticks([])
        ax_kpi.set_yticks([])

        val     = df[col].sum()
        mean    = df[col].mean()
        pct_chg = ((df[col].iloc[-1] - df[col].iloc[0]) / abs(df[col].iloc[0]) * 100) if len(df) > 1 and df[col].iloc[0] != 0 else 0

        # Format value
        if abs(val) >= 1_000_000:
            val_str = f"{val/1_000_000:.1f}M"
        elif abs(val) >= 1_000:
            val_str = f"{val/1_000:.1f}K"
        else:
            val_str = f"{val:.1f}"

        trend_color = POSITIVE if pct_chg >= 0 else NEGATIVE
        trend_arrow = "▲" if pct_chg >= 0 else "▼"

        ax_kpi.text(0.5, 0.72, col[:20], transform=ax_kpi.transAxes,
                    ha='center', va='center', fontsize=9, color=TEXT_MUTED)
        ax_kpi.text(0.5, 0.42, val_str, transform=ax_kpi.transAxes,
                    ha='center', va='center', fontsize=18, fontweight='bold', color=ACCENT)
        ax_kpi.text(0.5, 0.15, f"{trend_arrow} {abs(pct_chg):.1f}%",
                    transform=ax_kpi.transAxes,
                    ha='center', va='center', fontsize=9, color=trend_color)

# ── MAIN CHART (trend or bar) ─────────────────────────────────────────────────
if n_numeric >= 1:
    ax1 = fig.add_axes([0.02, 0.42, 0.45, 0.36])
    ax1.set_facecolor(CARD_BG)
    col = numeric_cols[0]

    if has_date:
        date_col = date_cols[0]
        df_sorted = df.sort_values(date_col)
        x = range(len(df_sorted))
        ax1.fill_between(x, df_sorted[col], alpha=0.15, color=ACCENT)
        ax1.plot(x, df_sorted[col], color=ACCENT, linewidth=2.5, marker='o',
                 markersize=3, markerfacecolor=ACCENT)
        ax1.set_title(f"{col} Over Time", pad=10, fontsize=11, fontweight='bold')
    elif has_cat:
        cat_col  = category_cols[0]
        grouped  = df.groupby(cat_col)[col].sum().sort_values(ascending=False).head(10)
        bars     = ax1.barh(range(len(grouped)), grouped.values,
                            color=CHART_COLORS[:len(grouped)], alpha=0.85, height=0.6)
        ax1.set_yticks(range(len(grouped)))
        ax1.set_yticklabels([str(x)[:15] for x in grouped.index], fontsize=8)
        ax1.set_title(f"{col} by {cat_col}", pad=10, fontsize=11, fontweight='bold')
        # Value labels
        for bar, val in zip(bars, grouped.values):
            ax1.text(val * 1.01, bar.get_y() + bar.get_height()/2,
                     f"{val:,.0f}", va='center', fontsize=7, color=TEXT_MUTED)
    else:
        ax1.plot(df[col].values, color=ACCENT, linewidth=2)
        ax1.set_title(col, pad=10, fontsize=11, fontweight='bold')

    ax1.grid(axis='x', alpha=0.3)
    ax1.spines['top'].set_visible(False)
    ax1.spines['right'].set_visible(False)

# ── SECONDARY CHART ───────────────────────────────────────────────────────────
if n_numeric >= 2:
    ax2 = fig.add_axes([0.52, 0.42, 0.45, 0.36])
    ax2.set_facecolor(CARD_BG)
    col2 = numeric_cols[1]

    if has_cat:
        cat_col = category_cols[0]
        grouped = df.groupby(cat_col)[col2].sum().sort_values(ascending=False).head(8)
        wedges, texts, autotexts = ax2.pie(
            grouped.values,
            labels=[str(x)[:12] for x in grouped.index],
            colors=CHART_COLORS[:len(grouped)],
            autopct='%1.1f%%',
            pctdistance=0.75,
            wedgeprops=dict(width=0.55, edgecolor=DARK_BG, linewidth=2)
        )
        for text in texts:      text.set_color(TEXT_MUTED); text.set_fontsize(8)
        for autotext in autotexts: autotext.set_color(TEXT); autotext.set_fontsize(7)
        ax2.set_title(f"{col2} Distribution", pad=10, fontsize=11, fontweight='bold')
    else:
        bars = ax2.bar(range(min(20, len(df))), df[col2].head(20),
                       color=ACCENT2, alpha=0.8)
        ax2.set_title(col2, pad=10, fontsize=11, fontweight='bold')
        ax2.grid(axis='y', alpha=0.3)
        ax2.spines['top'].set_visible(False)
        ax2.spines['right'].set_visible(False)

# ── BOTTOM: CORRELATION / SCATTER ────────────────────────────────────────────
if n_numeric >= 2:
    ax3 = fig.add_axes([0.02, 0.04, 0.30, 0.34])
    ax3.set_facecolor(CARD_BG)
    ax3.scatter(df[numeric_cols[0]], df[numeric_cols[1]],
                color=GOLD, alpha=0.6, s=20, edgecolors='none')
    # Trend line
    try:
        z   = np.polyfit(df[numeric_cols[0]].dropna(), df[numeric_cols[1]].dropna(), 1)
        p   = np.poly1d(z)
        x_l = np.linspace(df[numeric_cols[0]].min(), df[numeric_cols[0]].max(), 100)
        ax3.plot(x_l, p(x_l), color=ACCENT, linewidth=1.5, linestyle='--', alpha=0.8)
    except: pass
    ax3.set_xlabel(numeric_cols[0][:20], fontsize=8)
    ax3.set_ylabel(numeric_cols[1][:20], fontsize=8)
    ax3.set_title("Correlation", pad=8, fontsize=11, fontweight='bold')
    ax3.grid(alpha=0.2)
    ax3.spines['top'].set_visible(False)
    ax3.spines['right'].set_visible(False)

# ── BOTTOM: STATS TABLE ───────────────────────────────────────────────────────
if n_numeric >= 1:
    ax4 = fig.add_axes([0.36, 0.04, 0.30, 0.34])
    ax4.set_facecolor(CARD_BG)
    ax4.axis('off')

    stats_data = []
    for col in numeric_cols[:5]:
        stats_data.append([
            col[:15],
            f"{df[col].mean():,.1f}",
            f"{df[col].max():,.1f}",
            f"{df[col].min():,.1f}",
            f"{df[col].std():,.1f}",
        ])

    table = ax4.table(
        cellText=stats_data,
        colLabels=["Metric", "Mean", "Max", "Min", "StdDev"],
        cellLoc='center',
        loc='center',
        bbox=[0, 0, 1, 1]
    )
    table.auto_set_font_size(False)
    table.set_fontsize(8)
    for (r, c), cell in table.get_celld().items():
        cell.set_facecolor(DARK_BG if r == 0 else CARD_BG)
        cell.set_text_props(color=ACCENT if r == 0 else TEXT)
        cell.set_edgecolor(GRID_COLOR)
    ax4.set_title("Summary Statistics", pad=8, fontsize=11, fontweight='bold', color=TEXT,
                  x=0.5, y=0.98)

# ── BOTTOM: DISTRIBUTION ─────────────────────────────────────────────────────
if n_numeric >= 1:
    ax5 = fig.add_axes([0.70, 0.04, 0.28, 0.34])
    ax5.set_facecolor(CARD_BG)
    col = numeric_cols[0]
    ax5.hist(df[col].dropna(), bins=20, color=ACCENT2, alpha=0.7, edgecolor=DARK_BG)
    ax5.axvline(df[col].mean(), color=GOLD, linewidth=1.5, linestyle='--', label='Mean')
    ax5.set_title(f"{col[:20]} Distribution", pad=8, fontsize=11, fontweight='bold')
    ax5.legend(fontsize=8, facecolor=CARD_BG, edgecolor=GRID_COLOR, labelcolor=TEXT)
    ax5.grid(axis='y', alpha=0.3)
    ax5.spines['top'].set_visible(False)
    ax5.spines['right'].set_visible(False)

plt.savefig('powerbi_dashboard.png', dpi=150, bbox_inches='tight',
            facecolor=DARK_BG, edgecolor='none')
plt.show()
print("Dashboard rendered successfully!")
