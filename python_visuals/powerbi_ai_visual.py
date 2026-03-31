# ============================================================
# Power BI AI Agent — Smart Insights Python Visual
# Analyzes your data with AI and renders premium visuals
#
# SETUP:
# 1. pip install groq matplotlib pandas numpy
# 2. Replace YOUR_GROQ_KEY below with your key
# 3. In Power BI Desktop: Insert → Python visual
# 4. Drag your data fields into the visual
# 5. Paste this script and click Run
# ============================================================

import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import numpy as np
import json
import warnings
warnings.filterwarnings('ignore')

GROQ_API_KEY = "YOUR_GROQ_KEY"  # ← paste your Groq key here

# ── THEME ─────────────────────────────────────────────────────────────────────
DARK_BG  = "#0d1b2a"
CARD_BG  = "#1a2940"
ACCENT   = "#17a98e"
ACCENT2  = "#2e86de"
GOLD     = "#c9a84c"
TEXT     = "#e8f4f8"
MUTED    = "#7a9bb5"
GREEN    = "#22c55e"
RED      = "#ef4444"
GRID     = "#1e3a5f"
COLORS   = [ACCENT, ACCENT2, GOLD, "#9b59b6", "#e67e22", "#1abc9c", "#e74c3c"]

plt.rcParams.update({
    'figure.facecolor': DARK_BG, 'axes.facecolor': CARD_BG,
    'axes.edgecolor': GRID, 'axes.labelcolor': MUTED,
    'axes.titlecolor': TEXT, 'xtick.color': MUTED, 'ytick.color': MUTED,
    'text.color': TEXT, 'grid.color': GRID, 'grid.alpha': 0.3,
    'font.family': 'sans-serif', 'font.size': 9,
})

# ── DATA ──────────────────────────────────────────────────────────────────────
df          = dataset.copy()
num_cols    = df.select_dtypes(include=[np.number]).columns.tolist()
date_cols   = df.select_dtypes(include=['datetime64']).columns.tolist()
cat_cols    = df.select_dtypes(include=['object','category']).columns.tolist()

# ── AI INSIGHTS via Groq ──────────────────────────────────────────────────────
insights = []
try:
    from groq import Groq
    client = Groq(api_key=GROQ_API_KEY)

    summary = {
        "rows": len(df),
        "columns": list(df.columns),
        "numeric_stats": {col: {
            "mean": round(df[col].mean(), 2),
            "max":  round(df[col].max(), 2),
            "min":  round(df[col].min(), 2),
            "sum":  round(df[col].sum(), 2),
        } for col in num_cols[:5]},
        "categories": {col: df[col].value_counts().head(5).to_dict() for col in cat_cols[:2]},
    }

    response = client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        max_tokens=400,
        messages=[{
            "role": "user",
            "content": f"""Analyze this dataset and give 4 short business insights (max 12 words each).
Data: {json.dumps(summary)}
Format: return ONLY a JSON array of 4 strings, no other text.
Example: ["Revenue grew 20% YoY driven by Q4", "Top region is West with 45% share"]"""
        }]
    )
    text = response.choices[0].message.content.strip()
    text = text.replace("```json","").replace("```","").strip()
    insights = json.loads(text)
except Exception as e:
    insights = [
        f"Dataset: {len(df):,} rows × {len(df.columns)} cols",
        f"Top metric: {num_cols[0] if num_cols else 'N/A'} = {df[num_cols[0]].sum():,.0f}" if num_cols else "No numeric data",
        f"Categories: {', '.join(cat_cols[:3])}" if cat_cols else "No categories",
        f"Date range detected: {len(date_cols)} date columns" if date_cols else "No date columns",
    ]

# ── LAYOUT ────────────────────────────────────────────────────────────────────
fig = plt.figure(figsize=(16, 10), facecolor=DARK_BG)

# Header bar
header = fig.add_axes([0, 0.94, 1, 0.06])
header.set_facecolor(CARD_BG)
header.axis('off')
header.text(0.02, 0.5, "AI-Powered Dashboard", fontsize=16, fontweight='bold',
            color=TEXT, va='center', transform=header.transAxes)
header.text(0.98, 0.5, f"{len(df):,} records  ·  {len(df.columns)} fields",
            fontsize=10, color=MUTED, va='center', ha='right', transform=header.transAxes)
# Accent line
line = fig.add_axes([0, 0.935, 1, 0.005])
line.set_facecolor(ACCENT)
line.axis('off')

# ── KPI CARDS ────────────────────────────────────────────────────────────────
n_kpis = min(4, len(num_cols))
for i, col in enumerate(num_cols[:n_kpis]):
    ax = fig.add_axes([0.01 + i*0.245, 0.79, 0.23, 0.13])
    ax.set_facecolor(CARD_BG)
    for spine in ax.spines.values():
        spine.set_edgecolor(COLORS[i]); spine.set_linewidth(1.5)
    ax.set_xticks([]); ax.set_yticks([])
    val = df[col].sum()
    val_str = f"{val/1e6:.2f}M" if abs(val)>=1e6 else f"{val/1e3:.1f}K" if abs(val)>=1e3 else f"{val:.1f}"
    pct = 0
    if len(df) > 1 and df[col].iloc[0] != 0:
        pct = (df[col].iloc[-1] - df[col].iloc[0]) / abs(df[col].iloc[0]) * 100
    ax.text(0.5, 0.75, col[:18], transform=ax.transAxes, ha='center', fontsize=8, color=MUTED)
    ax.text(0.5, 0.44, val_str, transform=ax.transAxes, ha='center', fontsize=20,
            fontweight='bold', color=COLORS[i])
    trend_c = GREEN if pct >= 0 else RED
    ax.text(0.5, 0.13, f"{'▲' if pct>=0 else '▼'} {abs(pct):.1f}%",
            transform=ax.transAxes, ha='center', fontsize=9, color=trend_c)

# ── MAIN TREND CHART ─────────────────────────────────────────────────────────
ax1 = fig.add_axes([0.01, 0.42, 0.46, 0.34])
if num_cols:
    col = num_cols[0]
    if date_cols:
        df2 = df.sort_values(date_cols[0])
        x   = range(len(df2))
        ax1.fill_between(x, df2[col], alpha=0.12, color=ACCENT)
        ax1.plot(x, df2[col], color=ACCENT, lw=2.5, marker='o', ms=2.5)
    elif cat_cols:
        grp  = df.groupby(cat_cols[0])[col].sum().sort_values(ascending=False).head(8)
        bars = ax1.bar(range(len(grp)), grp.values, color=COLORS[:len(grp)], alpha=0.85, width=0.65)
        ax1.set_xticks(range(len(grp)))
        ax1.set_xticklabels([str(x)[:10] for x in grp.index], rotation=30, ha='right', fontsize=7)
        for bar, v in zip(bars, grp.values):
            ax1.text(bar.get_x()+bar.get_width()/2, bar.get_height()*1.01, f"{v:,.0f}",
                     ha='center', fontsize=6, color=MUTED)
    else:
        ax1.plot(df[col].values, color=ACCENT, lw=2)
    ax1.set_title(f"{col} Analysis", fontsize=11, fontweight='bold', pad=8)
    ax1.grid(axis='y', alpha=0.25)
    ax1.spines['top'].set_visible(False); ax1.spines['right'].set_visible(False)

# ── SECONDARY CHART ──────────────────────────────────────────────────────────
ax2 = fig.add_axes([0.52, 0.42, 0.46, 0.34])
if len(num_cols) >= 2 and cat_cols:
    col2 = num_cols[1]
    grp  = df.groupby(cat_cols[0])[col2].sum().sort_values(ascending=False).head(6)
    wedges, texts, autos = ax2.pie(
        grp.values, labels=[str(x)[:12] for x in grp.index],
        colors=COLORS[:len(grp)], autopct='%1.1f%%', pctdistance=0.78,
        wedgeprops=dict(width=0.55, edgecolor=DARK_BG, lw=2)
    )
    for t in texts:  t.set_color(MUTED); t.set_fontsize(8)
    for a in autos:  a.set_color(TEXT);  a.set_fontsize(7)
    ax2.set_title(f"{col2} Breakdown", fontsize=11, fontweight='bold', pad=8)
elif len(num_cols) >= 2:
    ax2.scatter(df[num_cols[0]], df[num_cols[1]], color=GOLD, alpha=0.5, s=15)
    ax2.set_title(f"{num_cols[0]} vs {num_cols[1]}", fontsize=11, fontweight='bold', pad=8)
    ax2.grid(alpha=0.2)
    ax2.spines['top'].set_visible(False); ax2.spines['right'].set_visible(False)

# ── AI INSIGHTS PANEL ────────────────────────────────────────────────────────
ax3 = fig.add_axes([0.01, 0.04, 0.46, 0.35])
ax3.set_facecolor(CARD_BG)
ax3.axis('off')
ax3.text(0.04, 0.92, "✦ AI Insights", transform=ax3.transAxes,
         fontsize=11, fontweight='bold', color=ACCENT)
for i, insight in enumerate(insights[:4]):
    y = 0.72 - i * 0.19
    # Bullet
    circle = plt.Circle((0.04, y+0.01), 0.015, transform=ax3.transAxes,
                         color=COLORS[i], zorder=5)
    ax3.add_patch(circle)
    ax3.text(0.08, y, insight, transform=ax3.transAxes,
             fontsize=9, color=TEXT, va='top', wrap=True,
             bbox=dict(boxstyle='round,pad=0.3', facecolor=DARK_BG, edgecolor=GRID, alpha=0.6))
for spine in ax3.spines.values():
    spine.set_edgecolor(ACCENT); spine.set_linewidth(1.5)

# ── DISTRIBUTION ─────────────────────────────────────────────────────────────
ax4 = fig.add_axes([0.52, 0.04, 0.46, 0.35])
if num_cols:
    col = num_cols[0]
    n, bins, patches_hist = ax4.hist(df[col].dropna(), bins=25,
                                      color=ACCENT2, alpha=0.75, edgecolor=DARK_BG, lw=0.5)
    # Color bars by value
    for patch, left in zip(patches_hist, bins[:-1]):
        patch.set_facecolor(ACCENT if left < df[col].median() else GOLD)
        patch.set_alpha(0.8)
    ax4.axvline(df[col].mean(),   color=GREEN, lw=1.5, ls='--', label='Mean')
    ax4.axvline(df[col].median(), color=GOLD,  lw=1.5, ls=':',  label='Median')
    ax4.set_title(f"{col[:20]} Distribution", fontsize=11, fontweight='bold', pad=8)
    ax4.legend(fontsize=8, facecolor=CARD_BG, edgecolor=GRID, labelcolor=TEXT)
    ax4.grid(axis='y', alpha=0.2)
    ax4.spines['top'].set_visible(False); ax4.spines['right'].set_visible(False)

plt.savefig('ai_dashboard.png', dpi=150, bbox_inches='tight',
            facecolor=DARK_BG, edgecolor='none')
plt.show()
