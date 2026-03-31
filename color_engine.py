"""
Color Engine — powered by Groq (free, no card, no regional restrictions)
Three color modes:
  1. describe  — user describes a vibe, AI picks a full palette
  2. browse    — curated palettes per page type
  3. auto      — AI infers palette from page name
"""

import json
import re
from groq import Groq


CURATED_PALETTES = {
    "sales": [
        {"name": "Navy & Gold",      "colors": ["#0D2B55","#1A4A8A","#C9A84C","#F2D98B","#FFFFFF"], "vibe": "premium"},
        {"name": "Electric Blue",    "colors": ["#0A1628","#1B3D6E","#2E86DE","#A8D8F0","#FFFFFF"], "vibe": "modern"},
        {"name": "Forest & Slate",   "colors": ["#1C3A2A","#2E6B4F","#4CAF82","#B2DFC9","#F5FAF7"], "vibe": "fresh"},
        {"name": "Crimson Pro",      "colors": ["#1A0A0A","#5C1A1A","#C0392B","#F1948A","#FFF5F5"], "vibe": "bold"},
    ],
    "marketing": [
        {"name": "Coral Campaign",   "colors": ["#1A0D00","#7A3B1E","#E07B39","#F5C99A","#FFF8F3"], "vibe": "energetic"},
        {"name": "Purple Pulse",     "colors": ["#110A1C","#4A1A7A","#8E44AD","#D2A8E8","#F9F0FF"], "vibe": "creative"},
        {"name": "Teal & Lime",      "colors": ["#061414","#0D5C5C","#17A98E","#7FE0D1","#F0FAFA"], "vibe": "vibrant"},
        {"name": "Midnight Rose",    "colors": ["#150810","#5C1A3A","#B03060","#E8A0BA","#FFF0F5"], "vibe": "bold"},
    ],
    "finance": [
        {"name": "Deep Charcoal",    "colors": ["#0F0F0F","#2D2D2D","#4A90D9","#A8C8F0","#F5F5F5"], "vibe": "serious"},
        {"name": "Green Ledger",     "colors": ["#081408","#1A4020","#27AE60","#A9DFBF","#F0FFF4"], "vibe": "growth"},
        {"name": "Slate & Amber",    "colors": ["#0D0D0A","#3D3D1A","#B8860B","#F0D080","#FFFFF0"], "vibe": "warm"},
        {"name": "Steel Blue",       "colors": ["#080D14","#1A2A3D","#2C5F8A","#8FBDD3","#EEF6FA"], "vibe": "corporate"},
    ],
    "operations": [
        {"name": "Industrial Grey",  "colors": ["#0A0A0A","#303030","#607D8B","#B0BEC5","#ECEFF1"], "vibe": "utility"},
        {"name": "Blueprint",        "colors": ["#040D1C","#0D2B55","#1565C0","#90CAF9","#E3F2FD"], "vibe": "technical"},
        {"name": "Olive & Rust",     "colors": ["#0C0A00","#3D3000","#8D6E1A","#D4B483","#FFF8E1"], "vibe": "earthy"},
        {"name": "Monochrome+",      "colors": ["#050505","#1A1A1A","#4A4A4A","#AAAAAA","#F0F0F0"], "vibe": "minimal"},
    ],
    "executive": [
        {"name": "Boardroom",        "colors": ["#08080F","#1A1A2E","#16213E","#4A90D9","#F5F5FF"], "vibe": "luxury"},
        {"name": "Platinum",         "colors": ["#0A0A0A","#2A2A2A","#7A7A7A","#C8C8C8","#F8F8F8"], "vibe": "minimal"},
        {"name": "Sapphire & Gold",  "colors": ["#060B14","#0D2040","#1A4A8A","#C9A84C","#FAFAFA"], "vibe": "premium"},
        {"name": "Forest Executive", "colors": ["#050D08","#0D2A14","#1A5C2A","#5AAF7A","#F0FAF2"], "vibe": "fresh"},
    ],
    "default": [
        {"name": "Ocean Deep",       "colors": ["#030D1A","#0A2A40","#1565C0","#64B5F6","#E3F2FD"], "vibe": "cool"},
        {"name": "Sunset Pro",       "colors": ["#1A0500","#4A1500","#C0392B","#F39C12","#FFFDE7"], "vibe": "warm"},
        {"name": "Arctic",           "colors": ["#050A0F","#0D1F30","#2196F3","#B3E5FC","#F0F9FF"], "vibe": "clean"},
        {"name": "Ember",            "colors": ["#0F0500","#3D1500","#E65100","#FFB74D","#FFF8E1"], "vibe": "bold"},
    ],
}

PAGE_CATEGORY_MAP = {
    "sales":      ["sales", "revenue", "deals", "pipeline", "quota"],
    "marketing":  ["marketing", "campaign", "funnel", "roi", "leads"],
    "finance":    ["finance", "p&l", "profit", "budget", "cost", "margin"],
    "operations": ["operations", "ops", "supply", "logistics", "inventory"],
    "executive":  ["executive", "summary", "overview", "ceo", "board", "kpi"],
}

PALETTE_PROMPT = """You are a Power BI dashboard design expert.
Generate a beautiful professional color palette for a Power BI dashboard page.

Page name: "{page_name}"
{vibe_line}

Return ONLY a valid JSON object, no markdown, no explanation, just raw JSON:
{{
  "background": "#hexcolor",
  "canvas_bg": "#hexcolor",
  "chart_colors": ["#hex1","#hex2","#hex3","#hex4","#hex5","#hex6"],
  "table_header": "#hexcolor",
  "table_header_font": "#hexcolor",
  "table_row_alt": "#hexcolor",
  "table_row_base": "#hexcolor",
  "table_font": "#hexcolor",
  "kpi_card_bg": "#hexcolor",
  "kpi_card_border": "#hexcolor",
  "kpi_value_color": "#hexcolor",
  "kpi_label_color": "#hexcolor",
  "font_color": "#hexcolor",
  "accent": "#hexcolor",
  "positive": "#hexcolor",
  "negative": "#hexcolor",
  "neutral": "#hexcolor",
  "vibe_summary": "one sentence describing the look"
}}

Rules:
- Strong contrast between background and font (WCAG AA)
- Chart colors must be visually distinct from each other
- Table header darker than row colors
- KPI card slightly elevated from background
- Accent color must pop against background
"""


class ColorEngine:
    def __init__(self, groq_api_key: str):
        self.client = Groq(api_key=groq_api_key)
        self.model  = "llama-3.3-70b-versatile"

    def _ask(self, prompt: str) -> dict:
        response = self.client.chat.completions.create(
            model=self.model,
            messages=[{"role": "user", "content": prompt}],
            max_tokens=800,
            temperature=0.7,
        )
        raw = response.choices[0].message.content.strip()
        return self._safe_parse(raw)

    def from_description(self, vibe: str, page_name: str) -> dict:
        prompt = PALETTE_PROMPT.format(
            page_name=page_name,
            vibe_line=f'Desired vibe/mood: "{vibe}"'
        )
        return self._ask(prompt)

    def curated_palettes_for_page(self, page_name: str) -> list:
        category = self._detect_category(page_name)
        return CURATED_PALETTES.get(category, CURATED_PALETTES["default"])

    def build_full_palette(self, curated: dict) -> dict:
        bg, dark, mid, light, white = curated["colors"]
        return {
            "background":        bg,
            "canvas_bg":         self._lighten(bg, 0.05),
            "chart_colors":      [mid, light, dark, self._blend(mid, light), self._blend(dark, mid), white],
            "table_header":      dark,
            "table_header_font": white,
            "table_row_alt":     self._lighten(light, 0.5),
            "table_row_base":    white,
            "table_font":        bg,
            "kpi_card_bg":       self._lighten(bg, 0.08),
            "kpi_card_border":   mid,
            "kpi_value_color":   light,
            "kpi_label_color":   self._lighten(light, 0.3),
            "font_color":        white,
            "accent":            mid,
            "positive":          "#27AE60",
            "negative":          "#E74C3C",
            "neutral":           "#95A5A6",
            "vibe_summary":      f"{curated['name']} — {curated['vibe']} palette",
        }

    def auto_detect(self, page_name: str) -> dict:
        prompt = PALETTE_PROMPT.format(
            page_name=page_name,
            vibe_line="Infer the best professional theme from the page name alone."
        )
        return self._ask(prompt)

    def _detect_category(self, page_name: str) -> str:
        name_lower = page_name.lower()
        for category, keywords in PAGE_CATEGORY_MAP.items():
            if any(kw in name_lower for kw in keywords):
                return category
        return "default"

    def _safe_parse(self, raw: str) -> dict:
        raw = re.sub(r"```json|```", "", raw).strip()
        try:
            return json.loads(raw)
        except json.JSONDecodeError:
            return {
                "background": "#0D1B2A", "canvas_bg": "#162436",
                "chart_colors": ["#2E86DE","#27AE60","#E67E22","#9B59B6","#E74C3C","#1ABC9C"],
                "table_header": "#1A3A5C", "table_header_font": "#FFFFFF",
                "table_row_alt": "#EAF2FB", "table_row_base": "#FFFFFF", "table_font": "#0D1B2A",
                "kpi_card_bg": "#1A2E45", "kpi_card_border": "#2E86DE",
                "kpi_value_color": "#64B5F6", "kpi_label_color": "#90CAF9",
                "font_color": "#FFFFFF", "accent": "#2E86DE",
                "positive": "#27AE60", "negative": "#E74C3C", "neutral": "#95A5A6",
                "vibe_summary": "Default ocean-deep corporate palette",
            }

    @staticmethod
    def _hex_to_rgb(h):
        h = h.lstrip("#")
        return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))

    @staticmethod
    def _rgb_to_hex(r, g, b):
        return "#{:02X}{:02X}{:02X}".format(int(r), int(g), int(b))

    def _lighten(self, hex_color, amount):
        r, g, b = self._hex_to_rgb(hex_color)
        return self._rgb_to_hex(min(255,r+(255-r)*amount), min(255,g+(255-g)*amount), min(255,b+(255-b)*amount))

    def _blend(self, h1, h2):
        r1,g1,b1 = self._hex_to_rgb(h1)
        r2,g2,b2 = self._hex_to_rgb(h2)
        return self._rgb_to_hex((r1+r2)//2,(g1+g2)//2,(b1+b2)//2)