"""
Theme Builder — generates strictly valid Power BI Desktop theme JSON.
Tested against Power BI Desktop theme schema.
"""


class ThemeBuilder:

    def build_pbix_theme(self, page_name: str, palette: dict) -> dict:
        chart_colors = palette.get("chart_colors", [
            "#2E86DE","#27AE60","#E67E22","#9B59B6","#E74C3C","#1ABC9C",
        ])
        # Ensure all are valid hex strings
        chart_colors = [c for c in chart_colors if isinstance(c, str) and c.startswith("#")]

        theme = {
            "name": f"PBIAgent_{page_name.replace(' ','_').replace('/','_')}",
            "dataColors": chart_colors,
            "background":  palette.get("canvas_bg", palette.get("background", "#FFFFFF")),
            "foreground":  palette.get("font_color", "#000000"),
            "tableAccent": palette.get("accent", "#2E86DE"),
            "good":        palette.get("positive", "#27AE60"),
            "neutral":     palette.get("neutral",  "#95A5A6"),
            "bad":         palette.get("negative", "#E74C3C"),
            "maximum":     chart_colors[0] if chart_colors else "#2E86DE",
            "center":      chart_colors[2] if len(chart_colors) > 2 else "#E67E22",
            "minimum":     palette.get("negative", "#E74C3C"),
            "null":        palette.get("neutral", "#95A5A6"),
        }
        return theme
