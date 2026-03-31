"""
Agent Brain v4 — clean version, no pbix_editor dependency.
Handles AI chat with full context about the loaded .pbix file.
"""
import json
from groq import Groq


class AgentBrain:
    def __init__(self, groq_api_key: str, pbix_path: str = None):
        self.client    = Groq(api_key=groq_api_key)
        self.pbix_path = pbix_path
        self.history   = []

    def chat(self, user_message: str, pages=None, schema=None, themes=None) -> dict:
        system = f"""You are an expert Power BI dashboard designer and data analyst AI agent.

LOADED FILE: {self.pbix_path or 'No file loaded'}
PAGES: {[p['displayName'] for p in (pages or [])] or 'None'}
DATASET: {json.dumps(schema) if schema else 'Not available'}
CURRENT THEMES: {json.dumps(themes) if themes else 'None generated yet'}

You help with:
1. DAX measures — write complete, correct DAX for any metric
2. Design advice — layout, color, visual type recommendations
3. Python visual scripts — generate or modify scripts for premium visuals
4. Data insights — analyze and suggest improvements
5. Theme recommendations — suggest palettes and explain choices

Be specific, actionable, and reference actual page/column names."""

        self.history.append({"role": "user", "content": user_message})
        messages = [{"role": "system", "content": system}] + self.history

        response = self.client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=messages,
            max_tokens=1000,
            temperature=0.4,
        )
        reply = response.choices[0].message.content
        self.history.append({"role": "assistant", "content": reply})

        if len(self.history) > 20:
            self.history = self.history[-20:]

        return {"reply": reply, "executed": False, "backup": None, "error": None}

    def reload(self):
        self.history = []