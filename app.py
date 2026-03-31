"""
Power BI AI Agent — Web App v6.0
Clean version — no pbix_editor dependency.
Features: Theme generation, AI chat, DAX advice, design recommendations.
Run: python app.py  →  open http://localhost:5000
"""

import os, json
from flask import Flask, request, jsonify, render_template
from flask_cors import CORS
from color_engine  import ColorEngine
from theme_builder import ThemeBuilder
from pbi_client    import PBIClient
from groq import Groq

GROQ_API_KEY = os.getenv("GROQ_API_KEY", "YOUR_GROQ_API_KEY")

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "change-me-in-production")
CORS(app)

color = ColorEngine(GROQ_API_KEY)
theme = ThemeBuilder()
pbi   = PBIClient()

# ── Global state ───────────────────────────────────────────────────────────────
state = {
    "pages":     [],
    "themes":    {},
    "pbix_path": None,
    "schema":    {},
    "chat_history": [],
}


@app.route("/")
def index():
    return render_template("index.html")


# ── File ops ──────────────────────────────────────────────────────────────────
@app.route("/api/scan", methods=["POST"])
def scan():
    found = pbi.find_pbix_files()
    return jsonify({"files": [str(f) for f in found]})


@app.route("/api/load", methods=["POST"])
def load():
    path = request.json.get("path", "")
    try:
        pbi.set_pbix(path)
        pages  = pbi.list_pages()
        full_ctx = pbi.get_full_report_context()
        state["pages"]         = pages
        state["pbix_path"]     = path
        state["themes"]        = {}
        state["schema"]        = full_ctx
        state["chat_history"]  = []
        return jsonify({"ok": True, "pages": pages, "schema": {"tables": full_ctx.get("tables",[]), "measures": full_ctx.get("measures",[]), "relationships": full_ctx.get("relationships",[])}})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)})


@app.route("/api/state", methods=["GET"])
def get_state():
    return jsonify({
        "pages":     state["pages"],
        "themes":    state["themes"],
        "pbix_path": state["pbix_path"],
    })


# ── Theme ops ─────────────────────────────────────────────────────────────────
@app.route("/api/auto", methods=["POST"])
def auto_detect():
    results = []
    for pg in state["pages"]:
        try:
            palette = color.auto_detect(pg["displayName"])
            state["themes"][pg["displayName"]] = palette
            results.append({"page": pg["displayName"], "palette": palette})
        except Exception as e:
            results.append({"page": pg["displayName"], "error": str(e)})
    return jsonify({"ok": True, "results": results})


@app.route("/api/describe", methods=["POST"])
def describe():
    page_name = request.json.get("page", "")
    vibe      = request.json.get("vibe", "")
    targets   = state["pages"] if page_name.lower() == "all" else \
                [p for p in state["pages"] if p["displayName"] == page_name]
    results = []
    for pg in targets:
        try:
            palette = color.from_description(vibe, pg["displayName"])
            state["themes"][pg["displayName"]] = palette
            results.append({"page": pg["displayName"], "palette": palette})
        except Exception as e:
            results.append({"page": pg["displayName"], "error": str(e)})
    return jsonify({"ok": True, "results": results})


@app.route("/api/browse", methods=["GET"])
def browse():
    page    = request.args.get("page", "")
    options = color.curated_palettes_for_page(page)
    return jsonify({"ok": True, "options": options})


@app.route("/api/browse/apply", methods=["POST"])
def browse_apply():
    page    = request.json.get("page", "")
    index   = request.json.get("index", 0)
    options = color.curated_palettes_for_page(page)
    palette = color.build_full_palette(options[index])
    state["themes"][page] = palette
    return jsonify({"ok": True, "palette": palette})


@app.route("/api/export", methods=["POST"])
def export():
    saved = []
    for page_name, palette in state["themes"].items():
        pbix_theme = theme.build_pbix_theme(page_name, palette)
        fname      = f"theme_{page_name.replace(' ', '_')}.json"
        with open(fname, "w") as f:
            json.dump(pbix_theme, f, indent=2)
        saved.append(fname)
    return jsonify({"ok": True, "files": saved})


# ── AI Chat ───────────────────────────────────────────────────────────────────
@app.route("/api/agent/chat", methods=["POST"])
def agent_chat():
    msg = request.json.get("message", "")
    if not msg:
        return jsonify({"ok": False, "error": "Empty message"})

    client = Groq(api_key=GROQ_API_KEY)

    # Build full file context
    ctx = state.get("schema", {})
    pages_sum = []
    for p in ctx.get("pages", []):
        vis_list = []
        for v in p.get("visuals", []):
            s = v["type"]
            if v.get("title"):  s += f" titled '{v['title']}'"
            if v.get("fields"): s += f" using [{', '.join(v['fields'])}]"
            vis_list.append(s)
        pages_sum.append(f"  Page '{p['name']}': {', '.join(vis_list) or 'no visuals'}")
    meas_sum = [f"  [{m['table']}].[{m['name']}] = {m['expression']}" for m in ctx.get("measures",[])]
    tbls_sum = [f"  {t['name']}: {', '.join([c['name'] for c in t.get('columns',[]) if not c.get('isHidden')][:10])}" for t in ctx.get("tables",[]) if not t.get("isHidden")]
    rels_sum = [f"  {r['from']} to {r['to']}" for r in ctx.get("relationships",[])]
    system = f"""You are an expert Power BI dashboard designer and data analyst AI agent with FULL access to the loaded .pbix file.

FILE: {state['pbix_path'] or 'No file loaded'}
VERSION: {ctx.get('version','unknown')}
THEME: {ctx.get('theme','unknown')}

PAGES AND VISUALS:
{chr(10).join(pages_sum) or 'None'}

DAX MEASURES (full expressions):
{chr(10).join(meas_sum) or 'None found — measures may not be readable from this file version'}

TABLES AND COLUMNS:
{chr(10).join(tbls_sum) or 'None'}

RELATIONSHIPS:
{chr(10).join(rels_sum) or 'None'}

REPORT FILTERS: {json.dumps(ctx.get('filters',[]))}
THEMES GENERATED: {list(state['themes'].keys()) if state['themes'] else 'None'}

You can:
1. Review DAX measures for correctness, syntax errors and logical issues — reference the actual expressions above
2. Review each page's visualizations — check chart types, fields used, layout
3. Write new DAX using actual table/column names from this file
4. Suggest design and layout improvements
5. Generate theme JSON for Power BI import
6. Detect issues: wrong relationships, missing measures, inefficient patterns

Always reference actual names from the file. Format DAX with line breaks."""

    # Maintain conversation history
    history = state["chat_history"]
    history.append({"role": "user", "content": msg})

    messages = [{"role": "system", "content": system}] + history

    response = client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=messages,
        max_tokens=1000,
        temperature=0.4,
    )
    reply = response.choices[0].message.content
    history.append({"role": "assistant", "content": reply})

    # Keep history manageable
    if len(history) > 20:
        state["chat_history"] = history[-20:]

    return jsonify({"ok": True, "reply": reply, "executed": False, "error": None})


@app.route("/api/agent/reload", methods=["POST"])
def agent_reload():
    """Reload file info."""
    if state["pbix_path"]:
        try:
            pbi.set_pbix(state["pbix_path"])
            pages  = pbi.list_pages()
            full_ctx = pbi.get_full_report_context()
            state["pages"]  = pages
            state["schema"] = schema
            return jsonify({"ok": True, "pages": pages})
        except Exception as e:
            return jsonify({"ok": False, "error": str(e)})
    return jsonify({"ok": False, "error": "No file loaded"})


if __name__ == "__main__":
    print("\n╔══════════════════════════════════════════════╗")
    print("║   Power BI AI Agent  v6.0                   ║")
    print("║   Open: http://localhost:5000               ║")
    print("╚══════════════════════════════════════════════╝\n")
    app.run(debug=True, port=5000)


@app.route("/api/download/<page_name>", methods=["GET"])
def download_theme(page_name):
    from flask import Response
    decoded = page_name.replace("_"," ")
    palette = None
    matched_key = None
    for key in state["themes"]:
        if key.lower() == decoded.lower() or key.replace(" ","_").lower() == page_name.lower():
            palette     = state["themes"][key]
            matched_key = key
            break
    if not palette:
        return jsonify({"error": "Theme not found — generate a theme first"}), 404
    pbix_theme = theme.build_pbix_theme(matched_key, palette)
    fname      = f"theme_{matched_key.replace(' ','_')}.json"
    return Response(
        json.dumps(pbix_theme, indent=2),
        mimetype="application/json",
        headers={"Content-Disposition": f"attachment; filename={fname}"}
    )


@app.route("/api/download_all", methods=["GET"])
def download_all_themes():
    import io, zipfile
    from flask import send_file
    if not state["themes"]:
        return jsonify({"error": "No themes generated yet"}), 404
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for page_name, palette in state["themes"].items():
            pbix_theme = theme.build_pbix_theme(page_name, palette)
            fname      = f"theme_{page_name.replace(' ','_')}.json"
            zf.writestr(fname, json.dumps(pbix_theme, indent=2))
    buf.seek(0)
    return send_file(buf, mimetype="application/zip",
                     as_attachment=True, download_name="pbi_themes.zip")