"""
PBI Client — Deep Local Reader
Extracts everything from a .pbix file:
- Pages, visuals, field bindings, titles, positions
- DAX measures with full expressions
- Tables, columns, relationships
- Report config, themes, filters
- DataModel metadata
"""

import os
import json
import zipfile
from pathlib import Path
from datetime import datetime


class PBIClient:
    def __init__(self, pbix_path: str = None, *args, **kwargs):
        self.pbix_path    = pbix_path
        self.workspace_id = "local"
        self._layout      = None

    # ── File selection ────────────────────────────────────────────────────────
    def set_pbix(self, path: str):
        path = Path(path)
        if not path.exists():
            raise FileNotFoundError(f"File not found: {path}")
        if path.suffix.lower() != ".pbix":
            raise ValueError(f"Expected a .pbix file, got: {path.suffix}")
        self.pbix_path = str(path)
        self._layout   = None
        print(f"   Loaded: {path.name}  ({path.stat().st_size // 1024} KB)")

    def find_pbix_files(self, search_dir: str = None) -> list:
        search_dirs = []
        if search_dir:
            search_dirs.append(Path(search_dir))
        home = Path.home()
        search_dirs += [
            home / "Documents", home / "Desktop", home / "Downloads",
            home / "OneDrive" / "Documents", home / "OneDrive" / "Desktop",
            Path(__file__).parent,
        ]
        found = []
        for d in search_dirs:
            if d.exists():
                found += list(d.rglob("*.pbix"))
        return list(set(found))

    # ── ZIP utilities ─────────────────────────────────────────────────────────
    def _list_zip_entries(self) -> list:
        with zipfile.ZipFile(self.pbix_path, "r") as z:
            return z.namelist()

    def _read_entry(self, entry_name: str) -> bytes:
        with zipfile.ZipFile(self.pbix_path, "r") as z:
            return z.read(entry_name)

    def _decode(self, raw: bytes) -> str:
        if raw[:2] == b'\xff\xfe':
            return raw[2:].decode("utf-16-le")
        elif raw[:2] == b'\xfe\xff':
            return raw[2:].decode("utf-16-be")
        elif raw[:3] == b'\xef\xbb\xbf':
            return raw[3:].decode("utf-8")
        else:
            try:    return raw.decode("utf-16-le")
            except: return raw.decode("utf-8", errors="replace")

    def _read_json_entry(self, entry_name: str) -> dict:
        raw  = self._read_entry(entry_name)
        text = self._decode(raw)
        return json.loads(text)

    # ── Layout ────────────────────────────────────────────────────────────────
    def _read_layout(self) -> dict:
        if self._layout:
            return self._layout
        with zipfile.ZipFile(self.pbix_path, "r") as z:
            entry = next((n for n in z.namelist() if n == "Report/Layout"), None)
            if not entry:
                raise ValueError("Report/Layout not found in .pbix")
            raw = z.read(entry)
        self._layout = json.loads(self._decode(raw))
        return self._layout

    # ── Basic accessors ───────────────────────────────────────────────────────
    def list_reports(self) -> list:
        if not self.pbix_path:
            return []
        return [{"id": "local", "name": Path(self.pbix_path).stem, "path": self.pbix_path}]

    def list_pages(self, report_id: str = "local") -> list:
        layout = self._read_layout()
        return sorted([{
            "name":        s.get("name", ""),
            "displayName": s.get("displayName", s.get("name", "Page")),
            "order":       s.get("ordinal", 0),
        } for s in layout.get("sections", [])], key=lambda x: x["order"])

    def list_visuals(self, report_id: str, page_name: str) -> list:
        layout = self._read_layout()
        for section in layout.get("sections", []):
            if section.get("displayName") == page_name or section.get("name") == page_name:
                result = []
                for vc in section.get("visualContainers", []):
                    try:    cfg = json.loads(vc.get("config", "{}"))
                    except: cfg = {}
                    v_type = cfg.get("singleVisual", {}).get("visualType", "unknown")
                    result.append({
                        "type": v_type, "x": vc.get("x", 0), "y": vc.get("y", 0),
                        "width": vc.get("width", 0), "height": vc.get("height", 0),
                    })
                return result
        return {}

    def get_dataset_schema(self, dataset_id: str = "local") -> dict:
        return self.get_full_report_context()

    # ── FULL DEEP READER ──────────────────────────────────────────────────────
    def get_full_report_context(self) -> dict:
        """
        Extract EVERYTHING from the .pbix file:
        Pages, visuals, field bindings, DAX measures, tables, columns,
        relationships, report config, themes, filters, metadata.
        """
        ctx = {
            "file":          Path(self.pbix_path).name,
            "zip_entries":   [],
            "metadata":      {},
            "version":       "",
            "pages":         [],
            "measures":      [],
            "tables":        [],
            "relationships": [],
            "report_config": {},
            "theme":         "",
            "filters":       [],
            "errors":        [],
        }

        # ── List all ZIP entries ──────────────────────────────────────────────
        try:
            ctx["zip_entries"] = self._list_zip_entries()
        except Exception as e:
            ctx["errors"].append(f"ZIP listing: {e}")

        # ── Version ──────────────────────────────────────────────────────────
        try:
            if "Version" in ctx["zip_entries"]:
                ctx["version"] = self._decode(self._read_entry("Version")).strip()
        except Exception as e:
            ctx["errors"].append(f"Version: {e}")

        # ── Metadata ─────────────────────────────────────────────────────────
        try:
            if "Metadata" in ctx["zip_entries"]:
                ctx["metadata"] = self._read_json_entry("Metadata")
        except Exception as e:
            ctx["errors"].append(f"Metadata: {e}")

        # ── Pages + Visuals from Report/Layout ────────────────────────────────
        try:
            layout = self._read_layout()
            ctx["theme"]         = layout.get("theme", "")
            ctx["report_config"] = self._safe_parse(layout.get("config", "{}"))

            for section in layout.get("sections", []):
                page = {
                    "name":        section.get("displayName", section.get("name", "")),
                    "ordinal":     section.get("ordinal", 0),
                    "width":       section.get("width", 1280),
                    "height":      section.get("height", 720),
                    "filters":     self._safe_parse(section.get("filters", "[]")),
                    "config":      self._safe_parse(section.get("config", "{}")),
                    "visuals":     [],
                }

                for vc in section.get("visualContainers", []):
                    try:    cfg = json.loads(vc.get("config", "{}"))
                    except: cfg = {}

                    sv     = cfg.get("singleVisual", {})
                    vtype  = sv.get("visualType", "unknown")

                    # Title
                    title = ""
                    try:
                        title = (sv.get("vcObjects", {})
                                   .get("title", [{}])[0]
                                   .get("properties", {})
                                   .get("text", {})
                                   .get("expr", {})
                                   .get("Literal", {})
                                   .get("Value", "")).strip("'")
                    except: pass

                    # Build alias → real table name map from From clause
                    from_map = {}
                    try:
                        for src in sv.get("prototypeQuery", {}).get("From", []):
                            alias    = src.get("Name", "")
                            entity   = src.get("Entity", "")
                            if alias and entity:
                                from_map[alias] = entity
                    except: pass

                    # Field bindings — use NativeReferenceName for real names
                    fields_used = []
                    try:
                        for proj in sv.get("prototypeQuery", {}).get("Select", []):
                            # NativeReferenceName gives the clean display name
                            native_name = proj.get("NativeReferenceName", "")
                            full_name   = proj.get("Name", "")  # Table.FieldName format

                            col  = proj.get("Column", {})
                            meas = proj.get("Measure", {})
                            agg  = proj.get("Aggregation", {})

                            if meas:
                                alias = meas.get("Expression", {}).get("SourceRef", {}).get("Source", "")
                                tbl   = from_map.get(alias, alias)
                                fld   = meas.get("Property", native_name)
                                if fld: fields_used.append(f"{tbl}[{fld}] (measure)")
                            elif col:
                                alias = col.get("Expression", {}).get("SourceRef", {}).get("Source", "")
                                tbl   = from_map.get(alias, alias)
                                fld   = col.get("Property", native_name)
                                if fld: fields_used.append(f"{tbl}[{fld}]")
                            elif agg:
                                alias = (agg.get("Expression", {})
                                            .get("Column", {})
                                            .get("Expression", {})
                                            .get("SourceRef", {})
                                            .get("Source", ""))
                                tbl   = from_map.get(alias, alias)
                                fld   = (agg.get("Expression", {})
                                            .get("Column", {})
                                            .get("Property", native_name))
                                fn    = agg.get("Function", "")
                                if fld: fields_used.append(f"{fn}({tbl}[{fld}])")
                            elif native_name:
                                fields_used.append(native_name)
                    except: pass

                    # Filters on visual
                    vis_filters = []
                    try:
                        vis_filters = self._safe_parse(vc.get("filters", "[]"))
                    except: pass

                    # Objects (formatting overrides)
                    objects = sv.get("objects", {})

                    page["visuals"].append({
                        "type":    vtype,
                        "title":   title,
                        "fields":  fields_used,
                        "filters": vis_filters,
                        "objects": objects,
                        "x":       vc.get("x", 0),
                        "y":       vc.get("y", 0),
                        "width":   vc.get("width", 0),
                        "height":  vc.get("height", 0),
                        "z":       vc.get("z", 0),
                    })

                ctx["pages"].append(page)

        except Exception as e:
            ctx["errors"].append(f"Layout: {e}")

        # ── DataModelSchema — measures, tables, columns, relationships ─────────
        try:
            entries      = ctx["zip_entries"]
            schema_entry = next((n for n in entries if "DataModelSchema" in n), None)

            if schema_entry:
                schema = self._read_json_entry(schema_entry)
                model  = schema.get("model", {})
                for table in model.get("tables", []):
                    tbl_name  = table.get("name", "")
                    is_hidden = table.get("isHidden", False)
                    columns   = []
                    for col in table.get("columns", []):
                        columns.append({
                            "name": col.get("name",""), "dataType": col.get("dataType",""),
                            "isHidden": col.get("isHidden", False),
                            "expression": col.get("expression",""),
                        })
                    for meas in table.get("measures", []):
                        ctx["measures"].append({
                            "table": tbl_name, "name": meas.get("name",""),
                            "expression": meas.get("expression","").strip(),
                            "formatString": meas.get("formatString",""),
                        })
                    ctx["tables"].append({"name": tbl_name, "isHidden": is_hidden, "columns": columns})
                for rel in model.get("relationships", []):
                    ctx["relationships"].append({
                        "from": f"{rel.get('fromTable','')}[{rel.get('fromColumn','')}]",
                        "to":   f"{rel.get('toTable','')}[{rel.get('toColumn','')}]",
                        "isActive": rel.get("isActive", True),
                    })
            else:
                # Extract tables/measures from visual field references
                ctx["schema_note"] = "DataModelSchema not in this file. Tables/measures extracted from visual references."
                tables_seen   = {}
                measures_seen = set()
                layout = self._read_layout()
                for section in layout.get("sections", []):
                    for vc in section.get("visualContainers", []):
                        try:
                            cfg = json.loads(vc.get("config", "{}"))
                            sv  = cfg.get("singleVisual", {})
                            pq  = sv.get("prototypeQuery", {})
                            from_map = {s["Name"]: s["Entity"] for s in pq.get("From",[]) if "Name" in s and "Entity" in s}

                            # Register ALL tables that appear in From clause
                            for entity in from_map.values():
                                if entity and entity not in tables_seen:
                                    tables_seen[entity] = set()

                            for proj in pq.get("Select", []):
                                native   = proj.get("NativeReferenceName", "")
                                name     = proj.get("Name", "")
                                tbl_part = name.split(".")[0] if "." in name else ""
                                meas     = proj.get("Measure", {})
                                col      = proj.get("Column", {})
                                agg      = proj.get("Aggregation", {})

                                if meas:
                                    alias = meas.get("Expression",{}).get("SourceRef",{}).get("Source","")
                                    tbl   = from_map.get(alias, tbl_part)
                                    fld   = meas.get("Property", native)
                                    key   = f"{tbl}.{fld}"
                                    if key not in measures_seen and fld:
                                        measures_seen.add(key)
                                        ctx["measures"].append({
                                            "table": tbl, "name": fld,
                                            "expression": "(stored in compressed DataModel)",
                                        })
                                    if tbl and tbl not in tables_seen:
                                        tables_seen[tbl] = set()
                                elif col:
                                    alias = col.get("Expression",{}).get("SourceRef",{}).get("Source","")
                                    tbl   = from_map.get(alias, tbl_part)
                                    fld   = col.get("Property", native)
                                    if tbl:
                                        if tbl not in tables_seen:
                                            tables_seen[tbl] = set()
                                        if fld:
                                            tables_seen[tbl].add(fld)
                                elif agg:
                                    alias = (agg.get("Expression",{}).get("Column",{})
                                               .get("Expression",{}).get("SourceRef",{}).get("Source",""))
                                    tbl   = from_map.get(alias, tbl_part)
                                    fld   = agg.get("Expression",{}).get("Column",{}).get("Property", native)
                                    if tbl:
                                        if tbl not in tables_seen:
                                            tables_seen[tbl] = set()
                                        if fld:
                                            tables_seen[tbl].add(fld)
                        except: pass

                for tbl_name, cols in tables_seen.items():
                    ctx["tables"].append({
                        "name":     tbl_name,
                        "isHidden": False,
                        "columns":  [{"name": c} for c in sorted(cols)],
                    })

        except Exception as e:
            ctx["errors"].append(f"DataModelSchema: {e}")

                # ── Report-level filters ──────────────────────────────────────────────
        try:
            layout = self._read_layout()
            ctx["filters"] = self._safe_parse(layout.get("filters", "[]"))
        except Exception as e:
            ctx["errors"].append(f"Filters: {e}")

        return ctx

    # ── Helpers ───────────────────────────────────────────────────────────────
    @staticmethod
    def _safe_parse(raw):
        if isinstance(raw, (dict, list)):
            return raw
        try:
            return json.loads(raw) if raw else {}
        except:
            return {}

    def apply_theme(self, report_id: str, theme_json: dict) -> bool:
        return False

    def execute_dax(self, dataset_id: str, dax_query: str) -> dict:
        print("DAX execution requires Power BI Service.")
        return {}

    def refresh_dataset(self, dataset_id: str) -> bool:
        return False