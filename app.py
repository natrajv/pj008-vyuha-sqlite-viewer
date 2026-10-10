from flask import Flask, render_template, request, jsonify, abort, send_file
from pathlib import Path
from werkzeug.utils import secure_filename
import sqlite3, json, fnmatch, re, io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

BASE_DIR = Path(__file__).resolve().parent
CONFIG_FILE = BASE_DIR / "config.json"
DEFAULT_CONFIG = {
    "version": "1.3.6", "database": "inventory.db", "page_size": 50,
    "host": "127.0.0.1", "port": 5000, "saved_queries": []
}
ALLOWED_DB_EXTENSIONS = {".db", ".sqlite", ".sqlite3"}
app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 1024 * 1024 * 1024

def load_config():
    try:
        data = json.loads(CONFIG_FILE.read_text(encoding="utf-8")) if CONFIG_FILE.exists() else {}
        return {**DEFAULT_CONFIG, **data}
    except Exception:
        return DEFAULT_CONFIG.copy()

CONFIG = load_config()

def save_config():
    CONFIG_FILE.write_text(json.dumps(CONFIG, indent=4, ensure_ascii=False), encoding="utf-8")

def active_queries():
    groups = CONFIG.setdefault("query_groups", {"Default": []})
    name = CONFIG.get("active_query_group", "Default")
    groups.setdefault(name, CONFIG.get("saved_queries", []))
    CONFIG["saved_queries"] = groups[name]
    return groups[name]

def persist_active_queries():
    CONFIG.setdefault("query_groups", {})[CONFIG.get("active_query_group", "Default")] = CONFIG.get("saved_queries", [])
    save_config()

def db_path():
    p = Path(CONFIG["database"])
    return p if p.is_absolute() else BASE_DIR / p

def connect(path=None):
    target = Path(path) if path else db_path()
    if not target.exists() or not target.is_file():
        raise FileNotFoundError(f"SQLite database not found: {target}")
    # URI read-only mode, plus query_only as defense in depth.
    con = sqlite3.connect(f"file:{target.resolve().as_posix()}?mode=ro", uri=True)
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA query_only = ON")
    con.execute("PRAGMA busy_timeout = 5000")
    return con

def quote_ident(name):
    return '"' + name.replace('"', '""') + '"'

def table_names():
    with connect() as con:
        return [r["name"] for r in con.execute("""
            SELECT name FROM sqlite_master WHERE type='table'
            AND name NOT LIKE 'sqlite_%' ORDER BY name COLLATE NOCASE
        """).fetchall()]

def is_safe_select(sql):
    # SQLite connection is read-only; also reject multiple statements and non-query verbs.
    stripped = sql.strip()
    if not stripped:
        return False, "Enter a SQL query."
    stripped = stripped.rstrip(";").strip()
    if ";" in stripped:
        return False, "Only one SQL statement is allowed."
    if not re.match(r"(?is)^(SELECT|WITH|EXPLAIN)\b", stripped):
        return False, "Only SELECT, WITH (read-only CTE), or EXPLAIN queries are allowed."
    return True, stripped

def wildcard_to_like(value):
    return value.replace("\\", "\\\\").replace("%", r"\%").replace("_", r"\_").replace("*", "%").replace("?", "_")

def get_query_rows(sql, params=()):
    ok, normalized = is_safe_select(sql)
    if not ok:
        raise ValueError(normalized)
    with connect() as con:
        cur = con.execute(normalized, params)
        cols = [d[0] for d in (cur.description or [])]
        rows = cur.fetchall()
        return cols, [{k: row[k] for k in row.keys()} for row in rows]

def apply_search(columns, rows, search="", mode="wildcard"):
    if not search:
        return rows
    if mode == "regex":
        try:
            pattern = re.compile(search, re.IGNORECASE)
        except re.error as exc:
            raise ValueError(f"Invalid regular expression: {exc}")
        return [r for r in rows if any(pattern.search("" if r.get(c) is None else str(r.get(c))) for c in columns)]
    pattern = search.lower()
    return [r for r in rows if any(fnmatch.fnmatchcase("" if r.get(c) is None else str(r.get(c)).lower(), pattern) for c in columns)]

@app.route("/")
def index():
    error = None
    try:
        tables = table_names()
    except Exception as exc:
        tables, error = [], str(exc)
    return render_template("index.html", tables=tables, config=CONFIG,
                           db_name=db_path().name, db_exists=db_path().exists(), error=error)

@app.post("/api/open-database")
def open_database():
    uploaded = request.files.get("database")
    if not uploaded or not uploaded.filename:
        return jsonify(error="Choose a SQLite database file."), 400
    filename = secure_filename(uploaded.filename)
    if Path(filename).suffix.lower() not in ALLOWED_DB_EXTENSIONS:
        return jsonify(error="Choose a .db, .sqlite, or .sqlite3 file."), 400
    folder = BASE_DIR / "databases"
    folder.mkdir(exist_ok=True)
    destination = folder / filename
    # Avoid overwriting the active database file in place.
    uploaded.save(destination)
    try:
        with connect(destination) as con:
            con.execute("SELECT name FROM sqlite_master LIMIT 1").fetchone()
    except Exception as exc:
        try: destination.unlink(missing_ok=True)
        except Exception: pass
        return jsonify(error=f"File is not a readable SQLite database: {exc}"), 400
    CONFIG["database"] = str(destination.relative_to(BASE_DIR))
    save_config()
    return jsonify(database=CONFIG["database"], name=destination.name, tables=table_names())

@app.get("/api/tables")
def api_tables():
    return jsonify(tables=table_names(), database=db_path().name)

@app.get("/api/table/<path:table>")
def api_table(table):
    if table not in table_names():
        abort(404, description="Table not found")
    try:
        page = max(1, int(request.args.get("page", 1)))
        page_size = max(1, min(500, int(request.args.get("page_size", CONFIG["page_size"]))))
    except ValueError:
        page, page_size = 1, CONFIG["page_size"]
    search = request.args.get("search", "")
    mode = request.args.get("mode", "wildcard")
    sort = request.args.get("sort", "")
    direction = "DESC" if request.args.get("direction", "asc").lower() == "desc" else "ASC"

    with connect() as con:
        info = con.execute(f"PRAGMA table_info({quote_ident(table)})").fetchall()
        columns = [r["name"] for r in info]
        if sort and sort not in columns and sort != "__rowid__":
            return jsonify(error="Invalid sort column"), 400
        order = f" ORDER BY {quote_ident('rowid' if sort == '__rowid__' else sort)} {direction}" if sort else ""
        all_rows = con.execute(f"SELECT rowid AS __rowid__, * FROM {quote_ident(table)}{order}").fetchall()
        rows = [{k: row[k] for k in row.keys()} for row in all_rows]
    all_columns = ["__rowid__"] + columns
    try:
        filtered = apply_search(all_columns, rows, search, mode)
    except ValueError as exc:
        return jsonify(error=str(exc)), 400
    total = len(filtered)
    pages = max(1, (total + page_size - 1) // page_size)
    page = min(page, pages)
    return jsonify(table=table, columns=all_columns,
                   rows=filtered[(page-1)*page_size:page*page_size],
                   page=page, page_size=page_size, total=total, pages=pages)

@app.post("/api/query")
def api_query():
    payload = request.get_json(silent=True) or {}
    sql = payload.get("sql", "")
    search = payload.get("search", "")
    mode = payload.get("mode", "wildcard")
    try:
        columns, rows = get_query_rows(sql)
        rows = apply_search(columns, rows, search, mode)
        return jsonify(columns=columns, rows=rows, total=len(rows), query=sql)
    except (ValueError, sqlite3.Error, FileNotFoundError) as exc:
        return jsonify(error=str(exc)), 400

@app.get("/api/saved-queries")
def get_saved_queries():
    return jsonify(queries=active_queries(), groups=sorted(CONFIG.setdefault("query_groups", {"Default": []}).keys(), key=str.casefold), active_group=CONFIG.get("active_query_group", "Default"))

@app.post("/api/saved-queries")
def save_saved_query():
    payload=request.get_json(silent=True) or {}
    name=str(payload.get("name","")).strip()
    sql=str(payload.get("sql","")).strip()
    group_name=str(payload.get("group", CONFIG.get("active_query_group", "Default"))).strip()
    if not name or not sql: return jsonify(error="Query name and SQL are required."),400
    if not group_name: return jsonify(error="Select a query group."),400
    ok,msg=is_safe_select(sql)
    if not ok: return jsonify(error=msg),400
    groups=CONFIG.setdefault("query_groups", {"Default": []})
    # Resolve an existing group case-insensitively; create it only if explicitly supplied and new.
    existing=next((g for g in groups if g.casefold()==group_name.casefold()),None)
    if existing is not None:
        group_name=existing
    else:
        groups[group_name]=[]
    CONFIG["active_query_group"]=group_name
    CONFIG["saved_queries"]=groups[group_name]
    qs=active_queries()
    old=next((q for q in qs if q["name"].casefold()==name.casefold()),None)
    if old: old["sql"]=sql
    else: qs.append({"name":name,"sql":sql})
    persist_active_queries()
    return jsonify(queries=qs, groups=sorted(groups,key=str.casefold), active_group=group_name)


@app.put("/api/saved-queries")
def reorder_saved_queries():
    payload=request.get_json(silent=True) or {}; names=payload.get("names",[]); qs=active_queries(); by={q["name"]:q for q in qs}
    ordered=[by[n] for n in names if n in by]; ordered.extend(q for q in qs if q["name"] not in names); CONFIG["saved_queries"]=ordered; persist_active_queries(); return jsonify(queries=ordered)

@app.delete("/api/saved-queries/<path:name>")
def delete_saved_query(name):
    CONFIG["saved_queries"]=[q for q in active_queries() if q["name"]!=name]; persist_active_queries(); return jsonify(queries=CONFIG["saved_queries"])

@app.post("/api/query-groups")
def create_query_group():
    payload=request.get_json(silent=True) or {}; name=str(payload.get("name","")).strip()
    if not name: return jsonify(error="Enter a group name."),400
    groups=CONFIG.setdefault("query_groups",{"Default":[]}); old=next((g for g in groups if g.casefold()==name.casefold()),None)
    if old: name=old
    else: groups[name]=[]
    CONFIG["active_query_group"]=name; CONFIG["saved_queries"]=groups[name]; save_config()
    return jsonify(groups=sorted(groups,key=str.casefold),active_group=name,queries=groups[name])

@app.put("/api/query-groups/active")
def activate_query_group():
    payload=request.get_json(silent=True) or {}; name=str(payload.get("name","")).strip(); groups=CONFIG.setdefault("query_groups",{"Default":[]})
    if name not in groups: return jsonify(error="Saved Queries group not found."),404
    CONFIG["active_query_group"]=name; CONFIG["saved_queries"]=groups[name]; save_config()
    return jsonify(groups=sorted(groups,key=str.casefold),active_group=name,queries=groups[name])

@app.put("/api/query-groups/rename")
def rename_query_group():
    payload=request.get_json(silent=True) or {}
    old_name=str(payload.get("old_name","")).strip()
    new_name=str(payload.get("new_name","")).strip()
    groups=CONFIG.setdefault("query_groups", {"Default": []})
    if not old_name or old_name not in groups:
        return jsonify(error="The selected query group was not found."),404
    if not new_name:
        return jsonify(error="Enter the new query group name."),400
    duplicate=next((g for g in groups if g.casefold()==new_name.casefold() and g!=old_name),None)
    if duplicate:
        return jsonify(error=f'A query group named "{duplicate}" already exists.'),400
    if old_name == new_name:
        return jsonify(error="The new group name is the same as the current name."),400
    groups[new_name]=groups.pop(old_name)
    if CONFIG.get("active_query_group", "Default")==old_name:
        CONFIG["active_query_group"]=new_name
    CONFIG["saved_queries"]=groups.get(CONFIG.get("active_query_group", "Default"), [])
    save_config()
    return jsonify(groups=sorted(groups,key=str.casefold),
                   active_group=CONFIG.get("active_query_group", "Default"),
                   queries=CONFIG["saved_queries"])

@app.post("/api/export")
def export_excel():
    payload = request.get_json(silent=True) or {}
    columns = payload.get("columns", [])
    rows = payload.get("rows", [])
    sheet_name = str(payload.get("sheet", "SQLite Data"))[:31] or "SQLite Data"
    if not columns:
        return jsonify(error="There is no grid data to export."), 400
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws.append(columns)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="145B86")
    for row in rows:
        ws.append([row.get(c) for c in columns])
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for idx, col in enumerate(columns, start=1):
        max_len = max([len(str(ws.cell(r, idx).value or "")) for r in range(1, min(ws.max_row, 200)+1)] + [len(str(col))])
        ws.column_dimensions[get_column_letter(idx)].width = min(max(max_len + 2, 10), 55)
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name="vyuha-sqlite-export.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.get("/api/row/<path:table>/<int:rowid>")
def api_row(table, rowid):
    if table not in table_names():
        abort(404, description="Table not found")
    with connect() as con:
        row = con.execute(f"SELECT rowid AS __rowid__, * FROM {quote_ident(table)} WHERE rowid=?", (rowid,)).fetchone()
    if row is None:
        abort(404, description="Row not found")
    return jsonify(table=table, row={k: row[k] for k in row.keys()})

if __name__ == "__main__":
    app.run(host=CONFIG["host"], port=int(CONFIG["port"]), debug=False)
