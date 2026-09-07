"""
Recherche de produits dans les fichiers Excel de Magda (par marque, ref, etc.)

Comment ca marche :
- Au demarrage / a chaque recherche, l'outil relit UNIQUEMENT les fichiers
  Excel du dossier ci-dessous qui sont nouveaux ou modifies depuis la
  derniere fois (grace a leur date de modification). Les fichiers supprimes
  disparaissent automatiquement des resultats. Pas besoin de reindexer
  manuellement quand tu ajoutes/enleves des fichiers.
- La recherche se fait sur TOUTES les cellules de TOUS les onglets de
  TOUS les fichiers (peu importe la structure de chaque fichier).

Pour lancer : python app.py
Puis ouvrir http://localhost:5000 (ou http://<IP-de-ce-PC>:5000 depuis un
autre poste du reseau).
"""

from __future__ import annotations

import datetime
import html
import os
import threading
from pathlib import Path

import openpyxl
import xlrd
from flask import Flask, render_template_string, request

# ---------------------------------------------------------------------------
# Configuration : a adapter si besoin
# ---------------------------------------------------------------------------
SEARCH_FOLDER = Path(r"X:\Purchase Deco\leveranciers\lotenleveranciers\Bestellijsten Magda")
INCLUDE_SUBFOLDERS = False  # racine uniquement, cf. dossier "oude lijsten"
EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}
MAX_RESULTS = 500
HOST = "0.0.0.0"
PORT = 5000

app = Flask(__name__)

# ---------------------------------------------------------------------------
# Index en memoire : { chemin_fichier: {"mtime": float, "sheets": [...], "error": str|None} }
# ---------------------------------------------------------------------------
_index: dict[str, dict] = {}
_index_lock = threading.Lock()
_last_scan: datetime.datetime | None = None


def _cell_to_str(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return f"{value:.2f}"
    if isinstance(value, (datetime.datetime, datetime.date)):
        return value.strftime("%Y-%m-%d")
    return str(value)


def _parse_xlsx(path: Path) -> list[dict]:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        sheets = []
        for ws in wb.worksheets:
            header = None
            rows = []
            for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
                values = [_cell_to_str(v) for v in row]
                if not any(v.strip() for v in values):
                    continue
                if header is None:
                    header = values  # 1ere ligne non vide = en-tetes de colonnes
                rows.append((row_idx, values))
            sheets.append({"name": ws.title, "header": header, "rows": rows})
        return sheets
    finally:
        wb.close()


def _parse_xls(path: Path) -> list[dict]:
    wb = xlrd.open_workbook(str(path))
    sheets = []
    for ws in wb.sheets():
        header = None
        rows = []
        for row_idx in range(ws.nrows):
            values = [_cell_to_str(v) for v in ws.row_values(row_idx)]
            if not any(v.strip() for v in values):
                continue
            if header is None:
                header = values
            rows.append((row_idx + 1, values))
        sheets.append({"name": ws.name, "header": header, "rows": rows})
    return sheets


def _list_excel_files() -> list[Path]:
    if not SEARCH_FOLDER.exists():
        return []
    pattern = "**/*" if INCLUDE_SUBFOLDERS else "*"
    files = []
    for p in SEARCH_FOLDER.glob(pattern):
        if not p.is_file():
            continue
        if p.name.startswith("~$"):  # fichier verrouille/temporaire Excel
            continue
        if p.suffix.lower() in EXCEL_EXTENSIONS:
            files.append(p)
    return files


def update_index(force: bool = False) -> None:
    """Ajoute les fichiers nouveaux/modifies, retire ceux qui ont disparu."""
    global _last_scan
    with _index_lock:
        current_files = _list_excel_files()
        current_paths = {str(p) for p in current_files}

        # retirer les fichiers qui n'existent plus
        for stale in set(_index) - current_paths:
            del _index[stale]

        for path in current_files:
            key = str(path)
            try:
                mtime = path.stat().st_mtime
            except OSError:
                continue

            cached = _index.get(key)
            if not force and cached and cached["mtime"] == mtime:
                continue  # deja a jour

            try:
                if path.suffix.lower() == ".xls":
                    sheets = _parse_xls(path)
                else:
                    sheets = _parse_xlsx(path)
                _index[key] = {"mtime": mtime, "sheets": sheets, "error": None}
            except Exception as exc:  # fichier corrompu, ouvert en exclusif, etc.
                _index[key] = {"mtime": mtime, "sheets": [], "error": str(exc)}

        _last_scan = datetime.datetime.now()


def search(query: str) -> tuple[list[dict], bool, list[str]]:
    query_low = query.lower().strip()
    results = []
    truncated = False
    errors = []

    with _index_lock:
        items = list(_index.items())

    for path, data in items:
        if data["error"]:
            errors.append(f"{Path(path).name} : {data['error']}")
            continue
        for sheet in data["sheets"]:
            for row_idx, values in sheet["rows"]:
                joined = " | ".join(values)
                if query_low in joined.lower():
                    if len(results) >= MAX_RESULTS:
                        truncated = True
                        break
                    results.append(
                        {
                            "file": Path(path).name,
                            "full_path": path,
                            "sheet": sheet["name"],
                            "row": row_idx,
                            "values": values,
                            "header": sheet["header"],
                        }
                    )
            if truncated:
                break
        if truncated:
            break

    return results, truncated, errors


def _format_row(header: list[str] | None, values: list[str]) -> str:
    """Associe chaque valeur non vide a son en-tete de colonne, ex: 'Prijs: 1,35'."""
    parts = []
    if header:
        for h, v in zip(header, values):
            v = v.strip()
            if not v:
                continue
            label = (h or "").strip()
            parts.append(f"{label}: {v}" if label else v)
        # colonnes en trop si la ligne est plus large que l'en-tete
        for v in values[len(header):]:
            v = v.strip()
            if v:
                parts.append(v)
    else:
        parts = [v for v in values if v.strip()]
    return " | ".join(parts)


def _highlight(text: str, query: str) -> str:
    escaped = html.escape(text)
    if not query:
        return escaped
    q_esc = html.escape(query)
    idx_low = escaped.lower().find(q_esc.lower())
    if idx_low == -1:
        return escaped
    return (
        escaped[:idx_low]
        + "<mark>"
        + escaped[idx_low : idx_low + len(q_esc)]
        + "</mark>"
        + escaped[idx_low + len(q_esc) :]
    )


PAGE = """
<!doctype html>
<html lang="fr">
<head>
<meta charset="utf-8">
<title>Recherche produits - Bestellijsten Magda</title>
<style>
  body { font-family: system-ui, Arial, sans-serif; margin: 2rem; background: #f7f7f5; color: #1f1f1f; }
  h1 { font-size: 1.4rem; margin-bottom: .25rem; }
  .meta { color: #666; font-size: .85rem; margin-bottom: 1.25rem; }
  form { display: flex; gap: .5rem; margin-bottom: 1.5rem; }
  input[type=text] { flex: 1; padding: .6rem .8rem; font-size: 1rem; border: 1px solid #ccc; border-radius: 6px; }
  button { padding: .6rem 1.1rem; font-size: 1rem; border: 0; border-radius: 6px; background: #2f7a3d; color: white; cursor: pointer; }
  button:hover { background: #256030; }
  button.secondary { background: #888; }
  button.secondary:hover { background: #666; }
  table { border-collapse: collapse; width: 100%; background: white; box-shadow: 0 1px 3px rgba(0,0,0,.08); }
  th, td { border-bottom: 1px solid #eee; padding: .5rem .6rem; text-align: left; font-size: .88rem; vertical-align: top; }
  th { background: #eef3ec; position: sticky; top: 0; }
  tr:hover { background: #fafcf8; }
  mark { background: #ffe58a; padding: 0 2px; }
  .count { margin-bottom: .75rem; font-size: .9rem; color: #333; }
  .warn { background: #fff4e5; border: 1px solid #f0c987; padding: .6rem .8rem; border-radius: 6px; margin-bottom: 1rem; font-size: .85rem; }
  .path { color: #888; font-size: .78rem; }
</style>
</head>
<body>
  <h1>Recherche produits &mdash; Bestellijsten Magda</h1>
  <div class="meta">
    {{ file_count }} fichiers indexes{% if last_scan %} &middot; derniere mise a jour de l'index : {{ last_scan }}{% endif %}
  </div>

  <form method="get" action="/">
    <input type="text" name="q" placeholder="Marque, reference, mot-cle..." value="{{ query|e }}" autofocus>
    <button type="submit">Rechercher</button>
    <button type="submit" name="refresh" value="1" class="secondary" title="Force une relecture complete de tous les fichiers">Reindexer</button>
  </form>

  {% if errors %}
  <div class="warn">
    <strong>{{ errors|length }} fichier(s) non lisible(s)</strong> (probablement ouverts dans Excel ou corrompus) :
    <ul>
      {% for e in errors %}<li>{{ e }}</li>{% endfor %}
    </ul>
  </div>
  {% endif %}

  {% if query %}
    <div class="count">
      {{ results|length }} resultat(s){% if truncated %} (affichage limite aux {{ max_results }} premiers, affine ta recherche){% endif %}
    </div>
    {% if results %}
    <table>
      <thead><tr><th>Fichier</th><th>Onglet</th><th>Ligne</th><th>Contenu</th></tr></thead>
      <tbody>
      {% for r in results %}
        <tr>
          <td>{{ r.file }}<div class="path">{{ r.full_path }}</div></td>
          <td>{{ r.sheet }}</td>
          <td>{{ r.row }}</td>
          <td>{{ r.content|safe }}</td>
        </tr>
      {% endfor %}
      </tbody>
    </table>
    {% endif %}
  {% endif %}
</body>
</html>
"""


@app.route("/")
def index_route():
    query = request.args.get("q", "").strip()
    force = request.args.get("refresh") == "1"

    update_index(force=force)

    results, truncated, errors = ([], False, [])
    if query:
        results, truncated, errors = search(query)
        for r in results:
            r["content"] = _highlight(_format_row(r["header"], r["values"]), query)

    return render_template_string(
        PAGE,
        query=query,
        results=results,
        truncated=truncated,
        max_results=MAX_RESULTS,
        errors=errors,
        file_count=len(_index),
        last_scan=_last_scan.strftime("%d/%m/%Y %H:%M") if _last_scan else None,
    )


if __name__ == "__main__":
    print(f"Indexation initiale de {SEARCH_FOLDER} ...")
    update_index(force=True)
    print(f"{len(_index)} fichiers indexes. Demarrage du serveur sur http://{HOST}:{PORT}")
    app.run(host=HOST, port=PORT, debug=False)
