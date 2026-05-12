from __future__ import annotations

import json
import os
import socket
import unicodedata
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse

import openpyxl


BASE_DIR = Path(__file__).resolve().parent
DATA_PATH = BASE_DIR / "data.json"
WEEKLY_GOALS_PATH = BASE_DIR / "weekly_goals.json"
HOST = "0.0.0.0"
PORT = int(os.environ.get("PORT", "8000"))
DEFAULT_WEEKLY_GOAL = 700000.0
WORKBOOK_PATTERNS = (
    "BASE PYTHON*.xlsx",
    "BASE PYTHON*.xlsm",
    "COMPROMISSO MES*.xlsx",
    "COMPROMISSO MES*.xlsm",
    "COMPROMISSO MAIO*.xlsx",
    "COMPROMISSO MAIO*.xlsm",
)


INDEX_HTML = r"""<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Acompanhamento Diario</title>
  <style>
    :root{color-scheme:dark;--bg:#05070d;--panel:#111827;--line:#263449;--ink:#f5f7fb;--soft:#a8b3c7;--blue:#38bdf8;--good:#22c55e;--bad:#fb7185}
    *{box-sizing:border-box}body{margin:0;font-family:"Segoe UI",Arial,sans-serif;background:var(--bg);color:var(--ink)}header{padding:20px 28px;border-bottom:1px solid var(--line);display:flex;justify-content:space-between;gap:16px;flex-wrap:wrap}h1{margin:0;font-size:30px}header span,.hint{color:var(--soft);font-weight:700}.pill{padding:8px 12px;border:1px solid var(--line);border-radius:999px;background:#0b1220}main{max-width:1500px;margin:auto;padding:24px 28px}.cards{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:14px}.card,.panel{border:1px solid var(--line);border-radius:8px;background:var(--panel)}.card{padding:18px;min-height:130px}.label{color:var(--soft);font-size:13px;font-weight:900;text-transform:uppercase}.value{margin-top:10px;font-size:32px;font-weight:900;word-break:break-word}.good{color:var(--good)}.bad{color:var(--bad)}.accent{color:var(--blue)}.layout{display:grid;grid-template-columns:minmax(0,1.5fr) minmax(330px,.8fr);gap:16px;margin-top:16px}.panel-header{padding:16px;border-bottom:1px solid var(--line);display:flex;justify-content:space-between;gap:12px;flex-wrap:wrap}h2{margin:0}.controls{display:flex;gap:8px;flex-wrap:wrap}input,select,button{height:40px;border:1px solid var(--line);border-radius:6px;background:#070d18;color:var(--ink);font:inherit;font-weight:700;padding:0 10px}button{background:#2563eb;border:0;cursor:pointer}.table-wrap{max-height:620px;overflow:auto}table{width:100%;border-collapse:collapse;min-width:820px}th,td{padding:12px 14px;border-bottom:1px solid var(--line);text-align:right;white-space:nowrap}th:first-child,td:first-child{text-align:left}th{position:sticky;top:0;background:#0b1220;color:var(--soft);font-size:12px;text-transform:uppercase;cursor:pointer}.seller,.money{font-weight:900}.tag{display:inline-flex;justify-content:center;min-width:80px;padding:6px 10px;border-radius:999px;font-size:12px;font-weight:900}.tag.ok{background:rgba(34,197,94,.15);color:#95f5b6}.tag.pending{background:rgba(251,113,133,.15);color:#ffb2c0}.weekly{display:grid;gap:12px;padding:16px}.week{border:1px solid var(--line);border-radius:8px;padding:14px;background:#090f1a}.week-top{display:flex;justify-content:space-between;font-weight:900}.bar{height:10px;background:#1b2638;border-radius:99px;overflow:hidden;margin:10px 0}.bar span{display:block;height:100%;background:linear-gradient(90deg,#38bdf8,#8b5cf6)}.week-values{display:grid;grid-template-columns:1fr 1fr;gap:8px;color:var(--soft);font-weight:800}.week-values strong{display:block;color:var(--ink)}.error{margin-bottom:16px;padding:14px;border:1px solid var(--bad);border-radius:8px;background:rgba(127,29,29,.35);color:#fecdd3;font-weight:800}@media(max-width:1000px){.cards,.layout{grid-template-columns:1fr 1fr}.layout{display:block}.panel{margin-top:14px}}@media(max-width:640px){header,main{padding-left:14px;padding-right:14px}.cards{grid-template-columns:1fr 1fr}.value{font-size:22px}.controls{display:grid;grid-template-columns:1fr 1fr;width:100%}.controls input{grid-column:1/-1}.controls>*{width:100%}.table-wrap{max-height:none;overflow:visible}thead{display:none}table,tbody,tr,td{display:block;min-width:0;width:100%}tr{margin:10px 0;padding:10px;border:1px solid var(--line);border-radius:8px}td{display:flex;justify-content:space-between;white-space:normal;padding:7px 0}td:before{content:attr(data-label);color:var(--soft);font-size:12px;font-weight:900;text-transform:uppercase}.seller{display:block}.seller:before{display:none}}
  </style>
</head>
<body>
  <header><div><h1>Acompanhamento Diario</h1><span>Relatorio comercial em tempo real</span></div><div><span class="pill" id="workbook">Carregando...</span> <span class="pill" id="updated"></span></div></header>
  <main>
    <div id="error"></div>
    <section class="cards">
      <article class="card"><div class="label">Compromisso</div><div class="value" id="commitment">R$ 0,00</div><div class="hint">Meta consolidada</div></article>
      <article class="card"><div class="label">Atingido</div><div class="value good" id="reached">R$ 0,00</div></article>
      <article class="card"><div class="label">Falta</div><div class="value bad" id="missing">R$ 0,00</div><div class="hint" id="pendingCount"></div></article>
      <article class="card"><div class="label">Realizado</div><div class="value accent" id="percent">0%</div><div class="hint" id="positiveCount"></div></article>
    </section>
    <section class="layout">
      <article class="panel"><div class="panel-header"><h2>Performance por vendedor</h2><div class="controls"><input id="search" type="search" placeholder="Buscar vendedor"><select id="status"><option value="all">Todos</option><option value="pending">Com falta</option><option value="ok">Superou</option></select><button id="refresh">Atualizar</button></div></div><div class="table-wrap"><table><thead><tr><th data-sort="seller">Vendedor</th><th data-sort="commitment">Compromisso</th><th data-sort="reached">Atingido</th><th data-sort="missing">Falta</th><th data-sort="percent">%</th><th>Status</th></tr></thead><tbody id="tableBody"></tbody></table></div></article>
      <aside class="panel"><div class="panel-header"><h2>Acompanhamento semanal</h2><span class="pill" id="weekCount"></span></div><div class="weekly" id="weekly"></div></aside>
    </section>
  </main>
  <script>
    const state={rows:[],weeks:[],sortKey:"missing",sortDir:"desc"};
    const brl=new Intl.NumberFormat("pt-BR",{style:"currency",currency:"BRL"});
    const byId=id=>document.getElementById(id);
    const money=v=>brl.format(Number(v||0));
    const setText=(id,v)=>byId(id).textContent=v;
    async function loadData(){byId("error").innerHTML="";try{const r=await fetch("/api/data?ts="+Date.now());if(!r.ok)throw new Error(await r.text());const data=await r.json();state.rows=data.rows||[];state.weeks=data.weeks||[];renderSummary(data);renderTable();renderWeeks()}catch(e){byId("error").innerHTML=`<div class="error">${e.message}</div>`}}
    function renderSummary(data){const s=data.summary||{};setText("workbook",data.workbook||"data.json");setText("updated","Atualizado em "+(data.lastModified||"-"));setText("commitment",money(s.commitment));setText("reached",money(s.reached));setText("missing",money(s.missing));setText("percent",`${s.percent||0}%`);setText("pendingCount",`${s.pendingCount||0} vendedores pendentes`);setText("positiveCount",`${s.positiveCount||0} acima da meta`)}
    function filteredRows(){const term=byId("search").value.trim().toLowerCase();const status=byId("status").value;return state.rows.filter(row=>!term||String(row.seller||"").toLowerCase().includes(term)).filter(row=>status==="all"||row.status===status).sort((a,b)=>{const av=a[state.sortKey];const bv=b[state.sortKey];const result=typeof av==="string"?av.localeCompare(bv):Number(av)-Number(bv);return state.sortDir==="asc"?result:-result})}
    function renderTable(){byId("tableBody").innerHTML=filteredRows().map(row=>`<tr><td class="seller" data-label="Vendedor">${row.seller}</td><td data-label="Compromisso">${money(row.commitment)}</td><td data-label="Atingido">${money(row.reached)}</td><td data-label="Falta" class="money ${Number(row.missing)<=0?"good":"bad"}">${money(row.missing)}</td><td data-label="%">${row.percent}%</td><td data-label="Status"><span class="tag ${row.status}">${row.status==="ok"?"Superou":"Falta"}</span></td></tr>`).join("")}
    function renderWeeks(){const weeks=state.weeks.slice(0,5);setText("weekCount",`${weeks.length} semanas`);byId("weekly").innerHTML=weeks.map(week=>{const p=Math.min(Number(week.percent||0),100);return`<div class="week"><div class="week-top"><span>${week.name}</span><span class="accent">${Number(week.percent||0).toFixed(1)}%</span></div><div class="bar"><span style="width:${p}%"></span></div><div class="week-values"><span>Realizado<strong>${money(week.reached)}</strong></span><span>Meta<strong>${money(week.goal)}</strong></span></div></div>`}).join("")}
    document.querySelectorAll("th[data-sort]").forEach(h=>h.addEventListener("click",()=>{const key=h.dataset.sort;if(state.sortKey===key){state.sortDir=state.sortDir==="asc"?"desc":"asc"}else{state.sortKey=key;state.sortDir=key==="seller"?"asc":"desc"}renderTable()}));
    byId("search").addEventListener("input",renderTable);byId("status").addEventListener("change",renderTable);byId("refresh").addEventListener("click",loadData);loadData();setInterval(loadData,30000);
  </script>
</body>
</html>
"""


def normalize_text(value):
    return "" if value is None else str(value).strip()


def normalize_key(value):
    text = unicodedata.normalize("NFKD", normalize_text(value).upper())
    return " ".join("".join(ch for ch in text if not unicodedata.combining(ch)).split())


def normalize_week_name(value):
    text = normalize_text(value).title()
    digits = "".join(ch for ch in text if ch.isdigit())
    return f"Semana {int(digits)}" if digits else text or "Semana 1"


def as_number(value):
    if value is None or value == "":
        return 0.0
    if isinstance(value, str) and value.strip().startswith("#"):
        raise ValueError(f"Formula com erro no Excel: {value}")
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def money(value):
    return round(as_number(value), 2)


def week_sort_key(item):
    digits = "".join(ch for ch in item.get("name", "") if ch.isdigit())
    return int(digits) if digits else 99


def default_weekly_goals():
    return {f"Semana {index}": DEFAULT_WEEKLY_GOAL for index in range(1, 6)}


def read_weekly_goals():
    goals = default_weekly_goals()
    if WEEKLY_GOALS_PATH.exists():
        try:
            stored_goals = json.loads(WEEKLY_GOALS_PATH.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            stored_goals = {}
        for name, value in stored_goals.items():
            goals[normalize_week_name(name)] = money(value)
    return goals


def save_weekly_goals(goals):
    normalized_goals = default_weekly_goals()
    for name, value in goals.items():
        normalized_goals[normalize_week_name(name)] = money(value)
    WEEKLY_GOALS_PATH.write_text(
        json.dumps(normalized_goals, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    return normalized_goals


def find_workbook():
    candidates = []
    for pattern in WORKBOOK_PATTERNS:
        candidates.extend(path for path in BASE_DIR.glob(pattern) if not path.name.startswith("~$"))
    return max(candidates, key=lambda path: path.stat().st_mtime) if candidates else None


def apply_weekly_goals(data):
    goals = read_weekly_goals()
    weeks = data.get("weeks") or []
    for index in range(1, 6):
        week_name = f"Semana {index}"
        if not any(normalize_week_name(item.get("name")) == week_name for item in weeks):
            weeks.append({"name": week_name, "goal": 0, "reached": 0, "missing": 0, "percent": 0})
    for week in weeks:
        week_name = normalize_week_name(week.get("name"))
        week["name"] = week_name
        week["goal"] = money(goals.get(week_name, DEFAULT_WEEKLY_GOAL))
        week["reached"] = money(week.get("reached"))
        week["missing"] = money(week["goal"] - week["reached"])
        week["percent"] = round((week["reached"] / week["goal"] * 100) if week["goal"] else 0, 1)
    data["weeks"] = sorted(weeks, key=week_sort_key)[:5]
    summary = data.setdefault("summary", {})
    current_week = data["weeks"][0] if data["weeks"] else {}
    summary["weekGoal"] = current_week.get("goal", 0)
    summary["weekRevenue"] = current_week.get("reached", 0)
    summary["weekMissing"] = current_week.get("missing", 0)
    summary["weekPercent"] = current_week.get("percent", 0)
    return data


def build_row(seller, commitment, reached, missing):
    commitment = money(commitment)
    reached = money(reached)
    missing = money(missing)
    return {
        "seller": seller,
        "commitment": commitment,
        "reached": reached,
        "missing": missing,
        "average": 0.0,
        "difference": 0.0,
        "percent": round((reached / commitment * 100) if commitment else 0, 1),
        "status": "ok" if missing <= 0 else "pending",
    }


def find_block_start(sheet, title):
    wanted = normalize_key(title)
    for column in range(1, sheet.max_column + 1):
        if normalize_key(sheet.cell(1, column).value) == wanted:
            return column
    raise ValueError(f"Bloco nao encontrado em {sheet.title}: {title}")


def make_payload(excel_path, total_commitment, total_reached, total_missing, rows, weeks):
    total_percent = (total_reached / total_commitment * 100) if total_commitment else 0
    return {
        "workbook": excel_path.name,
        "lastModified": datetime.fromtimestamp(excel_path.stat().st_mtime).strftime("%d/%m/%Y %H:%M"),
        "summary": {
            "commitment": money(total_commitment),
            "reached": money(total_reached),
            "missing": money(total_missing),
            "percent": round(total_percent, 1),
            "weekGoal": weeks[0]["goal"] if weeks else 0,
            "weekRevenue": weeks[0]["reached"] if weeks else 0,
            "weekMissing": money(weeks[0]["goal"] - weeks[0]["reached"]) if weeks else 0,
            "weekPercent": round((weeks[0]["reached"] / weeks[0]["goal"] * 100) if weeks and weeks[0]["goal"] else 0, 1),
            "positiveCount": sum(1 for item in rows if item["missing"] <= 0),
            "pendingCount": sum(1 for item in rows if item["missing"] > 0),
        },
        "rows": sorted(rows, key=lambda item: item["missing"], reverse=True),
        "history": [],
        "weeks": weeks,
    }


def read_base_python_planilha1(excel_path, workbook):
    sheet = workbook["Planilha1"]
    start = find_block_start(sheet, "VAREJO")
    seller_col, commitment_col, reached_col, missing_col = start, start + 1, start + 2, start + 3
    rows = []
    for row_number in range(10, sheet.max_row + 1):
        seller = normalize_text(sheet.cell(row_number, seller_col).value)
        if not seller:
            continue
        if normalize_key(seller) in {"SEMANAL", "TOTAL", "VENDEDOR"}:
            break
        rows.append(
            build_row(
                seller,
                sheet.cell(row_number, commitment_col).value,
                sheet.cell(row_number, reached_col).value,
                sheet.cell(row_number, missing_col).value,
            )
        )
    indicators = {}
    for row_number in range(4, 8):
        key = normalize_key(sheet.cell(row_number, seller_col).value)
        if key:
            indicators[key] = money(sheet.cell(row_number, commitment_col).value)
    weeks = []
    weekly_header_row = None
    for row_number in range(10, sheet.max_row + 1):
        if normalize_key(sheet.cell(row_number, seller_col).value) == "SEMANAL":
            weekly_header_row = row_number
            break
    if weekly_header_row:
        for row_number in range(weekly_header_row + 1, sheet.max_row + 1):
            raw_label = normalize_text(sheet.cell(row_number, seller_col).value)
            if not normalize_key(raw_label).startswith("SEMANA"):
                continue
            goal = money(sheet.cell(row_number, commitment_col).value)
            reached = money(sheet.cell(row_number, reached_col).value)
            weeks.append({"name": normalize_week_name(raw_label), "goal": goal, "reached": reached})
    total_commitment = money(indicators.get("META DIA", sum(item["commitment"] for item in rows)))
    total_reached = money(indicators.get("ATINGIDO", sum(item["reached"] for item in rows)))
    total_missing = money(indicators.get("FALTA", total_commitment - total_reached))
    return make_payload(excel_path, total_commitment, total_reached, total_missing, rows, weeks)


def read_base_sheet(excel_path, workbook):
    sheet = workbook["BASE"]
    rows = []
    for row_number in range(1, min(sheet.max_row, 25) + 1):
        headers = {
            normalize_key(sheet.cell(row_number, column).value): column
            for column in range(1, sheet.max_column + 1)
            if normalize_key(sheet.cell(row_number, column).value)
        }
        if {"VENDEDOR", "COMPROMISSO", "ATINGIDO"}.issubset(headers):
            seller_col = headers.get("NOME", headers["VENDEDOR"])
            for item_row in range(row_number + 1, sheet.max_row + 1):
                seller = normalize_text(sheet.cell(item_row, seller_col).value)
                if not seller:
                    continue
                if normalize_key(seller) in {"TOTAL", "VENDEDOR"}:
                    break
                commitment = sheet.cell(item_row, headers["COMPROMISSO"]).value
                reached = sheet.cell(item_row, headers["ATINGIDO"]).value
                missing = sheet.cell(item_row, headers["FALTA"]).value if "FALTA" in headers else money(as_number(commitment) - as_number(reached))
                rows.append(build_row(seller, commitment, reached, missing))
            break
    total_commitment = money(sum(item["commitment"] for item in rows))
    total_reached = money(sum(item["reached"] for item in rows))
    total_missing = money(total_commitment - total_reached)
    weeks = []
    if "RESUMO" in workbook.sheetnames:
        indicators = {}
        for row in workbook["RESUMO"].iter_rows(min_row=1, values_only=True):
            if row:
                indicators[normalize_key(row[0])] = money(row[1] if len(row) > 1 else 0)
        total_commitment = money(indicators.get("META DIA", total_commitment))
        total_reached = money(indicators.get("ATINGIDO", total_reached))
        total_missing = money(indicators.get("FALTA", total_missing))
        for index in range(1, 6):
            weeks.append({"name": f"Semana {index}", "goal": DEFAULT_WEEKLY_GOAL, "reached": money(indicators.get(f"SEMANA {index}", 0))})
    return make_payload(excel_path, total_commitment, total_reached, total_missing, rows, weeks)


def read_workbook_data():
    excel_path = find_workbook()
    if not excel_path:
        return None
    workbook = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)
    try:
        if excel_path.name.upper().startswith("BASE PYTHON") and "Planilha1" in workbook.sheetnames:
            return read_base_python_planilha1(excel_path, workbook)
        if "BASE" in workbook.sheetnames:
            return read_base_sheet(excel_path, workbook)
    finally:
        workbook.close()
    return None


def read_dashboard_data():
    data = read_workbook_data()
    if data:
        return apply_weekly_goals(data)
    if not DATA_PATH.exists():
        raise FileNotFoundError(f"Arquivo nao encontrado: {DATA_PATH}")
    return apply_weekly_goals(json.loads(DATA_PATH.read_text(encoding="utf-8")))


def get_local_addresses():
    addresses = {"127.0.0.1", "localhost"}
    try:
        hostname = socket.gethostname()
        addresses.add(hostname)
        for info in socket.getaddrinfo(hostname, None, family=socket.AF_INET):
            ip = info[4][0]
            if ip and not ip.startswith("127."):
                addresses.add(ip)
    except OSError:
        pass
    return [f"http://{address}:{PORT}" for address in sorted(addresses)]


class DashboardHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == "/":
            self.send_text(INDEX_HTML, "text/html; charset=utf-8")
            return
        if parsed.path == "/api/data":
            try:
                payload = json.dumps(read_dashboard_data(), ensure_ascii=False).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Cache-Control", "no-store")
                self.send_header("Content-Length", str(len(payload)))
                self.end_headers()
                self.wfile.write(payload)
            except Exception as exc:
                self.send_text(str(exc), "text/plain; charset=utf-8", status=500)
            return
        self.send_text("Nao encontrado", "text/plain; charset=utf-8", status=404)

    def do_POST(self):
        parsed = urlparse(self.path)
        if parsed.path == "/api/weekly-goals":
            try:
                length = int(self.headers.get("Content-Length", "0"))
                body = self.rfile.read(length).decode("utf-8") if length else "{}"
                goals = save_weekly_goals(json.loads(body).get("goals", {}))
                payload = json.dumps({"goals": goals}, ensure_ascii=False).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Cache-Control", "no-store")
                self.send_header("Content-Length", str(len(payload)))
                self.end_headers()
                self.wfile.write(payload)
            except Exception as exc:
                self.send_text(str(exc), "text/plain; charset=utf-8", status=500)
            return
        self.send_text("Nao encontrado", "text/plain; charset=utf-8", status=404)

    def log_message(self, format, *args):
        timestamp = datetime.now().strftime("%H:%M:%S")
        print(f"[{timestamp}] {self.address_string()} {format % args}")

    def send_text(self, text, content_type, status=200):
        data = text.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)


def main():
    os.chdir(BASE_DIR)
    server = ThreadingHTTPServer((HOST, PORT), DashboardHandler)
    print("Dashboard compartilhado.")
    for address in get_local_addresses():
        print(f"Acesse: {address}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nEncerrando dashboard.")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
