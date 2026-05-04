from __future__ import annotations

import json
import os
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse

BASE_DIR = Path(__file__).resolve().parent
DATA_PATH = BASE_DIR / "data.json"
HOST = "0.0.0.0"
PORT = int(os.environ.get("PORT", "8000"))

INDEX_HTML = """<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Acompanhamento Diario</title>
  <style>
    :root{--bg:#f5f7f9;--panel:#fff;--ink:#17202a;--muted:#637083;--line:#dce2e8;--good:#16834a;--bad:#c62828;--blue:#2166a6;--teal:#177e89;--shadow:0 12px 30px rgba(31,42,55,.08)}
    *{box-sizing:border-box}body{margin:0;font-family:Segoe UI,Arial,sans-serif;color:var(--ink);background:var(--bg)}
    .topbar{background:#101820;color:#fff;padding:18px 24px;border-bottom:4px solid var(--teal)}
    .topbar-inner{max-width:1220px;margin:0 auto;display:flex;align-items:center;justify-content:space-between;gap:16px;flex-wrap:wrap}
    h1{margin:0;font-size:24px}.meta{color:#c8d1dc;font-size:13px;display:flex;gap:12px;flex-wrap:wrap}
    main{max-width:1220px;margin:0 auto;padding:22px 18px 36px}.cards{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:14px;margin-bottom:18px}
    .card,.panel{background:var(--panel);border:1px solid var(--line);border-radius:8px;box-shadow:var(--shadow)}.card{padding:16px;min-height:112px}
    .label{color:var(--muted);font-size:12px;text-transform:uppercase;font-weight:700}.value{margin-top:8px;font-size:26px;font-weight:800;line-height:1.15;word-break:break-word}
    .good{color:var(--good)}.bad{color:var(--bad)}.progress{width:100%;height:10px;margin-top:12px;overflow:hidden;border-radius:99px;background:#e7ecf1}.progress span{display:block;height:100%;width:0;max-width:100%;background:linear-gradient(90deg,var(--teal),var(--blue))}
    .grid{display:grid;grid-template-columns:minmax(0,1.6fr) minmax(320px,.9fr);gap:16px;align-items:start}.panel-header{padding:14px 16px;border-bottom:1px solid var(--line);display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap}h2{margin:0;font-size:16px}
    .controls{display:flex;gap:8px;flex-wrap:wrap}input,select,button{height:36px;border:1px solid var(--line);border-radius:6px;background:#fff;color:var(--ink);font:inherit;padding:0 10px}button{cursor:pointer;font-weight:700;color:#fff;background:var(--blue);border-color:var(--blue)}
    .table-wrap{overflow-x:auto}table{width:100%;border-collapse:collapse;min-width:760px}th,td{padding:11px 12px;border-bottom:1px solid var(--line);text-align:right;white-space:nowrap}th:first-child,td:first-child{text-align:left}th{font-size:12px;text-transform:uppercase;color:var(--muted);background:#fbfcfd;cursor:pointer}tbody tr:hover{background:#f7fafc}.seller{font-weight:800}.money{font-weight:800}
    .tag{display:inline-flex;align-items:center;justify-content:center;min-width:82px;height:26px;border-radius:99px;font-size:12px;font-weight:800;border:1px solid}.tag.ok{color:var(--good);background:#eaf7ef;border-color:#c7e7d1}.tag.pending{color:var(--bad);background:#fff0f0;border-color:#f3c6c6}
    .side{display:grid;gap:16px}.ranking,.history{padding:14px 16px 16px}.rank-item{display:grid;grid-template-columns:1fr auto;gap:10px;padding:10px 0;border-bottom:1px solid var(--line)}.rank-item:last-child{border-bottom:0}.rank-name{font-weight:800}.rank-percent{color:var(--muted);font-weight:800}.bar{grid-column:1/-1;height:8px;background:#e9eef3;border-radius:99px;overflow:hidden}.bar span{display:block;height:100%;background:var(--teal);max-width:100%}
    .history-row{display:grid;grid-template-columns:82px 1fr auto;gap:10px;align-items:center;padding:9px 0;border-bottom:1px solid var(--line);font-size:13px}.history-row:last-child{border-bottom:0}.muted{color:var(--muted)}.error{margin:18px 0;padding:14px 16px;border-radius:8px;border:1px solid #f2b8b8;color:var(--bad);background:#fff3f3;font-weight:700}
    @media(max-width:980px){.cards{grid-template-columns:repeat(2,minmax(0,1fr))}.grid{grid-template-columns:1fr}}@media(max-width:560px){.topbar{padding:16px}main{padding:16px 12px 28px}.cards{grid-template-columns:1fr}.value{font-size:23px}input{width:100%}.controls{width:100%}select,button{flex:1}}
  </style>
</head>
<body>
  <header class="topbar"><div class="topbar-inner"><h1>Acompanhamento Diario</h1><div class="meta"><span id="workbook">Carregando...</span><span id="updated"></span></div></div></header>
  <main><div id="error"></div><section class="cards"><article class="card"><div class="label">Compromisso</div><div class="value" id="commitment">R$ 0,00</div></article><article class="card"><div class="label">Atingido</div><div class="value good" id="reached">R$ 0,00</div><div class="progress"><span id="totalProgress"></span></div></article><article class="card"><div class="label">Falta</div><div class="value bad" id="missing">R$ 0,00</div></article><article class="card"><div class="label">Semana 1</div><div class="value" id="week">0%</div><div class="progress"><span id="weekProgress"></span></div></article></section>
  <section class="grid"><article class="panel"><div class="panel-header"><h2>Vendedores</h2><div class="controls"><input id="search" type="search" placeholder="Buscar vendedor"><select id="status"><option value="all">Todos</option><option value="pending">Com falta</option><option value="ok">Superou</option></select><button id="refresh" type="button">Atualizar</button></div></div><div class="table-wrap"><table><thead><tr><th data-sort="seller">Vendedor</th><th data-sort="commitment">Compromisso</th><th data-sort="reached">Atingido</th><th data-sort="missing">Falta</th><th data-sort="percent">%</th><th>Status</th></tr></thead><tbody id="tableBody"></tbody></table></div></article><aside class="side"><article class="panel"><div class="panel-header"><h2>Melhores desempenhos</h2></div><div class="ranking" id="ranking"></div></article><article class="panel"><div class="panel-header"><h2>Movimento diario</h2></div><div class="history" id="history"></div></article></aside></section></main>
  <script>
    const state={rows:[],sortKey:"missing",sortDir:"desc"};const brl=new Intl.NumberFormat("pt-BR",{style:"currency",currency:"BRL"});const byId=id=>document.getElementById(id);const formatMoney=v=>brl.format(Number(v||0));const moneyClass=v=>Number(v)<=0?"good":"bad";const setText=(id,v)=>byId(id).textContent=v;
    async function loadData(){byId("error").innerHTML="";try{const r=await fetch("/api/data?ts="+Date.now());if(!r.ok)throw new Error(await r.text());const d=await r.json();state.rows=d.rows;renderSummary(d);renderTable();renderRanking();renderHistory(d.history)}catch(e){byId("error").innerHTML=`<div class="error">${e.message}</div>`}}
    function renderSummary(d){const s=d.summary;setText("workbook",d.workbook);setText("updated","Atualizado em "+d.lastModified);setText("commitment",formatMoney(s.commitment));setText("reached",formatMoney(s.reached));setText("missing",formatMoney(s.missing));setText("week",`${s.weekPercent}%`);byId("totalProgress").style.width=`${Math.min(s.percent,100)}%`;byId("weekProgress").style.width=`${Math.min(s.weekPercent,100)}%`}
    function filteredRows(){const term=byId("search").value.trim().toLowerCase();const status=byId("status").value;return state.rows.filter(r=>!term||r.seller.toLowerCase().includes(term)).filter(r=>status==="all"||r.status===status).sort((a,b)=>{const av=a[state.sortKey],bv=b[state.sortKey];const res=typeof av==="string"?av.localeCompare(bv):Number(av)-Number(bv);return state.sortDir==="asc"?res:-res})}
    function renderTable(){byId("tableBody").innerHTML=filteredRows().map(r=>`<tr><td class="seller">${r.seller}</td><td>${formatMoney(r.commitment)}</td><td>${formatMoney(r.reached)}</td><td class="money ${moneyClass(r.missing)}">${formatMoney(r.missing)}</td><td>${r.percent}%</td><td><span class="tag ${r.status}">${r.status==="ok"?"Superou":"Falta"}</span></td></tr>`).join("")}
    function renderRanking(){const top=[...state.rows].sort((a,b)=>b.percent-a.percent).slice(0,6);byId("ranking").innerHTML=top.map(r=>`<div class="rank-item"><div class="rank-name">${r.seller}</div><div class="rank-percent">${r.percent}%</div><div class="bar"><span style="width:${Math.min(r.percent,100)}%"></span></div></div>`).join("")}
    function renderHistory(h){byId("history").innerHTML=h.slice(-8).reverse().map(i=>`<div class="history-row"><strong>${i.date}</strong><span class="muted">${formatMoney(i.reached)} atingido</span><span class="money ${moneyClass(i.balance)}">${formatMoney(i.balance)}</span></div>`).join("")}
    document.querySelectorAll("th[data-sort]").forEach(h=>h.addEventListener("click",()=>{const k=h.dataset.sort;if(state.sortKey===k){state.sortDir=state.sortDir==="asc"?"desc":"asc"}else{state.sortKey=k;state.sortDir=k==="seller"?"asc":"desc"}renderTable()}));byId("search").addEventListener("input",renderTable);byId("status").addEventListener("change",renderTable);byId("refresh").addEventListener("click",loadData);loadData();setInterval(loadData,60000);
  </script>
</body>
</html>"""


def read_dashboard_data():
    if not DATA_PATH.exists():
        raise FileNotFoundError(f"Arquivo nao encontrado: {DATA_PATH}")
    return json.loads(DATA_PATH.read_text(encoding="utf-8"))


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

    def log_message(self, format, *args):
        print(format % args)

    def send_text(self, text, content_type, status=200):
        data = text.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)


def main():
    server = ThreadingHTTPServer((HOST, PORT), DashboardHandler)
    print(f"Dashboard no ar em http://{HOST}:{PORT}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("Encerrando dashboard.")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
