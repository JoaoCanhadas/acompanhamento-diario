from __future__ import annotations

import json
import os
import socket
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse


BASE_DIR = Path(__file__).resolve().parent
DATA_PATH = BASE_DIR / "data.json"
HOST = "0.0.0.0"
PORT = int(os.environ.get("PORT", "8000"))


INDEX_HTML = r"""<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Acompanhamento Diario</title>
  <style>
    :root {
      color-scheme: dark;
      --bg: #05070d;
      --line: #223047;
      --ink: #f4f7fb;
      --soft: #a8b3c7;
      --muted: #79869d;
      --blue: #38bdf8;
      --blue-2: #2563eb;
      --purple: #8b5cf6;
      --good: #22c55e;
      --bad: #fb7185;
      --shadow: 0 22px 60px rgba(0,0,0,.38);
    }
    * { box-sizing: border-box; }
    body { margin: 0; min-height: 100vh; font-family: "Segoe UI", Arial, sans-serif; color: var(--ink); background: radial-gradient(circle at 20% -10%, rgba(56,189,248,.18), transparent 30%), radial-gradient(circle at 90% 0%, rgba(139,92,246,.18), transparent 34%), var(--bg); overflow-x: hidden; }
    .topbar { border-bottom: 1px solid var(--line); background: rgba(5,7,13,.9); backdrop-filter: blur(18px); }
    .topbar-inner { max-width: 1520px; margin: 0 auto; padding: 20px 28px; display: flex; align-items: center; justify-content: space-between; gap: 18px; flex-wrap: wrap; }
    .title-block h1 { margin: 0; font-size: 30px; line-height: 1.1; font-weight: 800; }
    .title-block span { display: block; margin-top: 6px; color: var(--soft); font-size: 15px; font-weight: 600; }
    .meta { display: flex; gap: 10px; flex-wrap: wrap; justify-content: flex-end; color: var(--soft); font-size: 14px; font-weight: 700; }
    .pill { padding: 9px 12px; border: 1px solid var(--line); border-radius: 999px; background: rgba(17,24,39,.8); }
    main { max-width: 1520px; margin: 0 auto; padding: 24px 28px 34px; }
    .cards { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 16px; margin-bottom: 18px; }
    .card, .panel { border: 1px solid var(--line); border-radius: 8px; background: linear-gradient(180deg, rgba(17,24,39,.96), rgba(12,17,29,.96)); box-shadow: var(--shadow); }
    .card { min-height: 142px; padding: 20px; position: relative; overflow: hidden; }
    .card::after { content: ""; position: absolute; inset: auto 0 0; height: 3px; background: linear-gradient(90deg, var(--blue), var(--purple)); }
    .label { color: var(--soft); font-size: 15px; font-weight: 800; text-transform: uppercase; }
    .value { margin-top: 10px; font-size: 34px; line-height: 1.08; font-weight: 900; word-break: break-word; }
    .hint { margin-top: 10px; color: var(--muted); font-size: 14px; font-weight: 700; }
    .good { color: var(--good); } .bad { color: var(--bad); } .accent { color: var(--blue); }
    .progress { width: 100%; height: 12px; margin-top: 16px; overflow: hidden; border-radius: 999px; background: #1b2638; }
    .progress span { display: block; height: 100%; width: 0; max-width: 100%; border-radius: inherit; background: linear-gradient(90deg, var(--blue), var(--purple)); }
    .layout { display: grid; grid-template-columns: minmax(0, 1.55fr) minmax(360px, .9fr); gap: 18px; align-items: start; }
    .panel-header { padding: 18px 20px; border-bottom: 1px solid var(--line); display: flex; align-items: center; justify-content: space-between; gap: 12px; flex-wrap: wrap; }
    h2 { margin: 0; font-size: 20px; line-height: 1.2; font-weight: 900; }
    .controls { display: flex; gap: 10px; flex-wrap: wrap; }
    input, select, button { height: 42px; border: 1px solid var(--line); border-radius: 6px; background: #080d16; color: var(--ink); font: inherit; font-size: 15px; font-weight: 700; padding: 0 12px; }
    button { cursor: pointer; border-color: transparent; background: linear-gradient(135deg, var(--blue-2), var(--purple)); }
    .table-wrap { overflow-x: auto; }
    table { width: 100%; border-collapse: collapse; min-width: 860px; }
    th, td { padding: 14px 16px; border-bottom: 1px solid rgba(34,48,71,.8); text-align: right; white-space: nowrap; font-size: 16px; }
    th:first-child, td:first-child { text-align: left; }
    th { color: var(--soft); background: rgba(8,13,22,.7); font-size: 13px; text-transform: uppercase; cursor: pointer; user-select: none; }
    tbody tr:hover { background: rgba(56,189,248,.06); }
    .seller { color: #ffffff; font-weight: 900; }
    .money { font-weight: 900; }
    .tag { display: inline-flex; align-items: center; justify-content: center; min-width: 92px; height: 30px; border-radius: 999px; font-size: 13px; font-weight: 900; border: 1px solid; }
    .tag.ok { color: #8ef8b7; background: rgba(34,197,94,.12); border-color: rgba(34,197,94,.38); }
    .tag.pending { color: #ff9aad; background: rgba(251,113,133,.12); border-color: rgba(251,113,133,.38); }
    .weekly { display: grid; gap: 14px; padding: 18px; }
    .week-card { border: 1px solid var(--line); border-radius: 8px; background: rgba(8,13,22,.76); padding: 16px; }
    .week-top { display: flex; align-items: baseline; justify-content: space-between; gap: 12px; margin-bottom: 12px; }
    .week-name { font-size: 17px; font-weight: 900; }
    .week-percent { color: var(--blue); font-size: 22px; font-weight: 900; }
    .week-values { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; color: var(--soft); font-size: 14px; font-weight: 800; }
    .week-values strong { display: block; margin-top: 3px; color: var(--ink); font-size: 18px; }
    .error { margin-bottom: 18px; padding: 16px 18px; border-radius: 8px; border: 1px solid rgba(251,113,133,.5); color: #fecdd3; background: rgba(127,29,29,.4); font-weight: 800; }
    @media (max-width: 1180px) { .cards { grid-template-columns: repeat(2, minmax(0, 1fr)); } .layout { grid-template-columns: 1fr; } }
    @media (max-width: 640px) { .topbar-inner, main { padding-left: 16px; padding-right: 16px; } .title-block h1 { font-size: 25px; } .cards { grid-template-columns: 1fr; } .value { font-size: 30px; } .controls, input { width: 100%; } select, button { flex: 1; } }
  </style>
</head>
<body>
  <header class="topbar"><div class="topbar-inner"><div class="title-block"><h1>Acompanhamento Diario</h1><span>Relatorio comercial em tempo real</span></div><div class="meta"><span class="pill" id="workbook">Carregando...</span><span class="pill" id="updated"></span></div></div></header>
  <main><div id="error"></div>
    <section class="cards">
      <article class="card"><div class="label">Compromisso</div><div class="value" id="commitment">R$ 0,00</div><div class="hint">Meta consolidada do periodo</div></article>
      <article class="card"><div class="label">Atingido</div><div class="value good" id="reached">R$ 0,00</div><div class="progress"><span id="totalProgress"></span></div></article>
      <article class="card"><div class="label">Falta</div><div class="value bad" id="missing">R$ 0,00</div><div class="hint" id="pendingCount">0 vendedores pendentes</div></article>
      <article class="card"><div class="label">Realizado</div><div class="value accent" id="percent">0%</div><div class="hint" id="positiveCount">0 acima da meta</div></article>
    </section>
    <section class="layout">
      <article class="panel"><div class="panel-header"><h2>Performance por vendedor</h2><div class="controls"><input id="search" type="search" placeholder="Buscar vendedor"><select id="status"><option value="all">Todos</option><option value="pending">Com falta</option><option value="ok">Superou</option></select><button id="refresh" type="button">Atualizar</button></div></div><div class="table-wrap"><table><thead><tr><th data-sort="seller">Vendedor</th><th data-sort="commitment">Compromisso</th><th data-sort="reached">Atingido</th><th data-sort="missing">Falta</th><th data-sort="percent">%</th><th>Status</th></tr></thead><tbody id="tableBody"></tbody></table></div></article>
      <aside class="panel"><div class="panel-header"><h2>Acompanhamento semanal</h2><span class="pill" id="weekCount">4 semanas</span></div><div class="weekly" id="weekly"></div></aside>
    </section>
  </main>
  <script>
    const state = { rows: [], weeks: [], sortKey: "missing", sortDir: "desc" };
    const brl = new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL" });
    const byId = (id) => document.getElementById(id);
    const formatMoney = (value) => brl.format(Number(value || 0));
    const moneyClass = (value) => Number(value) <= 0 ? "good" : "bad";
    const setText = (id, value) => { byId(id).textContent = value; };
    async function loadData() { byId("error").innerHTML = ""; try { const response = await fetch("/api/data?ts=" + Date.now()); if (!response.ok) throw new Error(await response.text()); const data = await response.json(); state.rows = data.rows || []; state.weeks = normalizeWeeks(data); renderSummary(data); renderTable(); renderWeeks(); } catch (error) { byId("error").innerHTML = `<div class="error">${error.message}</div>`; } }
    function normalizeWeeks(data) { if (Array.isArray(data.weeks) && data.weeks.length) return data.weeks; const summary = data.summary || {}; const goal = Number(summary.weekGoal || 0); return [1, 2, 3, 4].map((week) => ({ name: `Semana ${week}`, goal, reached: week === 1 ? Number(summary.weekRevenue || 0) : 0, missing: week === 1 ? Number(summary.weekMissing || goal) : goal, percent: week === 1 ? Number(summary.weekPercent || 0) : 0 })); }
    function renderSummary(data) { const summary = data.summary || {}; setText("workbook", data.workbook || "data.json"); setText("updated", "Atualizado em " + (data.lastModified || "-")); setText("commitment", formatMoney(summary.commitment)); setText("reached", formatMoney(summary.reached)); setText("missing", formatMoney(summary.missing)); setText("percent", `${summary.percent || 0}%`); setText("pendingCount", `${summary.pendingCount || 0} vendedores pendentes`); setText("positiveCount", `${summary.positiveCount || 0} acima da meta`); byId("totalProgress").style.width = `${Math.min(Number(summary.percent || 0), 100)}%`; }
    function filteredRows() { const term = byId("search").value.trim().toLowerCase(); const status = byId("status").value; return state.rows.filter((row) => !term || row.seller.toLowerCase().includes(term)).filter((row) => status === "all" || row.status === status).sort((a, b) => { const av = a[state.sortKey]; const bv = b[state.sortKey]; const result = typeof av === "string" ? av.localeCompare(bv) : Number(av) - Number(bv); return state.sortDir === "asc" ? result : -result; }); }
    function renderTable() { byId("tableBody").innerHTML = filteredRows().map((row) => `<tr><td class="seller">${row.seller}</td><td>${formatMoney(row.commitment)}</td><td>${formatMoney(row.reached)}</td><td class="money ${moneyClass(row.missing)}">${formatMoney(row.missing)}</td><td>${row.percent}%</td><td><span class="tag ${row.status}">${row.status === "ok" ? "Superou" : "Falta"}</span></td></tr>`).join(""); }
    function renderWeeks() { const weeks = state.weeks.slice(0, 5); setText("weekCount", `${weeks.length} semanas`); byId("weekly").innerHTML = weeks.map((week) => { const percent = Math.min(Number(week.percent || 0), 100); return `<div class="week-card"><div class="week-top"><div class="week-name">${week.name}</div><div class="week-percent">${Number(week.percent || 0).toFixed(1)}%</div></div><div class="progress"><span style="width:${percent}%"></span></div><div class="week-values"><span>Realizado<strong>${formatMoney(week.reached)}</strong></span><span>Meta<strong>${formatMoney(week.goal)}</strong></span></div></div>`; }).join(""); }
    document.querySelectorAll("th[data-sort]").forEach((header) => { header.addEventListener("click", () => { const key = header.dataset.sort; if (state.sortKey === key) { state.sortDir = state.sortDir === "asc" ? "desc" : "asc"; } else { state.sortKey = key; state.sortDir = key === "seller" ? "asc" : "desc"; } renderTable(); }); });
    byId("search").addEventListener("input", renderTable); byId("status").addEventListener("change", renderTable); byId("refresh").addEventListener("click", loadData); loadData(); setInterval(loadData, 60000);
  </script>
</body>
</html>
"""


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
