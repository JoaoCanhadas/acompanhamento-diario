from __future__ import annotations

import json
import os
import socket
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse

import openpyxl


BASE_DIR = Path(__file__).resolve().parent
EXCEL_PATH = BASE_DIR / "COMPROMISSO MAIO.xlsm"
LEGACY_EXCEL_PATH = BASE_DIR / "base_vendas.xlsx"
DATA_PATH = BASE_DIR / "data.json"
WEEKLY_GOALS_PATH = BASE_DIR / "weekly_goals.json"
HOST = "0.0.0.0"
PORT = int(os.environ.get("PORT", "8000"))
DEFAULT_WEEKLY_GOAL = 700000.0


INDEX_HTML = r"""<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Acompanhamento Diario</title>
  <style>
    :root{color-scheme:dark;--bg:#05070d;--panel:#0c111d;--line:#223047;--ink:#f4f7fb;--soft:#a8b3c7;--muted:#79869d;--blue:#38bdf8;--blue-2:#2563eb;--purple:#8b5cf6;--good:#22c55e;--bad:#fb7185;--shadow:0 22px 60px rgba(0,0,0,.38)}
    *{box-sizing:border-box}body{margin:0;min-height:100vh;font-family:"Segoe UI",Arial,sans-serif;color:var(--ink);background:radial-gradient(circle at 20% -10%,rgba(56,189,248,.18),transparent 30%),radial-gradient(circle at 90% 0%,rgba(139,92,246,.18),transparent 34%),var(--bg);overflow-x:hidden}.topbar{border-bottom:1px solid var(--line);background:rgba(5,7,13,.9);backdrop-filter:blur(18px)}.topbar-inner{max-width:1520px;margin:0 auto;padding:20px 28px;display:flex;align-items:center;justify-content:space-between;gap:18px;flex-wrap:wrap}.title-block h1{margin:0;font-size:30px;line-height:1.1;font-weight:800}.title-block span{display:block;margin-top:6px;color:var(--soft);font-size:15px;font-weight:600}.meta{display:flex;gap:10px;flex-wrap:wrap;justify-content:flex-end;color:var(--soft);font-size:14px;font-weight:700}.pill{padding:9px 12px;border:1px solid var(--line);border-radius:999px;background:rgba(17,24,39,.8)}main{max-width:1520px;margin:0 auto;padding:24px 28px 34px}.cards{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:16px;margin-bottom:18px}.card,.panel{border:1px solid var(--line);border-radius:8px;background:linear-gradient(180deg,rgba(17,24,39,.96),rgba(12,17,29,.96));box-shadow:var(--shadow)}.card{min-height:142px;padding:20px;position:relative;overflow:hidden}.card:after{content:"";position:absolute;inset:auto 0 0;height:3px;background:linear-gradient(90deg,var(--blue),var(--purple))}.label{color:var(--soft);font-size:15px;font-weight:800;text-transform:uppercase}.value{margin-top:10px;font-size:34px;line-height:1.08;font-weight:900;word-break:break-word}.hint{margin-top:10px;color:var(--muted);font-size:14px;font-weight:700}.good{color:var(--good)}.bad{color:var(--bad)}.accent{color:var(--blue)}.progress{width:100%;height:12px;margin-top:16px;overflow:hidden;border-radius:999px;background:#1b2638}.progress span{display:block;height:100%;width:0;max-width:100%;border-radius:inherit;background:linear-gradient(90deg,var(--blue),var(--purple))}.layout{display:grid;grid-template-columns:minmax(0,1.55fr) minmax(360px,.9fr);gap:18px;align-items:start}.panel-header{padding:18px 20px;border-bottom:1px solid var(--line);display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap}h2{margin:0;font-size:20px;line-height:1.2;font-weight:900}.controls{display:flex;gap:10px;flex-wrap:wrap}input,select,button{height:42px;border:1px solid var(--line);border-radius:6px;background:#080d16;color:var(--ink);font:inherit;font-size:15px;font-weight:700;padding:0 12px}button{cursor:pointer;border-color:transparent;background:linear-gradient(135deg,var(--blue-2),var(--purple))}.table-wrap{max-height:620px;overflow:hidden;position:relative}.table-wrap:after{content:"";position:absolute;left:0;right:0;bottom:0;height:58px;pointer-events:none;background:linear-gradient(180deg,transparent,rgba(12,17,29,.96))}table{width:100%;border-collapse:collapse;min-width:860px}th,td{padding:14px 16px;border-bottom:1px solid rgba(34,48,71,.8);text-align:right;white-space:nowrap;font-size:16px}th:first-child,td:first-child{text-align:left}th{color:var(--soft);background:rgba(8,13,22,.7);font-size:13px;text-transform:uppercase;cursor:pointer;user-select:none;position:sticky;top:0;z-index:2}tbody.auto-scroll{animation:sellerLoop 80s linear infinite;will-change:transform}tbody.auto-scroll:hover{animation-play-state:paused}.loop-copy{opacity:.96}@keyframes sellerLoop{from{transform:translateY(0)}to{transform:translateY(-50%)}}tbody tr:hover{background:rgba(56,189,248,.06)}.seller{color:#fff;font-weight:900}.money{font-weight:900}.tag{display:inline-flex;align-items:center;justify-content:center;min-width:92px;height:30px;border-radius:999px;font-size:13px;font-weight:900;border:1px solid}.tag.ok{color:#8ef8b7;background:rgba(34,197,94,.12);border-color:rgba(34,197,94,.38)}.tag.pending{color:#ff9aad;background:rgba(251,113,133,.12);border-color:rgba(251,113,133,.38)}.weekly{display:grid;gap:14px;padding:18px}.week-card{border:1px solid var(--line);border-radius:8px;background:rgba(8,13,22,.76);padding:16px}.week-top{display:flex;align-items:baseline;justify-content:space-between;gap:12px;margin-bottom:12px}.week-name{font-size:17px;font-weight:900}.week-percent{color:var(--blue);font-size:22px;font-weight:900}.week-values{display:grid;grid-template-columns:1fr 1fr;gap:10px;color:var(--soft);font-size:14px;font-weight:800}.week-values strong{display:block;margin-top:3px;color:var(--ink);font-size:18px}.week-edit{display:flex;gap:10px;margin-top:12px}.week-edit input{width:100%;min-width:0;height:38px}.save-goals{width:100%;margin-top:4px}.error{margin-bottom:18px;padding:16px 18px;border-radius:8px;border:1px solid rgba(251,113,133,.5);color:#fecdd3;background:rgba(127,29,29,.4);font-weight:800}@media(max-width:1180px){.cards{grid-template-columns:repeat(2,minmax(0,1fr))}.layout{grid-template-columns:1fr}}@media(max-width:640px){body{background:var(--bg)}.topbar-inner,main{padding-left:14px;padding-right:14px}.topbar-inner{align-items:flex-start;gap:14px}.title-block h1{font-size:24px}.title-block span{font-size:14px}.meta{width:100%;justify-content:flex-start}.pill{max-width:100%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}main{padding-top:16px;padding-bottom:24px}.cards{grid-template-columns:repeat(2,minmax(0,1fr));gap:10px}.card{min-height:118px;padding:14px}.label{font-size:12px}.value{font-size:22px;line-height:1.12}.hint{font-size:12px}.progress{height:9px;margin-top:10px}.layout{gap:14px}.panel-header{padding:14px;align-items:flex-start}h2{font-size:18px}.controls{width:100%;display:grid;grid-template-columns:1fr 1fr}.controls input{grid-column:1/-1}.controls input,.controls select,.controls button{width:100%}.table-wrap{max-height:none;overflow:visible}.table-wrap:after,thead{display:none}table,tbody,tr,td{display:block;width:100%;min-width:0}tbody.auto-scroll{animation:none}tr.loop-copy{display:none}tbody tr{margin:12px;padding:12px;border:1px solid var(--line);border-radius:8px;background:rgba(8,13,22,.76)}td{display:flex;align-items:center;justify-content:space-between;gap:12px;padding:8px 0;border-bottom:1px solid rgba(34,48,71,.5);text-align:right;white-space:normal;font-size:14px}td:last-child{border-bottom:0}td:before{content:attr(data-label);color:var(--soft);font-size:12px;font-weight:900;text-transform:uppercase;text-align:left}td.seller{display:block;padding-top:0;color:var(--ink);font-size:16px;text-align:left}td.seller:before{display:none}.tag{min-width:76px;height:28px}.weekly{padding:14px}.week-card{padding:14px}.week-values strong{font-size:16px}}@media(prefers-reduced-motion:reduce){tbody.auto-scroll{animation:none}}
  </style>
</head>
<body>
  <header class="topbar"><div class="topbar-inner"><div class="title-block"><h1>Acompanhamento Diario</h1><span>Relatorio comercial em tempo real</span></div><div class="meta"><span class="pill" id="workbook">Carregando...</span><span class="pill" id="updated"></span></div></div></header>
  <main><div id="error"></div><section class="cards"><article class="card"><div class="label">Compromisso</div><div class="value" id="commitment">R$ 0,00</div><div class="hint">Meta consolidada do periodo</div></article><article class="card"><div class="label">Atingido</div><div class="value good" id="reached">R$ 0,00</div><div class="progress"><span id="totalProgress"></span></div></article><article class="card"><div class="label">Falta</div><div class="value bad" id="missing">R$ 0,00</div><div class="hint" id="pendingCount">0 vendedores pendentes</div></article><article class="card"><div class="label">Realizado</div><div class="value accent" id="percent">0%</div><div class="hint" id="positiveCount">0 acima da meta</div></article></section><section class="layout"><article class="panel"><div class="panel-header"><h2>Performance por vendedor</h2><div class="controls"><input id="search" type="search" placeholder="Buscar vendedor"><select id="status"><option value="all">Todos</option><option value="pending">Com falta</option><option value="ok">Superou</option></select><button id="refresh" type="button">Atualizar</button></div></div><div class="table-wrap"><table><thead><tr><th data-sort="seller">Vendedor</th><th data-sort="commitment">Compromisso</th><th data-sort="reached">Atingido</th><th data-sort="missing">Falta</th><th data-sort="percent">%</th><th>Status</th></tr></thead><tbody id="tableBody"></tbody></table></div></article><aside class="panel"><div class="panel-header"><h2>Acompanhamento semanal</h2><span class="pill" id="weekCount">4 semanas</span></div><div class="weekly" id="weekly"></div></aside></section></main>
  <script>
    const state={rows:[],weeks:[],sortKey:"missing",sortDir:"desc"};const brl=new Intl.NumberFormat("pt-BR",{style:"currency",currency:"BRL"});const byId=id=>document.getElementById(id);const formatMoney=value=>brl.format(Number(value||0));const moneyClass=value=>Number(value)<=0?"good":"bad";const setText=(id,value)=>{byId(id).textContent=value};
    async function loadData(){byId("error").innerHTML="";try{const response=await fetch("/api/data?ts="+Date.now());if(!response.ok)throw new Error(await response.text());const data=await response.json();state.rows=data.rows||[];state.weeks=normalizeWeeks(data);renderSummary(data);renderTable();renderWeeks()}catch(error){byId("error").innerHTML=`<div class="error">${error.message}</div>`}}
    function normalizeWeeks(data){if(Array.isArray(data.weeks)&&data.weeks.length)return data.weeks;const summary=data.summary||{};const goal=Number(summary.weekGoal||0);return[1,2,3,4].map(week=>({name:`Semana ${week}`,goal,reached:week===1?Number(summary.weekRevenue||0):0,missing:week===1?Number(summary.weekMissing||goal):goal,percent:week===1?Number(summary.weekPercent||0):0}))}
    function renderSummary(data){const summary=data.summary||{};setText("workbook",data.workbook||"data.json");setText("updated","Atualizado em "+(data.lastModified||"-"));setText("commitment",formatMoney(summary.commitment));setText("reached",formatMoney(summary.reached));setText("missing",formatMoney(summary.missing));setText("percent",`${summary.percent||0}%`);setText("pendingCount",`${summary.pendingCount||0} vendedores pendentes`);setText("positiveCount",`${summary.positiveCount||0} acima da meta`);byId("totalProgress").style.width=`${Math.min(Number(summary.percent||0),100)}%`}
    function filteredRows(){const term=byId("search").value.trim().toLowerCase();const status=byId("status").value;return state.rows.filter(row=>!term||row.seller.toLowerCase().includes(term)).filter(row=>status==="all"||row.status===status).sort((a,b)=>{const av=a[state.sortKey];const bv=b[state.sortKey];const result=typeof av==="string"?av.localeCompare(bv):Number(av)-Number(bv);return state.sortDir==="asc"?result:-result})}
    function rowMarkup(row,copy=false){return`<tr${copy?' class="loop-copy"':''}><td class="seller" data-label="Vendedor">${row.seller}</td><td data-label="Compromisso">${formatMoney(row.commitment)}</td><td data-label="Atingido">${formatMoney(row.reached)}</td><td data-label="Falta" class="money ${moneyClass(row.missing)}">${formatMoney(row.missing)}</td><td data-label="%">${row.percent}%</td><td data-label="Status"><span class="tag ${row.status}">${row.status==="ok"?"Superou":"Falta"}</span></td></tr>`}
    function renderTable(){const rows=filteredRows();const markup=rows.map(row=>rowMarkup(row)).join("");const duplicateMarkup=rows.map(row=>rowMarkup(row,true)).join("");byId("tableBody").innerHTML=rows.length>8?markup+duplicateMarkup:markup;byId("tableBody").classList.toggle("auto-scroll",rows.length>8)}
    function renderWeeks(){const weeks=state.weeks.slice(0,5);setText("weekCount",`${weeks.length} semanas`);byId("weekly").innerHTML=weeks.map(week=>{const percent=Math.min(Number(week.percent||0),100);const goal=Number(week.goal||0).toFixed(2);return`<div class="week-card"><div class="week-top"><div class="week-name">${week.name}</div><div class="week-percent">${Number(week.percent||0).toFixed(1)}%</div></div><div class="progress"><span style="width:${percent}%"></span></div><div class="week-values"><span>Realizado<strong>${formatMoney(week.reached)}</strong></span><span>Meta<strong>${formatMoney(week.goal)}</strong></span></div><div class="week-edit"><input class="goal-input" type="number" min="0" step="0.01" value="${goal}" data-week="${week.name}" aria-label="Meta ${week.name}"></div></div>`}).join("")+`<button class="save-goals" id="saveGoals" type="button">Salvar metas</button>`;byId("saveGoals").addEventListener("click",saveWeeklyGoals)}
    async function saveWeeklyGoals(){byId("error").innerHTML="";try{const goals={};document.querySelectorAll(".goal-input").forEach(input=>{goals[input.dataset.week]=Number(input.value||0)});const response=await fetch("/api/weekly-goals",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({goals})});if(!response.ok)throw new Error(await response.text());await loadData()}catch(error){byId("error").innerHTML=`<div class="error">${error.message}</div>`}}
    document.querySelectorAll("th[data-sort]").forEach(header=>{header.addEventListener("click",()=>{const key=header.dataset.sort;if(state.sortKey===key){state.sortDir=state.sortDir==="asc"?"desc":"asc"}else{state.sortKey=key;state.sortDir=key==="seller"?"asc":"desc"}renderTable()})});byId("search").addEventListener("input",renderTable);byId("status").addEventListener("change",renderTable);byId("refresh").addEventListener("click",loadData);loadData();setInterval(loadData,30000);
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


def as_number(value):
    if value is None or value == "":
        return 0.0
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def money(value):
    return round(as_number(value), 2)


def find_compromisso_workbook():
    if EXCEL_PATH.exists():
        return EXCEL_PATH
    candidates = [
        path
        for path in BASE_DIR.glob("COMPROMISSO MAIO*.xlsm")
        if not path.name.startswith("~$")
    ]
    return max(candidates, key=lambda path: path.stat().st_mtime) if candidates else None


def default_weekly_goals():
    return {f"Semana {index}": DEFAULT_WEEKLY_GOAL for index in range(1, 6)}


def normalize_text(value):
    if value is None:
        return ""
    return str(value).strip()


def normalize_week_name(value):
    text = normalize_text(value).title()
    digits = "".join(ch for ch in text if ch.isdigit())
    if digits:
        return f"Semana {int(digits)}"
    return text or "Semana 1"


def week_sort_key(item):
    digits = "".join(ch for ch in item["name"] if ch.isdigit())
    return int(digits) if digits else 99


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


def read_sales_rows(workbook):
    sheet = workbook["Planilha2"]
    rows = []
    for row in sheet.iter_rows(min_row=4, values_only=True):
        seller = normalize_text(row[1] if len(row) > 1 else None)
        if not seller:
            continue
        if seller.upper() == "TOTAL":
            break
        if seller.upper() == "VENDEDOR":
            continue
        commitment = money(row[2] if len(row) > 2 else 0)
        reached = money(row[3] if len(row) > 3 else 0)
        missing = money(row[4] if len(row) > 4 else commitment - reached)
        average = money(row[6] if len(row) > 6 else 0)
        difference = money(row[7] if len(row) > 7 else average - reached)
        percent = (reached / commitment * 100) if commitment else 0
        rows.append({"seller": seller, "commitment": commitment, "reached": reached, "missing": missing, "average": average, "difference": difference, "percent": round(percent, 1), "status": "ok" if missing <= 0 else "pending"})
    return rows


def read_weekly_blocks(workbook):
    sheet = workbook["Planilha1"]
    weeks = []
    daily_totals = []
    day_groups = [(3, 4, 5), (6, 7, 8), (9, 10, 11), (12, 13, 14), (15, 16, 17)]
    for row_number in range(1, sheet.max_row + 1):
        week_name = normalize_text(sheet.cell(row_number, 2).value)
        if not week_name.upper().startswith("SEMANA"):
            continue
        total_row = None
        for candidate in range(row_number + 2, min(sheet.max_row, row_number + 30) + 1):
            if normalize_text(sheet.cell(candidate, 2).value).upper() == "TOTAL":
                total_row = candidate
                break
        if not total_row:
            continue
        week_goal = 0.0
        week_reached = 0.0
        for commitment_col, reached_col, missing_col in day_groups:
            date_value = sheet.cell(row_number, commitment_col + 1).value
            day_name = normalize_text(sheet.cell(row_number, commitment_col).value)
            commitment = money(sheet.cell(total_row, commitment_col).value)
            reached = money(sheet.cell(total_row, reached_col).value)
            missing = money(sheet.cell(total_row, missing_col).value)
            if not (commitment or reached or missing):
                continue
            week_goal = money(week_goal + commitment)
            week_reached = money(week_reached + reached)
            daily_totals.append({"week": week_name, "day": day_name, "date": date_value.date() if isinstance(date_value, datetime) else None, "commitment": commitment, "reached": reached, "missing": missing})
        week_missing = money(week_goal - week_reached)
        week_percent = (week_reached / week_goal * 100) if week_goal else 0
        weeks.append({"name": week_name.title(), "goal": week_goal, "reached": week_reached, "missing": week_missing, "percent": round(week_percent, 1)})
    return sorted(weeks, key=week_sort_key)[:5], daily_totals


def select_daily_total(daily_totals):
    if not daily_totals:
        return None
    today = datetime.now().date()
    for item in daily_totals:
        if item["date"] == today:
            return item
    dated_items = [item for item in daily_totals if item["date"] and item["date"] <= today]
    if dated_items:
        return max(dated_items, key=lambda item: item["date"])
    return daily_totals[-1]


def read_compromisso_maio_data():
    excel_path = find_compromisso_workbook()
    if not excel_path:
        raise FileNotFoundError(f"Arquivo nao encontrado: {EXCEL_PATH.name}")
    workbook = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)
    try:
        rows = read_sales_rows(workbook)
        weeks, daily_totals = read_weekly_blocks(workbook)
    finally:
        workbook.close()
    daily_total = select_daily_total(daily_totals)
    total_commitment = money(daily_total["commitment"] if daily_total else sum(item["commitment"] for item in rows))
    total_reached = money(daily_total["reached"] if daily_total else sum(item["reached"] for item in rows))
    total_missing = money(daily_total["missing"] if daily_total else total_commitment - total_reached)
    total_percent = (total_reached / total_commitment * 100) if total_commitment else 0
    return {"workbook": excel_path.name, "lastModified": datetime.fromtimestamp(excel_path.stat().st_mtime).strftime("%d/%m/%Y %H:%M"), "summary": {"commitment": total_commitment, "reached": total_reached, "missing": total_missing, "percent": round(total_percent, 1), "weekGoal": weeks[0]["goal"] if weeks else 0, "weekRevenue": weeks[0]["reached"] if weeks else 0, "weekMissing": weeks[0]["missing"] if weeks else 0, "weekPercent": weeks[0]["percent"] if weeks else 0, "positiveCount": sum(1 for item in rows if item["missing"] <= 0), "pendingCount": sum(1 for item in rows if item["missing"] > 0)}, "rows": sorted(rows, key=lambda item: item["missing"], reverse=True), "history": [], "weeks": weeks}


def read_legacy_excel_dashboard_data():
    workbook = openpyxl.load_workbook(LEGACY_EXCEL_PATH, data_only=True, read_only=True)
    try:
        sheet = workbook["Vendas"]
        rows_by_seller = {}
        weeks_by_name = {}
        for row in sheet.iter_rows(min_row=2, values_only=True):
            date_value, week, seller, goal, reached, *_ = list(row) + [None] * 2
            if not seller:
                continue
            seller_name = str(seller).strip()
            week_name = str(week or "Semana 1").strip()
            goal_value = money(goal)
            reached_value = money(reached)
            seller_row = rows_by_seller.setdefault(seller_name, {"seller": seller_name, "commitment": 0.0, "reached": 0.0})
            seller_row["commitment"] = money(seller_row["commitment"] + goal_value)
            seller_row["reached"] = money(seller_row["reached"] + reached_value)
            week_row = weeks_by_name.setdefault(week_name, {"name": week_name, "goal": 0.0, "reached": 0.0})
            week_row["goal"] = money(week_row["goal"] + goal_value)
            week_row["reached"] = money(week_row["reached"] + reached_value)
    finally:
        workbook.close()
    rows = []
    for item in rows_by_seller.values():
        missing = money(item["commitment"] - item["reached"])
        percent = (item["reached"] / item["commitment"] * 100) if item["commitment"] else 0
        rows.append({**item, "missing": missing, "average": 0.0, "difference": 0.0, "percent": round(percent, 1), "status": "ok" if missing <= 0 else "pending"})
    weeks = []
    for item in sorted(weeks_by_name.values(), key=week_sort_key)[:5]:
        missing = money(item["goal"] - item["reached"])
        percent = (item["reached"] / item["goal"] * 100) if item["goal"] else 0
        weeks.append({**item, "missing": missing, "percent": round(percent, 1)})
    total_commitment = money(sum(item["commitment"] for item in rows))
    total_reached = money(sum(item["reached"] for item in rows))
    total_missing = money(total_commitment - total_reached)
    total_percent = (total_reached / total_commitment * 100) if total_commitment else 0
    return {"workbook": LEGACY_EXCEL_PATH.name, "lastModified": datetime.fromtimestamp(LEGACY_EXCEL_PATH.stat().st_mtime).strftime("%d/%m/%Y %H:%M"), "summary": {"commitment": total_commitment, "reached": total_reached, "missing": total_missing, "percent": round(total_percent, 1), "weekGoal": weeks[0]["goal"] if weeks else 0, "weekRevenue": weeks[0]["reached"] if weeks else 0, "weekMissing": weeks[0]["missing"] if weeks else 0, "weekPercent": weeks[0]["percent"] if weeks else 0, "positiveCount": sum(1 for item in rows if item["missing"] <= 0), "pendingCount": sum(1 for item in rows if item["missing"] > 0)}, "rows": sorted(rows, key=lambda item: item["missing"], reverse=True), "history": [], "weeks": weeks}


def read_dashboard_data():
    if find_compromisso_workbook():
        return apply_weekly_goals(read_compromisso_maio_data())
    if LEGACY_EXCEL_PATH.exists():
        return apply_weekly_goals(read_legacy_excel_dashboard_data())
    if not DATA_PATH.exists():
        raise FileNotFoundError(f"Arquivo nao encontrado: {DATA_PATH}")
    return apply_weekly_goals(json.loads(DATA_PATH.read_text(encoding="utf-8")))


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
                payload = json.loads(body)
                goals = save_weekly_goals(payload.get("goals", {}))
                response = json.dumps({"goals": goals}, ensure_ascii=False).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Cache-Control", "no-store")
                self.send_header("Content-Length", str(len(response)))
                self.end_headers()
                self.wfile.write(response)
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
