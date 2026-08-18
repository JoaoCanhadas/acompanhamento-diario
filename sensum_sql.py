from __future__ import annotations

import json
import os
from datetime import datetime, timedelta
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent
AJUSTES_REGIAO_PATH = BASE_DIR / "ajustes_regiao.json"
PANEL_JSON = {
    "sales": BASE_DIR / "data.json",
    "general": BASE_DIR / "geral.json",
    "keys": BASE_DIR / "keys.json",
    "milho": BASE_DIR / "positivacao_milho.json",
}

PANEL_TEMPLATE_JSON = {
    "sales": BASE_DIR / "template_data.json",
    "general": BASE_DIR / "template_geral.json",
    "keys": BASE_DIR / "template_keys.json",
    "milho": BASE_DIR / "template_positivacao_milho.json",
}


def load_env_file():
    env_path = BASE_DIR / ".env"
    if not env_path.exists():
        return
    for line in env_path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        name, value = line.split("=", 1)
        os.environ.setdefault(name.strip(), value.strip())


load_env_file()


PANEL_CONFIG = {
    "sales": {
        "title": "Faturamento",
        "default_week_goal": 700000.0,
        "weekly_goal_override": None,
    },
    "general": {
        "title": "Geral",
        "default_week_goal": 0.0,
        "weekly_goal_override": None,
    },
    "keys": {
        "title": "Keys",
        "default_week_goal": 0.0,
        "weekly_goal_override": None,
    },
    "milho": {
        "title": "Postivação",
        "default_week_goal": 0.0,
        "weekly_goal_override": None,
    },
}


def enabled():
    return bool(
        os.environ.get("SENSUM_SQL_CONNECTION_STRING")
        or (
            os.environ.get("SENSUM_SQL_SERVER")
            and os.environ.get("SENSUM_SQL_DATABASE")
            and os.environ.get("SENSUM_SQL_USER")
            and os.environ.get("SENSUM_SQL_PASSWORD")
        )
    )


def status():
    if not enabled():
        return {
            "mode": "json-fallback",
            "sqlEnabled": False,
            "sqlOk": False,
            "message": "Variaveis SQL nao configuradas.",
        }

    try:
        view_name = os.environ.get("SENSUM_SQL_VIEW", "dbo.VIW_IATAGEM_PEDIDO")
        rows = sql_fetch(f"SELECT TOP 1 1 AS ok FROM {view_name}")
        return {
            "mode": "sql-live",
            "sqlEnabled": True,
            "sqlOk": bool(rows),
            "view": view_name,
            "message": "Conexao SQL ativa.",
        }
    except Exception as exc:
        return {
            "mode": "json-fallback",
            "sqlEnabled": True,
            "sqlOk": False,
            "message": str(exc),
        }


def read_panel(panel):
    if not enabled():
        return None
    if panel not in PANEL_CONFIG:
        raise ValueError(f"Painel SQL desconhecido: {panel}")

    source_mode = os.environ.get("SENSUM_SQL_SOURCE_MODE", "pedido").lower()
    if source_mode == "pedido":
        return read_pedido_panel(panel)

    rows = query_rows(panel)
    if not rows:
        raise ValueError(f"A view SQL nao retornou linhas para o painel {panel}")

    row_items = [normalize_row(item) for item in rows if row_kind(item) == "row"]
    week_items = [normalize_week(item, panel) for item in rows if row_kind(item) == "week"]
    summary_items = [item for item in rows if row_kind(item) == "summary"]
    if not row_items:
        raise ValueError(f"A view SQL nao retornou linhas tipo 'row' para o painel {panel}")

    row_items = sorted(row_items, key=lambda item: item["missing"], reverse=True)
    weeks = sorted(week_items, key=week_sort_key)[:5]
    summary = build_summary(row_items, weeks, summary_items, panel)
    last_modified = max(
        (coalesce(item, "updated_at", "last_modified", "data_atualizacao") for item in rows),
        default=None,
    )

    return {
        "workbook": f"Sensum SQL - {PANEL_CONFIG[panel]['title']}",
        "lastModified": format_timestamp(last_modified),
        "summary": summary,
        "rows": row_items,
        "history": [],
        "weeks": weeks,
    }


def query_rows(panel):
    view_name = os.environ.get("SENSUM_SQL_VIEW", "dbo.vw_acompanhamento_diario")
    sql = f"SELECT * FROM {view_name} WHERE painel = ?"
    return sql_fetch(sql, panel)


def read_pedido_panel(panel):
    template = load_template(panel)
    reached_rows = query_pedido_reached(panel)
    row_items = merge_template_rows(panel, template, reached_rows)
    weeks = merge_template_weeks(panel, template, query_pedido_weeks(panel))
    summary = build_pedido_summary(row_items, weeks, template, panel)

    return {
        "workbook": f"Sensum SQL - {PANEL_CONFIG[panel]['title']}",
        "lastModified": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "summary": summary,
        "rows": sorted(row_items, key=lambda item: item["missing"], reverse=True),
        "history": [],
        "weeks": weeks,
    }


def query_pedido_reached(panel):
    select_name, where_extra, aggregate = pedido_panel_sql(panel)
    start_date, end_date = reached_period_range(panel)
    view_name = os.environ.get("SENSUM_SQL_VIEW", "dbo.VIW_IATAGEM_PEDIDO")
    area_filter = os.environ.get("SENSUM_SQL_AREA", "IATAGAM")
    date_expr = pedido_date_expr()
    sql = f"""
        SELECT
            {select_name} AS seller,
            {aggregate} AS reached
        FROM {view_name}
        WHERE {date_expr} >= ? AND {date_expr} < ?
          AND (? = '' OR AREA = ?)
          {where_extra}
        GROUP BY {select_name}
    """
    return sql_fetch(sql, start_date, end_date, area_filter, area_filter)


def query_pedido_weeks(panel):
    _, where_extra, aggregate = pedido_panel_sql(panel)
    start_date, end_date = period_range()
    view_name = os.environ.get("SENSUM_SQL_VIEW", "dbo.VIW_IATAGEM_PEDIDO")
    area_filter = os.environ.get("SENSUM_SQL_AREA", "IATAGAM")
    date_expr = pedido_date_expr()
    sql = f"""
        SELECT
            DATEPART(ISO_WEEK, {date_expr}) AS week_number,
            {aggregate} AS reached
        FROM {view_name}
        WHERE {date_expr} >= ? AND {date_expr} < ?
          AND (? = '' OR AREA = ?)
          {where_extra}
        GROUP BY DATEPART(ISO_WEEK, {date_expr})
        ORDER BY DATEPART(ISO_WEEK, {date_expr})
    """
    return sql_fetch(sql, start_date, end_date, area_filter, area_filter)


def sql_fetch(sql, *params):
    connection, driver = open_connection()
    try:
        if driver == "pymssql":
            sql = sql.replace("?", "%s")
        cursor = connection.cursor()
        cursor.execute(sql, params)
        columns = [column[0].lower() for column in cursor.description]
        return [dict(zip(columns, row)) for row in cursor.fetchall()]
    finally:
        connection.close()


def open_connection():
    connection_string = os.environ.get("SENSUM_SQL_CONNECTION_STRING")
    if connection_string:
        try:
            import pyodbc
        except ImportError as exc:
            raise RuntimeError(
                "Instale o pyodbc para ler o SQL por connection string: python -m pip install pyodbc"
            ) from exc
        return pyodbc.connect(connection_string, timeout=30), "pyodbc"

    try:
        import pymssql
    except ImportError as exc:
        raise RuntimeError(
            "Instale o pymssql para ler o SQL: python -m pip install pymssql"
        ) from exc

    return (
        pymssql.connect(
            server=os.environ["SENSUM_SQL_SERVER"],
            database=os.environ["SENSUM_SQL_DATABASE"],
            user=os.environ["SENSUM_SQL_USER"],
            password=os.environ["SENSUM_SQL_PASSWORD"],
            port=int(os.environ.get("SENSUM_SQL_PORT", "1433")),
            login_timeout=30,
            timeout=30,
        ),
        "pymssql",
    )


def general_region_sql():
    if not AJUSTES_REGIAO_PATH.exists():
        return "REGIAO"

    try:
        ajustes = json.loads(
            AJUSTES_REGIAO_PATH.read_text(encoding="utf-8")
        )
    except Exception:
        return "REGIAO"

    mes_atual = datetime.now().strftime("%Y-%m")
    ajustes_mes = ajustes.get(mes_atual, {})

    if not ajustes_mes:
        return "REGIAO"

    grupos = {}

    for pedido, regiao in ajustes_mes.items():
        pedido = str(pedido).strip()
        regiao = str(regiao).strip()

        if not pedido or not regiao:
            continue

        grupos.setdefault(regiao, []).append(pedido)

    if not grupos:
        return "REGIAO"

    cases = []

    for regiao, pedidos in grupos.items():
        pedidos_sql = ", ".join(
            f"'{pedido}'"
            for pedido in pedidos
        )

        regiao_sql = regiao.replace("'", "''")

        cases.append(
            f"WHEN CAST(PED AS VARCHAR(50)) IN ({pedidos_sql}) "
            f"THEN '{regiao_sql}'"
        )

    return (
        "CASE "
        + " ".join(cases)
        + " ELSE REGIAO END"
    )


def pedido_panel_sql(panel):
    seller_column = os.environ.get("SENSUM_SQL_SELLER_COLUMN", "REGIAO")

    if panel == "general":
        return general_region_sql(), "", "SUM(TOTAL)"

    if panel == "sales":
        return seller_column, sales_filter_sql(), "SUM(TOTAL)"

    if panel == "keys":
        return "REGIAO", keys_filter_sql(), "SUM(TOTAL)"

    if panel == "milho":
        metric = os.environ.get(
            "SENSUM_SQL_MILHO_METRIC",
            "COUNT(DISTINCT COD_CLIENTE)"
        )
        return "REGIAO", milho_filter_sql(), metric

    raise ValueError(f"Painel SQL desconhecido: {panel}")


def pedido_date_expr():
    return os.environ.get("SENSUM_SQL_DATE_EXPR", "DATEFROMPARTS(ANO, MES, DIA)")


def sales_filter_sql():
    value = os.environ.get("SENSUM_SQL_SALES_FILTER", "UPPER(REGIAO) NOT LIKE 'KEY%'")
    return f" AND ({value})" if value else ""


def keys_filter_sql():
    value = os.environ.get(
        "SENSUM_SQL_KEYS_FILTER",
        "UPPER(REGIAO) LIKE 'KEY%' OR UPPER(REP) LIKE 'KEY%'",
    )
    return f" AND ({value})" if value else ""


def milho_filter_sql():
    value = os.environ.get(
        "SENSUM_SQL_MILHO_FILTER",
        "(UPPER(GRUPO) LIKE '%BRIOCHE%' OR UPPER(PRODUTO) LIKE '%BRIOCHE%') AND UPPER(REGIAO) NOT LIKE 'KEY%'",
    )
    return f" AND ({value})" if value else ""


def period_range():
    today = datetime.now()
    year = int(os.environ.get("SENSUM_SQL_YEAR", today.year))
    month = int(os.environ.get("SENSUM_SQL_MONTH", today.month))
    start = datetime(year, month, 1)
    if month == 12:
        end = datetime(year + 1, 1, 1)
    else:
        end = datetime(year, month + 1, 1)
    until_day = os.environ.get("SENSUM_SQL_UNTIL_DAY")
    if until_day:
        end = min(end, datetime(year, month, int(until_day)) + timedelta(days=1))
    return start, end


def reached_period_range(panel):
    if panel == "sales":
        return current_day_range()
    if panel == "keys":
        return current_week_range()
    return period_range()


def current_day_range():
    month_start, month_end = period_range()
    today = datetime.now()
    year = int(os.environ.get("SENSUM_SQL_YEAR", today.year))
    month = int(os.environ.get("SENSUM_SQL_MONTH", today.month))
    day = int(os.environ.get("SENSUM_SQL_UNTIL_DAY", today.day))
    current = datetime(year, month, day)
    return max(month_start, current), min(month_end, current + timedelta(days=1))


def current_week_range():
    month_start, month_end = period_range()
    today = datetime.now()
    year = int(os.environ.get("SENSUM_SQL_YEAR", today.year))
    month = int(os.environ.get("SENSUM_SQL_MONTH", today.month))
    day = int(os.environ.get("SENSUM_SQL_UNTIL_DAY", today.day))
    current = datetime(year, month, day)
    week_start = current - timedelta(days=current.weekday())
    return max(month_start, week_start), min(month_end, current + timedelta(days=1))


def load_template(panel):
    path = PANEL_TEMPLATE_JSON.get(panel, PANEL_JSON[panel])
    if not path.exists():
        path = PANEL_JSON[panel]
    if not path.exists():
        return {"summary": {}, "rows": [], "weeks": []}
    return json.loads(path.read_text(encoding="utf-8"))


def merge_template_rows(panel, template, reached_rows):
    reached_by_name = {
        normalize_name(item.get("seller")): money(item.get("reached"))
        for item in reached_rows
    }

    rows = []
    template_rows = template.get("rows") or []

    for item in template_rows:
        seller = text(item.get("seller"))
        reference = text(item.get("reference"))

        reached = reached_by_name.get(
            normalize_name(seller),
            reached_by_name.get(normalize_name(reference), 0.0),
        )

        commitment = money(item.get("commitment"))
        missing = money(commitment - reached)

        rows.append(
            {
                "reference": reference,
                "seller": seller,
                "commitment": commitment,
                "reached": reached,
                "missing": missing,
                "average": money(item.get("average")),
                "difference": money(item.get("difference")),
                "percent": percent_value(None, reached, commitment),
                "status": "ok" if missing <= 0 else "pending",
            }
        )

    known_names = {normalize_name(item.get("seller")) for item in rows}
    known_refs = {normalize_name(item.get("reference")) for item in rows}

    for item in reached_rows:
        seller = text(item.get("seller"))
        key = normalize_name(seller)

        if panel == "milho":
            continue

        if panel == "general" and key == "KEY PIRACICABA":
            continue

        if key in known_names or key in known_refs:
            continue

        reached = money(item.get("reached"))

        rows.append(
            {
                "reference": seller if panel != "sales" else "",
                "seller": seller,
                "commitment": 0.0,
                "reached": reached,
                "missing": money(-reached),
                "average": 0.0,
                "difference": 0.0,
                "percent": 0,
                "status": "ok",
            }
        )

    return rows


def merge_template_weeks(panel, template, reached_weeks):
    reached_values = [money(item.get("reached")) for item in reached_weeks]
    template_weeks = template.get("weeks") or []
    weeks = []

    for index, item in enumerate(template_weeks[:5]):
        goal = PANEL_CONFIG[panel]["weekly_goal_override"]

        if goal is None:
            goal = round(
                money(
                    item.get(
                        "goal",
                        PANEL_CONFIG[panel]["default_week_goal"]
                    )
                )
            )

        reached = reached_values[index] if index < len(reached_values) else 0.0
        missing = money(goal - reached)

        weeks.append(
            {
                "name": text(item.get("name") or f"Semana {index + 1}"),
                "goal": goal,
                "reached": reached,
                "missing": missing,
                "percent": percent_value(None, reached, goal),
            }
        )

    return weeks


def build_pedido_summary(rows, weeks, template, panel):
    template_summary = template.get("summary") or {}
    commitment = money(template_summary.get("commitment", sum(item["commitment"] for item in rows)))
    reached = money(sum(item["reached"] for item in rows))
    missing = money(commitment - reached)
    current_week = weeks[0] if weeks else {}
    return {
        "commitment": commitment,
        "reached": reached,
        "missing": missing,
        "percent": percent_value(None, reached, commitment),
        "weekGoal": current_week.get("goal", 0),
        "weekRevenue": current_week.get("reached", 0),
        "weekMissing": current_week.get("missing", 0),
        "weekPercent": current_week.get("percent", 0),

        "positiveCount": sum(
            1
            for item in rows
            if item["commitment"] > 0
            and item["missing"] <= 0
        ),

        "pendingCount": sum(
            1
            for item in rows
            if item["commitment"] > 0
            and item["missing"] > 0
        ),
    }


def normalize_name(value):
    return text(value).upper().strip()


def normalize_row(item):
    reference = text(coalesce(item, "reference", "referencia"))
    seller = text(coalesce(item, "seller", "vendedor", "nome"))
    commitment = money(coalesce(item, "commitment", "meta", "compromisso", "meta_mes"))
    reached = money(coalesce(item, "reached", "atingido", "realizado"))
    missing = money(coalesce(item, "missing", "falta", default=commitment - reached))
    percent = percent_value(coalesce(item, "percent", "percentual", default=None), reached, commitment)

    return {
        "reference": reference,
        "seller": seller or reference,
        "commitment": commitment,
        "reached": reached,
        "missing": missing,
        "average": money(coalesce(item, "average", "media", default=0)),
        "difference": money(coalesce(item, "difference", "diferenca", default=0)),
        "percent": percent,
        "status": "ok" if missing <= 0 else "pending",
    }


def normalize_week(item, panel):
    config = PANEL_CONFIG[panel]
    goal = config["weekly_goal_override"]
    if goal is None:
        goal = money(coalesce(item, "goal", "meta", "commitment", "compromisso", default=config["default_week_goal"]))
    reached = money(coalesce(item, "reached", "atingido", "realizado", default=0))
    missing = money(coalesce(item, "missing", "falta", default=goal - reached))
    return {
        "name": text(coalesce(item, "week_name", "semana", "name", default="Semana")),
        "goal": goal,
        "reached": reached,
        "missing": missing,
        "percent": percent_value(coalesce(item, "percent", "percentual", default=None), reached, goal),
    }


def build_summary(rows, weeks, summary_items, panel):
    summary_row = summary_items[0] if summary_items else {}
    commitment = money(coalesce(summary_row, "commitment", "meta", "compromisso", "meta_mes", default=sum(item["commitment"] for item in rows)))
    reached = money(coalesce(summary_row, "reached", "atingido", "realizado", default=sum(item["reached"] for item in rows)))
    missing = money(coalesce(summary_row, "missing", "falta", default=commitment - reached))
    current_week = weeks[0] if weeks else {}
    return {
        "commitment": commitment,
        "reached": reached,
        "missing": missing,
        "percent": percent_value(coalesce(summary_row, "percent", "percentual", default=None), reached, commitment),
        "weekGoal": current_week.get("goal", 0),
        "weekRevenue": current_week.get("reached", 0),
        "weekMissing": current_week.get("missing", 0),
        "weekPercent": current_week.get("percent", 0),
        "positiveCount": sum(
            1
            for item in rows
            if item["commitment"] > 0
            and item["missing"] <= 0
        ),
        "pendingCount": sum(
            1
            for item in rows
            if item["commitment"] > 0
            and item["missing"] > 0
        ),
    }


def row_kind(item):
    value = text(coalesce(item, "tipo", "kind", "linha_tipo", default="row")).lower()
    if value in {"semana", "week", "weekly"}:
        return "week"
    if value in {"resumo", "summary", "total"}:
        return "summary"
    return "row"


def coalesce(item, *names, default=""):
    for name in names:
        if name in item and item[name] not in (None, ""):
            return item[name]
    return default


def text(value):
    return "" if value is None else str(value).strip()


def money(value):
    try:
        return round(float(value or 0), 2)
    except (TypeError, ValueError):
        return 0.0


def percent_value(value, reached, goal):
    if value not in (None, ""):
        return round(float(value), 1)
    return round((reached / goal * 100) if goal else 0, 1)


def week_sort_key(item):
    name = item.get("name", "")
    digits = "".join(char for char in name if char.isdigit())
    return int(digits or 99), name


def format_timestamp(value):
    if isinstance(value, datetime):
        return value.strftime("%d/%m/%Y %H:%M")
    if value:
        return str(value)
    return datetime.now().strftime("%d/%m/%Y %H:%M")
