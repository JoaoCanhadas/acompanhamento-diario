from __future__ import annotations

import sys
from datetime import datetime

import sensum_sql


def main():
    if not sensum_sql.enabled():
        print("SQL nao configurado. Rode configurar_sql_dashboard.bat primeiro.")
        return 1

    panels = [
        ("sales", "Faturamento"),
        ("general", "Geral"),
        ("milho", "Postivacao"),
        ("keys", "Keys"),
    ]

    for panel, label in panels:
        data = sensum_sql.read_panel(panel)
        rows = data.get("rows", [])
        summary = data.get("summary", {})
        print(
            f"OK {label:<12} {len(rows):>3} linhas | "
            f"atingido={summary.get('reached', 0)} | "
            f"atualizado={data.get('lastModified', '-')}"
        )

    print()
    print("TESTE DIRETO DA VIEW")

    today = datetime.now()
    start_date = datetime(today.year, today.month, 1)

    if today.month == 12:
        end_date = datetime(today.year + 1, 1, 1)
    else:
        end_date = datetime(today.year, today.month + 1, 1)

    sql = """
        SELECT
            MAX(DATEFROMPARTS(ANO, MES, DIA)) AS ultima_data,
            SUM(TOTAL) AS total_geral,
            COUNT(*) AS quantidade_linhas
        FROM dbo.VIW_IATAGEM_PEDIDO
        WHERE DATEFROMPARTS(ANO, MES, DIA) >= ?
          AND DATEFROMPARTS(ANO, MES, DIA) < ?
    """

    resultado = sensum_sql.sql_fetch(sql, start_date, end_date)

    if resultado:
        linha = resultado[0]
        print(f"Ultima data da view: {linha.get('ultima_data')}")
        print(f"Total direto da view: {linha.get('total_geral')}")
        print(f"Quantidade de linhas: {linha.get('quantidade_linhas')}")
    else:
        print("A consulta direta nao retornou dados.")

    return 0


if __name__ == "__main__":
    sys.exit(main())