from __future__ import annotations

import sys

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
    return 0


if __name__ == "__main__":
    sys.exit(main())
