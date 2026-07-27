from __future__ import annotations

import json
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE_DIR))

from dashboard import (
    BASE_DIR,
    DATA_PATH,
    GERAL_DATA_PATH,
    KEYS_DATA_PATH,
    POSITIVACAO_MILHO_DATA_PATH,
    read_dashboard_data,
    read_general_data,
    read_keys_data,
    read_positivacao_milho_data,
)

EXPORTS = [
    ("Faturamento", DATA_PATH, read_dashboard_data),
    ("Geral", GERAL_DATA_PATH, read_general_data),
    ("Postivação", POSITIVACAO_MILHO_DATA_PATH, read_positivacao_milho_data),
    ("Keys", KEYS_DATA_PATH, read_keys_data),
]


def main():
    print("Lendo fonte de dados e gerando arquivos...\n")
    erros = []

    for label, path, reader in EXPORTS:
        try:
            data = reader()
            path.write_text(
                json.dumps(data, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
            )
            workbook = data.get("workbook", "")
            modified = data.get("lastModified", "")
            rows = len(data.get("rows", []))
            print(f"  OK  {path.name:<30} {rows} linhas  ({workbook} - {modified})")
        except Exception as exc:
            print(f"  ERRO {label}: {exc}")
            erros.append(label)

    print()
    if erros:
        print(f"Falha em: {', '.join(erros)}")
        print("Verifique se o Excel esta fechado e sem erros de formula.")
        sys.exit(1)

    print("Todos os arquivos gerados com sucesso.")


if __name__ == "__main__":
    main()
