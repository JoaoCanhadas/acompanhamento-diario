from __future__ import annotations

import json
import sys

import openpyxl

import dashboard


EXPORTS = [
    ("Faturamento", "template_data.json", dashboard.read_base_python_planilha1_data),
    ("Geral", "template_geral.json", dashboard.read_general_planilha1_data),
    ("Postivacao", "template_positivacao_milho.json", dashboard.read_positivacao_milho_planilha1_data),
    ("Keys", "template_keys.json", dashboard.read_keys_planilha1_data),
]


def main():
    excel_path = dashboard.PYTHON_EXCEL_PATH
    if not excel_path.exists():
        print(f"Arquivo de metas nao encontrado: {excel_path}")
        return 1

    print("Lendo BASE PYTHON.xlsx e atualizando templates de metas...\n")
    workbook = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)
    try:
        for label, filename, reader in EXPORTS:
            data = reader(excel_path, workbook)
            path = dashboard.BASE_DIR / filename
            path.write_text(
                json.dumps(data, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
            )
            print(f"  OK  {filename:<35} {len(data.get('rows', []))} linhas  ({label})")
    finally:
        workbook.close()

    print("\nTemplates de metas atualizados.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
