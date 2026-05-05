# Hospedagem do dashboard

Este projeto esta pronto para hospedar como um servico web Python.

## Render

1. Crie uma conta em https://render.com.
2. No Render, clique em **New > Web Service**.
3. Conecte o repositorio do GitHub.
4. Use estas configuracoes:
   - **Runtime:** Python
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `python dashboard.py`
5. Escolha o plano **Free** e publique.

## Base Excel

O dashboard prioriza o arquivo `COMPROMISSO MAIO.xlsm`. A tabela de vendedores vem da
`Planilha2`, e o compromisso diario vem dos blocos semanais da `Planilha1`.

Se `COMPROMISSO MAIO.xlsm` nao existir no servidor, o sistema usa `base_vendas.xlsx`
e depois `data.json` como fallback.

Colunas da aba `Vendas` no fallback `base_vendas.xlsx`:

- `Data`
- `Semana`
- `Nome do vendedor`
- `Meta`
- `Realizado`
- `% de atingimento`
- `Observacoes`

Para criar o modelo inicial:

```bash
python gerar_base_excel.py
```

## Atualizacao automatica

A tela consulta `/api/data` a cada 30 segundos. Esse endpoint rele o Excel a cada chamada,
entao novas alteracoes salvas em `COMPROMISSO MAIO.xlsm` aparecem no dashboard sem recarregar
manualmente.

## Metas semanais

As metas semanais ficam em `weekly_goals.json`. O valor inicial esta em `700000`
para as semanas 1 a 5.

Tambem da para ajustar essas metas no proprio dashboard, nos campos do painel
`Acompanhamento semanal`, usando o botao `Salvar metas`.
