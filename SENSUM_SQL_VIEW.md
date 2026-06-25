# Integracao com a view do Sensum

View criada pelo suporte:

```sql
dbo.VIW_IATAGEM_PEDIDO
```

Essa view traz pedidos linha a linha. O dashboard agora consegue usar essa view para calcular o realizado/atingido dos painéis.

## Campos usados

| Campo da view | Uso no dashboard |
| --- | --- |
| `AREA` | Filtro da empresa/local. Padrao: `IATAGAM`. |
| `DATA`, `DIA`, `MES`, `ANO` | Filtro do periodo e semanas. |
| `REGIAO` | Agrupamento principal dos painéis. |
| `REP` | Filtro auxiliar para Keys. |
| `COD_CLIENTE` | Contagem de postivação por cliente distinto. |
| `PRODUTO`, `GRUPO` | Filtro de produtos de milho. |
| `TOTAL` | Valor realizado dos painéis de faturamento. |

## Configuracao

No computador/servidor que for gerar os dados:

```powershell
$env:SENSUM_SQL_CONNECTION_STRING="DRIVER={ODBC Driver 17 for SQL Server};SERVER=servidor;DATABASE=banco;UID=usuario;PWD=senha;TrustServerCertificate=yes"
$env:SENSUM_SQL_VIEW="dbo.VIW_IATAGEM_PEDIDO"
$env:SENSUM_SQL_SOURCE_MODE="pedido"
$env:SENSUM_SQL_AREA="IATAGAM"
python exportar_dados.py
```

Filtros opcionais:

```powershell
$env:SENSUM_SQL_YEAR="2026"
$env:SENSUM_SQL_MONTH="5"
$env:SENSUM_SQL_UNTIL_DAY="29"
```

## Regras atuais

- `sales`: soma `TOTAL`, agrupado por `REGIAO` por padrao.
- `general`: soma `TOTAL`, agrupado por `REGIAO`.
- `keys`: soma `TOTAL`, filtrando `REGIAO` ou `REP` iniciando com `KEY`.
- `milho`: conta `COUNT(DISTINCT COD_CLIENTE)`, filtrando `GRUPO` ou `PRODUTO` contendo `MILHO`.
- As metas continuam vindo dos JSONs atuais enquanto nao existir tabela/view de metas no Sensum.

## Ajustes possiveis por variavel

```powershell
$env:SENSUM_SQL_SELLER_COLUMN="REGIAO"
$env:SENSUM_SQL_SALES_FILTER="COD_TIPO_OPERACAO IN (1,2,3)"
$env:SENSUM_SQL_KEYS_FILTER="UPPER(REGIAO) LIKE 'KEY%' OR UPPER(REP) LIKE 'KEY%'"
$env:SENSUM_SQL_MILHO_FILTER="UPPER(GRUPO) LIKE '%MILHO%' OR UPPER(PRODUTO) LIKE '%MILHO%'"
$env:SENSUM_SQL_MILHO_METRIC="COUNT(DISTINCT COD_CLIENTE)"
```

## Ainda falta definir

Para eliminar o Excel 100%, precisamos de uma fonte SQL para as metas:

- meta por vendedor/regiao;
- meta geral;
- meta Keys;
- meta Postivação;
- metas semanais, se forem diferentes das metas fixas atuais.

Enquanto isso, o SQL atualiza o realizado e o dashboard usa as metas salvas nos JSONs.
