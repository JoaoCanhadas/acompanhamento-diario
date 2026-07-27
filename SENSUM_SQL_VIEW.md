# Integracao com a view do Sensum

View criada pelo suporte:

```sql
dbo.VIW_IATAGEM_PEDIDO
```

Essa view traz pedidos linha a linha. O dashboard usa essa view para calcular o realizado/atingido dos paineis.

Com o SQL configurado, as APIs do dashboard consultam o banco primeiro:

- `/api/data`
- `/api/geral`
- `/api/keys`
- `/api/positivacao-milho`

Os arquivos JSON/Excel ficam apenas como fallback operacional.

## Campos usados

| Campo da view | Uso no dashboard |
| --- | --- |
| `AREA` | Filtro da empresa/local. Padrao: `IATAGAM`. |
| `DATA`, `DIA`, `MES`, `ANO` | Filtro do periodo e semanas. |
| `REGIAO` | Agrupamento principal dos paineis. |
| `REP` | Filtro auxiliar para Keys. |
| `COD_CLIENTE` | Contagem de postivacao por cliente distinto. |
| `PRODUTO`, `GRUPO` | Filtro de produtos de milho/brioche conforme variavel. |
| `TOTAL` | Valor realizado dos paineis de faturamento. |

## Configuracao local

No computador que for testar ou publicar, rode:

```bat
configurar_sql_dashboard.bat
```

Ele pede servidor, banco, usuario, senha, porta, view e area. Depois salva estas variaveis no usuario do Windows:

```powershell
$env:SENSUM_SQL_SERVER="servidor"
$env:SENSUM_SQL_DATABASE="banco"
$env:SENSUM_SQL_USER="powerbi"
$env:SENSUM_SQL_PASSWORD="senha"
$env:SENSUM_SQL_PORT="1433"
$env:SENSUM_SQL_VIEW="dbo.VIW_IATAGEM_PEDIDO"
$env:SENSUM_SQL_SOURCE_MODE="pedido"
$env:SENSUM_SQL_AREA="IATAGAM"
```

Para testar depois:

```bat
testar_sql_dashboard.bat
```

Tambem funciona por connection string ODBC, se a maquina tiver driver ODBC:

```powershell
$env:SENSUM_SQL_CONNECTION_STRING="DRIVER={ODBC Driver 17 for SQL Server};SERVER=servidor;DATABASE=banco;UID=usuario;PWD=senha;TrustServerCertificate=yes"
$env:SENSUM_SQL_VIEW="dbo.VIW_IATAGEM_PEDIDO"
$env:SENSUM_SQL_SOURCE_MODE="pedido"
$env:SENSUM_SQL_AREA="IATAGAM"
python exportar_dados.py
```

## Configuracao no Render

Para o dashboard online ficar em tempo real, configure no Render as mesmas variaveis:

```text
SENSUM_SQL_SERVER
SENSUM_SQL_DATABASE
SENSUM_SQL_USER
SENSUM_SQL_PASSWORD
SENSUM_SQL_PORT
SENSUM_SQL_VIEW
SENSUM_SQL_SOURCE_MODE=pedido
SENSUM_SQL_AREA=IATAGAM
```

Depois publique os arquivos atualizados com `publicar.bat`.

## Filtros opcionais

```powershell
$env:SENSUM_SQL_YEAR="2026"
$env:SENSUM_SQL_MONTH="7"
$env:SENSUM_SQL_UNTIL_DAY="27"
$env:SENSUM_SQL_SELLER_COLUMN="REGIAO"
$env:SENSUM_SQL_SALES_FILTER="COD_TIPO_OPERACAO IN (1,2,3)"
$env:SENSUM_SQL_KEYS_FILTER="UPPER(REGIAO) LIKE 'KEY%' OR UPPER(REP) LIKE 'KEY%'"
$env:SENSUM_SQL_MILHO_FILTER="UPPER(GRUPO) LIKE '%MILHO%' OR UPPER(PRODUTO) LIKE '%MILHO%'"
$env:SENSUM_SQL_MILHO_METRIC="COUNT(DISTINCT COD_CLIENTE)"
```

## Regras atuais

- `sales`: soma `TOTAL`, agrupado por `REGIAO` por padrao.
- `general`: soma `TOTAL`, agrupado por `REGIAO`.
- `keys`: soma `TOTAL`, filtrando `REGIAO` ou `REP` iniciando com `KEY`.
- `milho`: conta `COUNT(DISTINCT COD_CLIENTE)`, filtrando `GRUPO` ou `PRODUTO` contendo `MILHO` por padrao.
- O realizado/atingido vem do SQL em tempo real quando as variaveis estiverem configuradas.
- As metas continuam vindo dos JSONs atuais enquanto nao existir tabela/view de metas no Sensum.

## Ainda falta definir para eliminar o Excel 100%

Para eliminar o Excel tambem das metas, precisamos de uma fonte SQL para:

- meta por vendedor/regiao;
- meta geral;
- meta Keys;
- meta Postivacao;
- metas semanais, se forem diferentes das metas fixas atuais.
