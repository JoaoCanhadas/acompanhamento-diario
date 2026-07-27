$ErrorActionPreference = "Stop"

function Read-Default {
    param(
        [string]$Prompt,
        [string]$Default
    )
    $Value = Read-Host "$Prompt [$Default]"
    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $Default
    }
    return $Value.Trim()
}

function Secure-To-Text {
    param([securestring]$Value)
    $Bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Value)
    try {
        return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($Bstr)
    } finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($Bstr)
    }
}

function Set-UserEnv {
    param(
        [string]$Name,
        [string]$Value
    )
    [Environment]::SetEnvironmentVariable($Name, $Value, "User")
    Set-Item -Path "Env:$Name" -Value $Value
}

function Pause-Exit {
    Write-Host ""
    Read-Host "Pressione Enter para fechar"
}

Write-Host ""
Write-Host "============================================"
Write-Host "  CONFIGURAR SQL DO DASHBOARD"
Write-Host "============================================"
Write-Host ""

$Server = Read-Host "Servidor SQL"
$Database = Read-Host "Banco de dados"
$User = Read-Default "Usuario SQL" "powerbi"
$Password = Secure-To-Text (Read-Host "Senha SQL" -AsSecureString)
$Port = Read-Default "Porta SQL" "1433"
$View = Read-Default "View de pedidos" "dbo.VIW_IATAGAM_PEDIDO"
$Area = Read-Default "Area/filtro" "IATAGAM"

if ([string]::IsNullOrWhiteSpace($Server) -or
    [string]::IsNullOrWhiteSpace($Database) -or
    [string]::IsNullOrWhiteSpace($User) -or
    [string]::IsNullOrWhiteSpace($Password)) {
    Write-Host ""
    Write-Host "ERRO: servidor, banco, usuario e senha sao obrigatorios."
    Pause-Exit
    exit 1
}

Set-UserEnv "SENSUM_SQL_SERVER" $Server.Trim()
Set-UserEnv "SENSUM_SQL_DATABASE" $Database.Trim()
Set-UserEnv "SENSUM_SQL_USER" $User.Trim()
Set-UserEnv "SENSUM_SQL_PASSWORD" $Password
Set-UserEnv "SENSUM_SQL_PORT" $Port.Trim()
Set-UserEnv "SENSUM_SQL_VIEW" $View.Trim()
Set-UserEnv "SENSUM_SQL_SOURCE_MODE" "pedido"
Set-UserEnv "SENSUM_SQL_AREA" $Area.Trim()

Write-Host ""
Write-Host "Instalando/conferindo dependencia SQL..."
python -m pip install pymssql
if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "ERRO: nao consegui instalar pymssql."
    Write-Host "Verifique se o Python/pip esta funcionando nessa maquina."
    Pause-Exit
    exit 1
}

Write-Host ""
Write-Host "Testando a conexao..."
python sensum_sql_teste.py
if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "As variaveis foram salvas, mas o teste SQL falhou."
    Write-Host "Confira servidor, banco, usuario, senha, porta e liberacao de rede."
    Pause-Exit
    exit 1
}

Write-Host ""
Write-Host "============================================"
Write-Host "  SQL CONFIGURADO COM SUCESSO"
Write-Host "============================================"
Write-Host ""
Write-Host "Agora reinicie o dashboard/publicacao para usar a view em tempo real."
Pause-Exit
