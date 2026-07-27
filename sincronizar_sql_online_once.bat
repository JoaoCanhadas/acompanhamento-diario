@echo off
cd /d "%~dp0"
if exist ".env" (
  for /f "usebackq tokens=1,* delims==" %%A in (".env") do set "%%A=%%B"
)
python sincronizar_sql_online.py --once >> sincronizacao_sql_online.log 2>&1
