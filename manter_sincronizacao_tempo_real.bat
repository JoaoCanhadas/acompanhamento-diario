@echo off
cd /d "%~dp0"
if exist ".env" (
  for /f "usebackq tokens=1,* delims==" %%A in (".env") do set "%%A=%%B"
)
python sincronizar_sql_online.py --interval 60
pause
