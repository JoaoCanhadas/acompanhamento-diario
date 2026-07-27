@echo off
cd /d "%~dp0"
python sincronizar_sql_online.py --once >> sincronizacao_sql_online.log 2>&1
