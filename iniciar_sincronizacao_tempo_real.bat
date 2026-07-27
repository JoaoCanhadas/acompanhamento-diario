@echo off
cd /d "%~dp0"
title Sincronizar SQL Online
python sincronizar_sql_online.py --interval 60
pause
