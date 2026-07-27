@echo off
cd /d "%~dp0"
title Configurar SQL Dashboard
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0configurar_sql_dashboard.ps1"
