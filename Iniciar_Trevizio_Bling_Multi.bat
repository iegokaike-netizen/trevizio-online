@echo off
title Trevizio Bling Multi
cd /d "%~dp0"
echo Iniciando Trevizio Bling Multi...
echo Se aparecer erro nessa janela, me mande print.
echo.
powershell -NoProfile -ExecutionPolicy Bypass -NoExit -File ".\trevizio_bling_multi.ps1"
echo.
echo O PowerShell fechou.
pause
