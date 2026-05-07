@echo off
title Sistema de Compras - GAQ
cd /d "%~dp0"

:loop
echo.
echo  Iniciando Sistema de Compras - Painel GAQ...
echo  [%date% %time%]
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0iniciar_servidor.ps1"
echo.
echo  [%date% %time%] Servidor encerrado (codigo: %errorlevel%). Reiniciando em 5 segundos...
echo  (Feche esta janela para parar definitivamente)
echo.
powershell -NoProfile -Command "Start-Sleep -Seconds 5"
goto loop
