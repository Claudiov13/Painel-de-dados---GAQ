@echo off
chcp 65001 >nul
title Painel GAQ - Configurar Agendamento

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0configurar_agendamento.ps1"

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo  [ERRO] Falha na configuracao. Tente executar como Administrador:
    echo  Clique direito neste .bat ^> Executar como administrador
    echo.
    pause
)
