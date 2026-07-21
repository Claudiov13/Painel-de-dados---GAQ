@echo off
chcp 65001 >nul
title Painel GAQ - Atualizando base_fracionamento.js
echo.
echo  ====================================================
echo   Painel GAQ - Atualizacao de base_fracionamento.js
echo  ====================================================
echo.
echo  Lendo Excel e gerando base_fracionamento.js...
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0gerar_base_fracionamento.ps1"

if %ERRORLEVEL% EQU 0 (
    echo.
    echo  [OK] base_fracionamento.js atualizado com sucesso!
    echo  O painel SharePoint ja pode ser recarregado.
) else (
    echo.
    echo  [ERRO] Algo deu errado. Verifique se a planilha:
    echo  Base de fracionamento MXM.xlsx
    echo  esta na pasta KB_GAQ no OneDrive.
)

echo.
echo  Esta janela fechara em 5 segundos...
powershell -NoProfile -Command "Start-Sleep -Seconds 5"
