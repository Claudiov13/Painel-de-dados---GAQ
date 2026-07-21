@echo off
chcp 65001 >nul
title Painel GAQ - Atualizando dados.js
echo.
echo  ====================================================
echo   Painel GAQ - Atualizacao de dados.js
echo  ====================================================
echo.
echo  Lendo Excel e gerando dados.js...
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0gerar_dados_js.ps1"

if %ERRORLEVEL% EQU 0 (
    echo.
    echo  [OK] dados.js atualizado com sucesso!
    echo  O painel SharePoint ja pode ser recarregado.
) else (
    echo.
    echo  [ERRO] Algo deu errado. Verifique o arquivo:
    echo  %~dp0gerar_dados_js.log
)

echo.
echo  Esta janela fechara em 5 segundos...
powershell -NoProfile -Command "Start-Sleep -Seconds 5"
