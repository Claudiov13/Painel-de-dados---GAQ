@echo off
title Sistema de Compras - GAQ
cd /d "%~dp0"
set NODE="%USERPROFILE%\node\bin\node.exe"

:loop
echo.
echo  Iniciando Sistema de Compras - Painel GAQ...
echo  [%date% %time%]
echo.
%NODE% server.js 2>&1
echo.
echo  [%date% %time%] Servidor encerrado (codigo: %errorlevel%). Reiniciando em 5 segundos...
echo  (Feche esta janela para parar definitivamente)
echo.
timeout /t 5 /nobreak >nul
goto loop
