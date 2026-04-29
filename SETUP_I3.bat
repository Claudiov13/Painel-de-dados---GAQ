@echo off
:: Executa o SETUP_I3.ps1 como Administrador
echo Iniciando setup do servidor dedicado GAQ...
echo.
powershell -ExecutionPolicy Bypass -Command "Start-Process powershell -ArgumentList '-ExecutionPolicy Bypass -File ""%~dp0SETUP_I3.ps1""' -Verb RunAs -Wait"
pause
