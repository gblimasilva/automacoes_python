@echo off
title Automação Completa - Execução
echo ===========================================
echo     INICIANDO PROCESSO AUTOMATIZADO...
echo ===========================================
echo.

:: Caminho do Python que você usa
set PYTHON_PATH=C:\Users\User\AppData\Local\Programs\Python\Python314\python.exe

:: Caminho do script principal
set SCRIPT_PATH=C:\Users\User\Desktop\automacoes_python\automacao_completa.py

echo Executando automação completa...
"%PYTHON_PATH%" "%SCRIPT_PATH%"
echo.
echo ===========================================
echo     PROCESSO FINALIZADO!
echo ===========================================
pause
