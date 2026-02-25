@echo off
chcp 65001 > nul
title Sistema Processos v5.0 — Light Mode
color 1F
cls
echo.
echo  ╔══════════════════════════════════════════════════════════╗
echo  ║  SISTEMA DE PROCESSOS DE PAGAMENTO  v5.0                ║
echo  ║  Interface Web — Light Mode Moderno                      ║
echo  ║  Controladoria Geral — Gov. Edison Lobão/MA             ║
echo  ╚══════════════════════════════════════════════════════════╝
echo.
python --version 2>nul
if errorlevel 1 (
    echo  ERRO: Python nao encontrado.
    echo  Instale em: https://python.org/downloads
    echo  IMPORTANTE: Marque "Add Python to PATH"
    pause & exit /b 1
)
echo  [1/2] Instalando dependencias...
pip install -r requirements.txt --quiet --no-warn-script-location
echo  [2/2] Iniciando sistema...
echo.
echo  Acesso local:    http://localhost:8501
echo  Acesso na rede:  Use o IP mostrado abaixo
echo  Login padrao:    admin / admin123
echo.
echo  Para fechar: pressione CTRL+C ou feche esta janela.
echo.
streamlit run controle_pagamentos.py --server.address 0.0.0.0 --server.port 8501
pause
