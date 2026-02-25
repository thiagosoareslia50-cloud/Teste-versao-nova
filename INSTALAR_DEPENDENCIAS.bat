@echo off
chcp 65001 > nul
title Instalador v5.0
echo Instalando dependencias v5.0...
python -m pip install --upgrade pip --quiet
pip install streamlit>=1.32.0
pip install pandas>=2.0.0
pip install openpyxl>=3.1.0
pip install reportlab>=4.0.0
pip install plotly>=5.0.0
pip install gspread>=6.0.0
pip install google-auth>=2.28.0
echo.
echo Instalacao concluida! Execute INICIAR.bat
pause
