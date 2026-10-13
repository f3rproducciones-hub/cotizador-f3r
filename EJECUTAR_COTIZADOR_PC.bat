@echo off
chcp 65001 > nul
title F3R Producciones - Cotizador Inteligente y CRM PC
echo ======================================================================
echo           F3R PRODUCCIONES S.A.S. - SUITE COMERCIAL Y CRM
echo                     MODO LOCAL PARA COMPUTADOR (PC)
echo ======================================================================
echo.
echo Iniciando servidor local...
echo Tu navegador se abrira automaticamente en http://localhost:8501
echo.
start "" "http://localhost:8501"
py -m streamlit run app.py
if %ERRORLEVEL% NEQ 0 (
    python -m streamlit run app.py
)
pause
