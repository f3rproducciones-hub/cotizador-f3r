@echo off
title F3R PRODUCCIONES S.A.S. - CRM COMERCIAL
echo ======================================================================
echo           INICIANDO CRM Y COTIZADOR F3R PRODUCCIONES S.A.S.
echo ======================================================================
echo.
echo Abriendo servidor local en el puerto 8502...
echo En tu navegador ingresa a: http://localhost:8502
echo.
py -m streamlit run app.py --server.port 8502
pause
