@echo off
title Sistema de Laudos - PRODUCAO (Porta 8501)
echo Iniciando ambiente de PRODUCAO na porta 8501 para os usuarios...
streamlit run app.py --server.port 8501
pause
