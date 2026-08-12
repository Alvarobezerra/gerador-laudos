@echo off
echo Iniciando Servidor API e PWA Mobile...
start /b uvicorn api:app --host 0.0.0.0 --port 8000

echo Iniciando Streamlit...
start /b streamlit run app.py

echo.
echo ====================================================
echo SERVIDORES INICIADOS!
echo ====================================================
echo.
echo - Para acessar o PWA no CELULAR, descubra o IP deste computador
echo   e acesse: http://SEU-IP:8000
echo.
echo - Para acessar o Streamlit no PC:
echo   http://localhost:8501
echo.
echo Pressione qualquer tecla para encerrar os servidores...
pause > nul
taskkill /f /im python.exe > nul 2>&1
echo Encerrado.
