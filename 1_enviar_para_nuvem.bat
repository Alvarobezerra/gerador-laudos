@echo off
chcp 65001 >nul
echo ==================================================
echo. ENVIAR APLICATIVO PARA O GITHUB (NUVEM)
echo ==================================================
echo.
echo Antes de continuar, certifique-se de que voce:
echo 1. Criou uma conta no site github.com
echo 2. Clicou em "New repository" e criou um repositorio vazio
echo.
set /p giturl="Cole aqui o link do seu repositorio (ex: https://github.com/SeuUsuario/gerador-laudos.git): "
echo.
echo Configurando e enviando seus arquivos...
git remote add origin "%giturl%"
git branch -M main
git push -u origin main
echo.
echo ==================================================
echo PRONTO! Seus arquivos foram enviados para o GitHub.
echo Agora acesse share.streamlit.io, faca login com seu GitHub,
echo clique em "New app" e selecione o seu repositorio para publicar!
echo ==================================================
pause
