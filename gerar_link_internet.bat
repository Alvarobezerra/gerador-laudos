@echo off
echo ==========================================================
echo GERADOR DE LINK PARA ACESSO EXTERNO (INTERNET)
echo ==========================================================
echo.
echo Este script cria um tunel seguro (HTTPS) usando Serveo.
echo Isso permite que voce acesse o app PWA de QUALQUER LUGAR,
echo mesmo no 4G da viatura, sem estar no Wi-Fi da delegacia!
echo.
echo IMPORTANTE: O "start_all.bat" precisa estar rodando antes.
echo.
echo ==========================================================
echo Copie o link verde "https://..." que vai aparecer abaixo
echo e acesse no celular:
echo ==========================================================
echo.
ssh -o StrictHostKeyChecking=no -R 80:localhost:8501 nokey@localhost.run
pause
