@echo off
echo =======================================================
echo.
echo HOLA, SOY TU ASISTENTE.
echo.
echo Sigamos estos pasos para conectar tu cuenta:
echo 1. Se abrira tu navegador web con una pagina de GitHub.
echo 2. Pega el codigo de 8 letras/numeros que aparecera abajo.
echo 3. Autoriza la aplicacion y vuelve a esta ventana.
echo.
echo =======================================================
echo.
gh auth login --hostname github.com --git-protocol https --web
echo.
echo.
echo EXCELENTE! YA ESTAS CONECTADO.
echo CIERRA ESTA VENTANA NEGRA Y AVISAME EN EL CHAT.
echo.
pause
