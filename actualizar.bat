@echo off
title Actualizador Dashboard Auditorias

echo ========================================================
echo      ACTUALIZACION DASHBOARD AUDITORIAS
echo ========================================================
echo.

set CARPETA=C:\AuditoriasTerreno
set EXCEL=%CARPETA%\data\BBDD_Supervisores.xlsx

:: -- Verificar que existe el Excel --------------------------
if not exist "%EXCEL%" (
    echo [ERROR] No se encontro el archivo Excel en:
    echo.
    echo         C:\AuditoriasTerreno\data\BBDD_Supervisores.xlsx
    echo.
    echo Asegurate que el archivo se llame exactamente "BBDD_Supervisores.xlsx"
    echo.
    pause
    exit /b 1
)

:: -- Verificar que git esta instalado ------------------------
git --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Git no esta instalado.
    echo.
    echo Descargalo desde: https://git-scm.com/download/win
    pause
    exit /b 1
)

:: -- Asegurar que la cuenta activa de GitHub sea JhonasVK ------
:: (otros scripts, como el de Reincidencias, dejan activada otra
:: cuenta al terminar; esto evita el error 403 al subir)
gh auth switch --hostname github.com --user JhonasVK >nul 2>nul

:: -- Git: agregar, confirmar y subir --------------------------
echo [1/3] Registrando cambios...
git -C "%CARPETA%" add data/BBDD_Supervisores.xlsx
echo       OK

echo [2/3] Guardando version con fecha y hora...
for /f "tokens=1-3 delims=/" %%a in ('date /t') do set FECHA=%%c-%%b-%%a
for /f "tokens=1-2 delims=: " %%a in ('time /t') do set HORA=%%a:%%b
git -C "%CARPETA%" commit -m "Actualizacion %FECHA% %HORA%"
echo       OK

echo [3/3] Subiendo a GitHub...
git -C "%CARPETA%" push origin main
if errorlevel 1 (
    echo.
    echo [ERROR] No se pudo subir. Verifica tu conexion o token de GitHub.
    pause
    exit /b 1
)

echo.
echo ========================================================
echo   Listo! Pagina actualizada correctamente
echo ========================================================
echo.
echo  En ~1 minuto veras los nuevos datos en tu pagina web.
echo.
pause
