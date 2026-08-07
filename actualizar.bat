@echo off
chcp 65001 >nul
title Actualizador Dashboard Auditorías

echo.
echo ╔══════════════════════════════════════════════════╗
echo ║     ACTUALIZACIÓN DASHBOARD AUDITORÍAS           ║
echo ╚══════════════════════════════════════════════════╝
echo.

set CARPETA=C:\AuditoriasTerreno
set EXCEL=%CARPETA%\data\BBDD_Supervisores.xlsx

:: ── Verificar que existe el Excel ─────────────────────────────
if not exist "%EXCEL%" (
    echo [ERROR] No se encontró el archivo Excel en:
    echo.
    echo         C:\AuditoriasTerreno\data\BBDD_Supervisores.xlsx
    echo.
    echo Asegúrate que el archivo se llame exactamente "BBDD_Supervisores.xlsx"
    echo.
    pause
    exit /b 1
)

:: ── Verificar que git está instalado ─────────────────────────
git --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Git no está instalado.
    echo.
    echo Descárgalo desde: https://git-scm.com/download/win
    pause
    exit /b 1
)

:: ── Git: agregar, confirmar y subir ──────────────────────────
echo [1/3] Registrando cambios...
git -C "%CARPETA%" add data/BBDD_Supervisores.xlsx >nul 2>&1
echo       OK

echo [2/3] Guardando versión con fecha y hora...
for /f "tokens=1-3 delims=/" %%a in ('date /t') do set FECHA=%%c-%%b-%%a
for /f "tokens=1-2 delims=: " %%a in ('time /t') do set HORA=%%a:%%b
git -C "%CARPETA%" commit -m "Actualización %FECHA% %HORA%" >nul 2>&1
echo       OK

echo [3/3] Subiendo a GitHub...
git -C "%CARPETA%" push origin main
if errorlevel 1 (
    echo.
    echo [ERROR] No se pudo subir. Verifica tu conexión o token de GitHub.
    pause
    exit /b 1
)

echo.
echo ╔══════════════════════════════════════════════════╗
echo ║  ✓  ¡Listo! Página actualizada correctamente    ║
echo ╚══════════════════════════════════════════════════╝
echo.
echo  En ~1 minuto verás los nuevos datos en tu página web.
echo.
timeout /t 6
