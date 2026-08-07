@echo off
chcp 65001 >nul
title Configuración inicial - Dashboard Auditorías

echo.
echo ╔══════════════════════════════════════════════════╗
echo ║     CONFIGURACIÓN INICIAL (Solo una vez)         ║
echo ╚══════════════════════════════════════════════════╝
echo.
echo Este script configura Git y conecta tu carpeta
echo con tu repositorio de GitHub.
echo.

:: ── Verificar Git ────────────────────────────────────────────
git --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Git no está instalado.
    echo.
    echo Descárgalo desde: https://git-scm.com/download/win
    echo Instálalo y vuelve a ejecutar este script.
    pause
    exit /b 1
)
echo ✓ Git encontrado
echo.

:: ── Pedir usuario de GitHub ──────────────────────────────────
set /p GITHUB_USER="Ingresa tu usuario de GitHub: "
if "%GITHUB_USER%"=="" (
    echo [ERROR] Debes ingresar tu usuario de GitHub
    pause
    exit /b 1
)

set REPO_URL=https://github.com/%GITHUB_USER%/dashboard-auditorias.git
echo.
echo Se conectará al repositorio:
echo   %REPO_URL%
echo.
echo IMPORTANTE: Debes crear ese repositorio en GitHub primero:
echo   1. Ve a https://github.com/new
echo   2. Nombre del repositorio: dashboard-auditorias
echo   3. Déjalo PÚBLICO
echo   4. NO marques "Add README"
echo   5. Haz clic en "Create repository"
echo.
pause

:: ── Configurar identidad Git ─────────────────────────────────
set /p GIT_EMAIL="Ingresa tu email de GitHub: "
set /p GIT_NOMBRE="Ingresa tu nombre: "
git config --global user.email "%GIT_EMAIL%"
git config --global user.name "%GIT_NOMBRE%"
echo ✓ Identidad Git configurada
echo.

:: ── Inicializar o verificar repositorio ──────────────────────
set CARPETA=%~dp0

if exist "%CARPETA%.git" (
    echo ✓ Repositorio Git ya existe, actualizando URL...
    git -C "%CARPETA%" remote set-url origin "%REPO_URL%" >nul 2>&1
) else (
    echo Inicializando repositorio Git...
    git -C "%CARPETA%" init >nul
    git -C "%CARPETA%" remote add origin "%REPO_URL%"
)

:: ── Primer commit y push ─────────────────────────────────────
echo.
echo Subiendo todos los archivos a GitHub...
echo (Se te pedirá usuario y contraseña/token de GitHub)
echo.

git -C "%CARPETA%" add .
git -C "%CARPETA%" commit -m "Configuración inicial del dashboard"
git -C "%CARPETA%" branch -M main
git -C "%CARPETA%" push -u origin main

if errorlevel 1 (
    echo.
    echo [ERROR] No se pudo subir a GitHub.
    echo.
    echo Si te pide contraseña, debes usar un TOKEN de GitHub:
    echo   1. Ve a https://github.com/settings/tokens
    echo   2. Haz clic en "Generate new token (classic)"
    echo   3. Marca el permiso "repo"
    echo   4. Copia el token y úsalo como contraseña
    echo.
    pause
    exit /b 1
)

:: ── Activar GitHub Pages ─────────────────────────────────────
echo.
echo ╔══════════════════════════════════════════════════╗
echo ║  ✓  Archivos subidos a GitHub                   ║
echo ╚══════════════════════════════════════════════════╝
echo.
echo ÚLTIMO PASO - Activa GitHub Pages:
echo   1. Ve a https://github.com/%GITHUB_USER%/dashboard-auditorias
echo   2. Haz clic en "Settings" (Configuración)
echo   3. En el menú izquierdo: "Pages"
echo   4. En "Branch" selecciona "main" y carpeta "/ (root)"
echo   5. Haz clic en "Save"
echo.
echo En ~2 minutos tu sitio estará en:
echo   https://%GITHUB_USER%.github.io/dashboard-auditorias/
echo.

:: Guardar la URL en un archivo de texto para referencia
echo Tu URL del dashboard: https://%GITHUB_USER%.github.io/dashboard-auditorias/ > mi_url.txt
echo Usuario de GitHub: %GITHUB_USER% >> mi_url.txt
echo. >> mi_url.txt
echo Para actualizar los datos diariamente, ejecuta: actualizar.bat >> mi_url.txt
echo Archivo mi_url.txt creado con tu URL ✓
echo.

start https://github.com/%GITHUB_USER%/dashboard-auditorias/settings/pages

pause
