@echo off
title Subiendo a GitHub...
echo =======================================================
echo.
echo HOLA, SOY TU ASISTENTE. 
echo ESTOY SUBIENDO TUS ARCHIVOS A GITHUB AUTOMATICAMENTE...
echo.
echo =======================================================
echo.

cd C:\AuditoriasTerreno
git config --global user.email "jvodnizza@gmail.com"
git config --global user.name "JhonasVK"

if not exist ".git" (
    git init
    git remote add origin https://github.com/JhonasVK/dashboard-auditorias.git
) else (
    git remote set-url origin https://github.com/JhonasVK/dashboard-auditorias.git
)

git add .
git commit -m "Configuración inicial del dashboard"
git branch -M main
git push -u origin main

echo.
echo =======================================================
echo.
echo EXCELENTE. LA SUBIDA HA TERMINADO.
echo AHORA SOLO QUEDA ACTIVAR GITHUB PAGES.
echo SE ABRIRA TU NAVEGADOR EN LA PAGINA CORRECTA.
echo.
pause
start https://github.com/JhonasVK/dashboard-auditorias/settings/pages
exit
