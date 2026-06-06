@echo off
chcp 65001 >nul
title Actualizando Dashboard Google Ads...
echo.
echo  Leyendo GOOGLE.xlsx...
powershell -ExecutionPolicy Bypass -NoProfile -File "%~dp0actualizar.ps1"
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo  ERROR al procesar el archivo. Verifica que Excel este cerrado.
    pause
    exit /b 1
)
echo.
echo  Listo. Regresa al navegador y presiona F5.
timeout /t 3 /nobreak >nul
