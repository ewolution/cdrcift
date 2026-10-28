@echo off
rem cdrcift - kurulumu kaldirir
title cdrcift kaldirma
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0kur.ps1" -Kaldir %*
echo.
pause
