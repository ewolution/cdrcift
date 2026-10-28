@echo off
rem cdrcift - CorelDRAW Cift Cizgi Denetleyicisi
rem Cift tiklayarak kurulum. Yonetici hakki gerekmez.
title cdrcift kurulum
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0kur.ps1" %*
echo.
pause
