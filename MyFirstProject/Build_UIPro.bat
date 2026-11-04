@echo off
title Build Civil Tool - UI Pro
cd /d "%~dp0"
echo Dang bien dich MyFirstProject (UI Pro)...
dotnet build MyFirstProject.csproj -c Debug -nologo > "%~dp0build_log_uipro.txt" 2>&1
if %errorlevel% equ 0 (echo BUILD THANH CONG!) else (echo CO LOI - xem build_log_uipro.txt)
timeout /t 5 > nul
