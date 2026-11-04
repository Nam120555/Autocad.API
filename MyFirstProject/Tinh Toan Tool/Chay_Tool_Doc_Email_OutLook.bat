@echo off
chcp 65001 > NUL
title TOOL TRÍCH XUẤT EMAIL OUTLOOK TỰ ĐỘNG
cd /d "C:\OutLook"
if exist "OutLookTool.exe" (
    start "" "C:\OutLook\OutLookTool.exe"
) else (
    echo [LOI] Khong tim thay C:\OutLook\OutLookTool.exe
    pause
)
