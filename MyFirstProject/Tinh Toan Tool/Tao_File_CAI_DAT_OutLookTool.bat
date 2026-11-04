@echo off
title Tao Bo Cai Dat Cho Phan Mem Loc Email Outlook - NXSOFT
color 0B
chcp 65001 >nul

echo =============================================================
echo   TẠO BỘ CÀI ĐẶT WINDOWS (.EXE) CHO PHẦN MỀM LỌC EMAIL OUTLOOK
echo   Tác Giả: Huỳnh Ngọc Nam - NXSOFT
echo =============================================================
echo.

set "PROJECT_DIR=D:\Autocad 2026_API\Autocad 2026_API"
set "OUTPUT_DIR=D:\Autocad 2026_API\OutLookTool_Installer"

:: 1. Build & Publish Release OutLookTool
echo [1/2] Đang biên dịch phần mềm chính OutLookTool.exe...
cd /d "%PROJECT_DIR%\OutLookTool"
dotnet publish -c Release -r win-x64 --self-contained false -p:PublishSingleFile=true -o "C:\OutLook" --nologo -v quiet

if %errorlevel% neq 0 (
    color 0C
    echo [LỖI] Biên dịch OutLookTool.exe thất bại! Vui lòng kiểm tra lại.
    pause
    exit /b 1
)
echo [OK] Đã xuất OutLookTool.exe vào C:\OutLook\
echo.

:: 2. Build & Publish Setup Installer App
echo [2/2] Đang đóng gói bộ cài đặt Setup_OutLookTool_v1.0.0.exe...
cd /d "%PROJECT_DIR%\Installer\OutLookTool_SetupApp"
if exist "%OUTPUT_DIR%" rmdir /s /q "%OUTPUT_DIR%"
mkdir "%OUTPUT_DIR%" 2>nul

dotnet publish -c Release -r win-x64 --self-contained false -p:PublishSingleFile=true -o "%OUTPUT_DIR%" --nologo -v quiet

if %errorlevel% neq 0 (
    color 0C
    echo [LỖI] Đóng gói bộ cài đặt thất bại!
    pause
    exit /b 1
)

echo.
color 0A
echo =============================================================
echo  THÀNH CÔNG! ĐÃ TẠO BỘ CÀI ĐẶT PHẦN MỀM LỌC EMAIL OUTLOOK
echo.
echo  Vị trí bộ cài đặt:
dir "%OUTPUT_DIR%\*.exe" /b 2>nul
echo =============================================================
echo.
echo Bấm phím bất kỳ để mở thư mục chứa bộ cài đặt...
pause > nul
explorer "%OUTPUT_DIR%"
