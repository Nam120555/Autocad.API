@echo off
title Tao File Cai Dat Cho Civil Tool - NXSOFT
color 0B

echo =============================================================
echo   TAO FILE CAI DAT WINDOWS (.exe) CHO PLUGIN CIVIL 3D
echo   Tac Gia: Huynh Ngoc Nam
echo =============================================================
echo.

:: Dinh nghia cac duong dan
set "PROJECT_DIR=D:\Autocad 2026_API\Autocad 2026_API"
set "BUILD_DIR=%PROJECT_DIR%\MyFirstProject\bin\Debug"
set "ISS_FILE=%PROJECT_DIR%\Installer\setup_civiltool.iss"
set "OUTPUT_DIR=D:\Autocad 2026_API\Civil_Tool_Installer"

:: Tim Inno Setup Compiler (iscc.exe)
set "ISCC="
if exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" set "ISCC=C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if exist "C:\Program Files\Inno Setup 6\ISCC.exe"       set "ISCC=C:\Program Files\Inno Setup 6\ISCC.exe"
if exist "%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe" set "ISCC=%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe"

if "%ISCC%"=="" (
    color 0C
    echo [LOI] Khong tim thay Inno Setup. Ban can cai dat Inno Setup truoc.
    pause
    exit /b 1
)

:: 1. Build Project
echo [1/3] Dang bien dich Civil Tool plugin...
cd /d "%PROJECT_DIR%\MyFirstProject"
dotnet build "MyFirstProject.csproj" -c Debug --nologo -v quiet

if %errorlevel% neq 0 (
    color 0C
    echo [LOI] Build that bai! Vui long kiem tra lai code.
    pause
    exit /b 1
)
echo [OK] Bien dich hoan tat.
echo.

:: 2. Kiem tra thu muc
echo [2/3] Kiem tra file...
if not exist "%BUILD_DIR%\MyFirstProject.dll" (
    color 0C
    echo [LOI] Khong tim thay MyFirstProject.dll.
    pause
    exit /b 1
)
echo [OK] File da san sang.
echo.

:: 3. Chay Inno Setup de tao bo cai
echo [3/3] Dang goi Inno Setup de tao file .exe...
if exist "%OUTPUT_DIR%" rmdir /s /q "%OUTPUT_DIR%"
mkdir "%OUTPUT_DIR%" 2>nul

cd /d "%PROJECT_DIR%"
"%ISCC%" "%ISS_FILE%"

echo.
if %errorlevel% equ 0 (
    color 0A
    echo =============================================================
    echo  THANH CONG! DA TAO BO CAI DAT CIVIL TOOL THUC THU
    echo.
    echo  Vi tri bo cai:
    dir "%OUTPUT_DIR%\*.exe" /b 2>nul
    echo.
    echo  UU DIEM TU DONG KHI CAI DAT BO NAY:
    echo  - Tu dong cai plugin vao ApplicationPlugins
    echo  - Tu dong Load vao AutoCAD Civil 3D ma KHONG CAN chuot phai NETLOAD
    echo  - Loai tru tat ca dll thua cua Autocad tranh lam crash phan mem
    echo =============================================================
    echo.
    echo Bam phim bat ky de mo thu muc chua file cai dat...
    pause > nul
    explorer "%OUTPUT_DIR%"
) else (
    color 0C
    echo =============================================================
    echo  THAT BAI! Chay Inno Setup bi loi.
    echo =============================================================
    pause
)
