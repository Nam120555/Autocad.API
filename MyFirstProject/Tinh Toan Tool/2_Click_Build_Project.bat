@echo off
title Build AutoCAD Civil 3D Plugin - MyFirstProject
color 0B
echo ====================================================================
echo        DANG BIEN DICH CODE SANG FILE DLL CHO AUTOCAD...
echo        Vui long doi trong giay lat...
echo ====================================================================
echo.

cd /d "%~dp0\MyFirstProject"

:: Luu log vao file de AI doc duoc
dotnet build MyFirstProject.csproj -c Debug > "%~dp0\build_log.txt" 2>&1

echo ====================================================================
if %errorlevel% equ 0 (
    color 0A
    echo    BUILD THANH CONG! TAI THAC DONG OK!
    echo    (Luu y: Vui long tat AutoCAD neu tool bao loi file chong cheo)
) else (
    color 0C
    echo    CO LOI XAY RA TRONG QUA TRINH BIEN DICH!
    echo    Toi da luu ban ghi loi vao file "build_log.txt".
    echo    Ban phan hoi lai tren chat de toi kiem tra log tu dong nhe!
)
echo ====================================================================
echo Nhan phim bat ky de thoat...
pause > nul
