@echo off
color 0A
title TRINH BIEN DICH PHAN MEM - PAVEMENT DESIGN APP
echo ========================================================
echo DANG TIEN HANH BIEN DICH PHAN MEM SANG FILE .EXE
echo Khong can mo Visual Studio - Su dung truc tiep .NET SDK
echo ========================================================
echo.

:: Di chuyen den dung thu muc chua project
cd /d "%~dp0"

:: Xoa thu muc build cu neu co
if exist "App_HoanThien" rmdir /s /q "App_HoanThien"

:: Chay lenh publish de tao file EXE, dong goi dep de nhat
echo Dang ban giao lenh Build cho he thong...
dotnet publish "PavementDesignApp.csproj" -c Release -r win-x64 --self-contained false -p:PublishSingleFile=true -o "App_HoanThien"

echo.
if exist "App_HoanThien\PavementDesignApp.exe" (
    echo ========================================================
    echo THANH CONG CHUC MUNG ANH!
    echo File chay PavementDesignApp.exe da duoc tao tai:
    echo %~dp0App_HoanThien\
    echo ========================================================
    echo Dang tu dong mo thu muc...
    explorer "App_HoanThien"
) else (
    echo [LOI] Co van de xay ra trong qua trinh build, hay kiem tra lai.
    color 0C
)

pause
