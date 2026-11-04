@echo off
color 0A
title Cong CU Dich Nguoc NXSoft
echo =======================================================
echo TU DONG CAI DAT ILSPY VA DECOMPILE NXSOFT_DOTNET.DLL
echo =======================================================
echo.
echo [1/2] Dang cai dat/cap nhat cong cu ILSpyCmd...
dotnet tool install --global ilspycmd 2>nul
if %errorlevel% neq 0 (
    dotnet tool update --global ilspycmd 2>nul
)

echo.
echo [2/2] Dang tien hanh dich nguoc (decompile) C:\NXsoft\NXsoft_DotNet.dll...
echo Vui long cho trong it phut de phan mem hoat dong. Toan bo ma nguon se luu thu muc NXSOFT_Source...
md "%~dp0NXSOFT_Source" 2>nul

:: Goi ilspycmd de luu ra project
ilspycmd -p -o "%~dp0NXSOFT_Source" "C:\NXsoft\NXsoft_DotNet.dll"

echo.
echo =============================================================
echo HOAN THANH! Ma nguon da duoc xuat ra thu muc: NXSOFT_Source
echo Hay quay lai Chat va bao: "Da xong, hay doc di" nhe! 
echo =============================================================
pause

