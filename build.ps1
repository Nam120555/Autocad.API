# =====================================================================
#  build.ps1 - Build Release + tao bo cai dat Civil Tool (thay cho 3_Tao_File_CAI_DAT_CivilTool.bat)
#  Phien ban lay tu <Version> trong MyFirstProject\MyFirstProject.csproj
#
#  Cach dung (PowerShell tai thu muc nay):
#     .\build.ps1              -> build Release + tao bo cai
#     .\build.ps1 -NoInstaller -> chi build Release
# =====================================================================
param([switch]$NoInstaller)
$ErrorActionPreference = 'Continue'
$Root    = Split-Path -Parent $MyInvocation.MyCommand.Path
$Proj    = Join-Path $Root 'MyFirstProject\MyFirstProject.csproj'
$OutDir  = Join-Path $Root 'MyFirstProject\bin\Release'
$Iss     = Join-Path $Root 'Installer\setup_civiltool.iss'
$PkgXml  = Join-Path $Root 'Installer\PackageContents.xml'
$Log     = Join-Path $Root '_build_check.log'

function Ok($t)   { Write-Host "  [OK] $t" -ForegroundColor Green }
function Fail($t) { Write-Host "  [LOI] $t" -ForegroundColor Red; exit 1 }

# 1. Doc version
[xml]$csproj = Get-Content $Proj -Encoding UTF8
$Version = ($csproj.Project.PropertyGroup | Where-Object { $_.Version } | Select-Object -First 1).Version
if (-not $Version) { Fail 'Khong doc duoc <Version> trong csproj' }
Write-Host "`n=== Civil Tool v$Version ===" -ForegroundColor Cyan

# 2. Build Release
Write-Host '[1/3] Build Release...'
dotnet build $Proj -c Release -nologo -v minimal *> $Log
if ($LASTEXITCODE -ne 0) {
    Get-Content $Log | Select-String ' error ' | Select-Object -First 15 | ForEach-Object { Write-Host $_.Line -ForegroundColor Red }
    Fail "Build loi - xem chi tiet: $Log"
}
$warn = (Get-Content $Log | Select-String 'Warning\(s\)' | Select-Object -Last 1).Line
Ok "Build thanh cong. $warn"
if (-not (Test-Path (Join-Path $OutDir 'MyFirstProject.dll'))) { Fail "Khong thay MyFirstProject.dll trong $OutDir" }
if ($NoInstaller) { exit 0 }

# 3. Cap nhat version trong PackageContents.xml
Write-Host '[2/3] Cap nhat PackageContents.xml...'
$pc = Get-Content $PkgXml -Raw -Encoding UTF8
$pc = $pc -replace 'AppVersion="[^"]*"', "AppVersion=`"$Version`"" -replace '(<ComponentEntry[^>]*?)Version="[^"]*"', "`$1Version=`"$Version`""
[IO.File]::WriteAllText($PkgXml, $pc, (New-Object Text.UTF8Encoding $false))
Ok "Version = $Version"

# 4. Inno Setup
Write-Host '[3/3] Tao bo cai dat...'
$iscc = @("${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe", "$env:ProgramFiles\Inno Setup 6\ISCC.exe", "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe") |
        Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $iscc) { Write-Host '  [!] Chua cai Inno Setup 6 -> bo qua buoc tao bo cai (DLL Release van dung duoc).' -ForegroundColor Yellow; exit 0 }
Push-Location (Join-Path $Root 'Installer')
& $iscc "/DAppVersion=$Version" /Q $Iss
$code = $LASTEXITCODE
Pop-Location
if ($code -ne 0) { Fail 'Inno Setup bao loi' }
Ok "Bo cai: D:\Autocad 2026_API\Civil_Tool_Installer\Civil_Tool_NXSOFT_Setup_v$Version.exe"
