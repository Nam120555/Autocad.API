# =====================================================================
#  GIAI DOAN 1 + 2 - Build chuan, phien ban, log loi
#  Chay bang NangCap_GiaiDoan1_2.bat
# =====================================================================
$ErrorActionPreference = 'Continue'
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root
$Proj = Join-Path $Root 'MyFirstProject\MyFirstProject.csproj'
$Log  = Join-Path $Root '_build_check.log'
function Step($t) { Write-Host "`n=== $t ===" -ForegroundColor Cyan }
function Ok($t)   { Write-Host "  [OK] $t" -ForegroundColor Green }
function Warn($t) { Write-Host "  [!] $t" -ForegroundColor Yellow }
function Stop-Here($t) { Write-Host "  [LOI] $t" -ForegroundColor Red; Start-Sleep 600; exit 1 }

$acadOpen = [bool](Get-Process acad -ErrorAction SilentlyContinue)

Step '1/4  Build Debug (ban dung hang ngay / NETRELOAD)'
if ($acadOpen) { Warn 'AutoCAD dang mo (DLL Debug co the bi khoa) -> bo qua build Debug, chi kiem tra bang build Release.' } else {
dotnet build $Proj -c Debug -nologo -v minimal *> $Log
if ($LASTEXITCODE -ne 0) {
  Get-Content $Log | Select-String ' error ' | Select-Object -First 15 | ForEach-Object { Write-Host $_.Line -ForegroundColor Red }
  Warn "Chua commit gi ca. File log loi: $Log  (gui cho Claude de sua)"
  Warn 'Muon huy toan bo thay doi giai doan 1-2:  git checkout -- .  &  git clean -fd MyFirstProject\Core'
  Stop-Here 'Build Debug loi'
}
$w = (Get-Content $Log | Select-String 'Warning\(s\)' | Select-Object -Last 1).Line
Ok "Build Debug thanh cong. $w  (truoc day: 4300 canh bao)"
Copy-Item $Log (Join-Path $Root '_build_debug.log') -Force
}

Step '2/4  Build Release + tao bo cai'
& (Join-Path $Root 'build.ps1')
if ($LASTEXITCODE -ne 0) { Stop-Here "Build Release loi - xem $Log" }

Step '3/4  Don script cu'
$old = Join-Path $Root 'MyFirstProject\Tinh Toan Tool\3_Tao_File_CAI_DAT_CivilTool.bat'
if (Test-Path $old) { git rm -q -- $old 2>$null; if (Test-Path $old) { Remove-Item $old -Force }; Ok 'Xoa 3_Tao_File_CAI_DAT_CivilTool.bat (thay bang build.bat)' }
foreach ($f in 'DonDep_GiaiDoan0.ps1','DonDep_GiaiDoan0.bat') { $p = Join-Path $Root $f; if (Test-Path $p) { Remove-Item $p -Force; Ok "Xoa $f (da xong)" } }

Step '4/4  Commit + tag v1.2.0'
git add -A
git -c core.quotepath=false commit -q -m "feat: giai doan 1-2 - v1.2.0, build Release, bo tham chieu AutoCAD 2027 gay xung dot, NxLog/AppInfo, thay 143 catch rong bang ghi log, lenh CT_LOG"
if ($LASTEXITCODE -eq 0) { Ok 'Da commit' } else { Warn 'Khong co gi de commit' }
git tag -f v1.2.0 | Out-Null
Ok 'Tag v1.2.0'
git log --oneline -3
Write-Host "`nHOAN TAT. Mo Civil 3D, kiem tra dong chao co hien 'v1.2.0', go thu CT_LOG." -ForegroundColor Green
Write-Host 'Khi on:  git push origin master --tags' -ForegroundColor Green
Write-Host "`nCua so se tu dong sau 10 phut."; Start-Sleep 600
