# =====================================================================
#  GIAI DOAN 0 - DON DEP DU AN CIVIL TOOL
#  Chay bang file DonDep_GiaiDoan0.bat (dat cung thu muc)
# =====================================================================
$ErrorActionPreference = 'Continue'
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root
$Proj = Join-Path $Root 'MyFirstProject'

function Step($t) { Write-Host "`n=== $t ===" -ForegroundColor Cyan }
function Ok($t)   { Write-Host "  [OK] $t" -ForegroundColor Green }
function Warn($t) { Write-Host "  [!] $t" -ForegroundColor Yellow }
function Fail($t) { Write-Host "  [LOI] $t" -ForegroundColor Red; Read-Host "Bam Enter de thoat"; exit 1 }

if (-not (Get-Command git -ErrorAction SilentlyContinue)) { Fail 'Khong tim thay git. Cai Git for Windows truoc.' }
if (-not (Test-Path (Join-Path $Root '.git'))) { Fail "Khong thay thu muc .git trong $Root" }
$acad = Get-Process acad -ErrorAction SilentlyContinue
if ($acad) { Fail 'AutoCAD/Civil 3D dang mo - hay tat AutoCAD roi chay lai (DLL dang bi khoa).' }

# ---------------------------------------------------------------------
Step '1/5  Bo cac thu muc build/IDE khoi git (van giu tren o dia)'
$old = Join-Path $Proj 'Tinh Toan Tool\.gitignore'
if (Test-Path $old) { Remove-Item $old -Force; Ok '.gitignore cu (Tinh Toan Tool) da xoa - dung .gitignore o thu muc goc' }
git rm -r -q --cached --ignore-unmatch -- `
    '*/bin/*' '*/obj/*' '.vs/*' '*.user' '*.pdb' '*.log' 'MyFirstProject/build_log*.txt' 2>$null | Out-Null
Ok 'Da go bin/ obj/ .vs/ log khoi theo doi cua git'

# ---------------------------------------------------------------------
Step '2/5  Commit LUU TOAN BO cong viec hien tai (moc an toan)'
git add -A
git -c core.quotepath=false commit -q -m "chore: luu toan bo thay doi tu 04/2026 truoc khi don dep (moc an toan)" 2>$null
if ($LASTEXITCODE -eq 0) { Ok 'Da tao commit moc an toan' } else { Warn 'Khong co gi moi de commit (bo qua)' }
git tag -f truoc-don-dep | Out-Null
Ok "Da gan tag 'truoc-don-dep' - muon quay lai: git checkout truoc-don-dep"

# ---------------------------------------------------------------------
Step '3/5  Xoa lenh rong / file rac'
$xoa = @(
  # Lenh vo rong (chi in thong bao, khong lam gi)
  'Civil Tool\36.CTP_FixTDTN.cs',
  'Civil Tool\37.CTSV_ChenhCaoTimCoc.cs',
  'Civil Tool\38.CTPI_ThietKePipeNetwork.cs',
  'Civil Tool\39.CTA_NhapNTD.cs',
  'Civil Tool\40.CTP_DienCaoDoThietKe.cs',
  # File chi con comment / class rong
  'Civil Tool\14.SanNen.cs',
  'Civil Tool\26.BanKinhRePhai.cs',
  'Civil Tool\20.NXsoftMasterSuite.cs',
  'Tinh Toan Tool\01.BanKinhRePhai.cs',
  'Tinh Toan Tool\TinhToan\TT_ExcelGenerator.cs',
  # Form mau cua template (lenh ShowForm)
  'Menu form\TestForm.cs',
  'Menu form\TestForm.Designer.cs',
  'Menu form\TestForm.resx',
  # File 0 byte / khong duoi
  'Civil Tool\41.BANG DANH MUC .CS',
  'Civil Tool\100.ThayDoimauLayer',
  'Civil Tool\15.BeMatTruBeMat',
  'Mô Hình\Mô Hình',
  'Mô Hình\THONG_KE_LENH_COMMAND_METHOD.md',
  # Log build 1.3 MB
  'build_log_uipro.txt'
)
$n = 0
foreach ($f in $xoa) {
  $p = Join-Path $Proj $f
  if (Test-Path -LiteralPath $p) { Remove-Item -LiteralPath $p -Force; Ok "Xoa: $f"; $n++ }
  else { Warn "Khong thay (bo qua): $f" }
}
$bk = Join-Path $Root '_Backup_UIPro'
if (Test-Path $bk) { Remove-Item $bk -Recurse -Force; Ok 'Xoa: _Backup_UIPro\ (da nam trong commit moc an toan)'; $n++ }
Write-Host "  => Da xoa $n muc" -ForegroundColor Green

# ---------------------------------------------------------------------
Step '4/5  Build kiem tra'
Push-Location $Proj
dotnet build MyFirstProject.csproj -c Debug -nologo -v quiet -clp:ErrorsOnly
$code = $LASTEXITCODE
Pop-Location
if ($code -ne 0) {
  Warn 'BUILD LOI - chua commit buoc don dep.'
  Warn 'Muon hoan tac:  git checkout -- .   (hoac: git reset --hard truoc-don-dep)'
  Read-Host 'Bam Enter de thoat'; exit 1
}
Ok 'Build thanh cong'

# ---------------------------------------------------------------------
Step '5/5  Commit don dep'
git add -A
git -c core.quotepath=false commit -q -m "chore: giai doan 0 - xoa 5 lenh vo rong, file rong, TestForm/AdskGreeting; chuyen .gitignore ra goc"
if ($LASTEXITCODE -eq 0) { Ok 'Da commit' } else { Warn 'Khong co thay doi de commit' }

Write-Host "`nHOAN TAT. Kiem tra lai bang:  git log --oneline -3" -ForegroundColor Green
Write-Host "Day len GitHub khi da kiem tra xong:  git push origin master --tags" -ForegroundColor Green
git log --oneline -3
Read-Host "`nBam Enter de dong"
