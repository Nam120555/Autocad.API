# =====================================================================
#  ci.ps1 - Kiem tra build + commit tu dong (dung chung cho cac buoc nang cap)
#  Doc thong diep commit tu file _ci_message.txt (Claude ghi san)
#  Ket qua ghi ra _ci_ketqua.txt de Claude doc lai.
# =====================================================================
$ErrorActionPreference = 'Continue'
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root
$Proj   = Join-Path $Root 'MyFirstProject\MyFirstProject.csproj'
$Log    = Join-Path $Root '_build_check.log'
$Result = Join-Path $Root '_ci_ketqua.txt'
$MsgF   = Join-Path $Root '_ci_message.txt'
function Popup($text, $icon) {
    try { Add-Type -AssemblyName System.Windows.Forms; [void][System.Windows.Forms.MessageBox]::Show($text, 'Civil Tool - Kiem tra build', 'OK', $icon) } catch { }
}
function Res($t) { Add-Content -Path $Result -Value $t -Encoding UTF8; Write-Host $t }
Set-Content -Path $Result -Value "START $(Get-Date -Format s)" -Encoding UTF8

if (-not (Test-Path $MsgF)) { Res 'FAIL: thieu _ci_message.txt'; exit 1 }
$msg = (Get-Content $MsgF -Raw -Encoding UTF8).Trim()

dotnet build $Proj -c Release -nologo -v minimal *> $Log
if ($LASTEXITCODE -ne 0) {
    Res 'BUILD: FAIL'
    Get-Content $Log | Select-String ' error ' | Select-Object -First 30 | ForEach-Object { Res $_.Line }
    Popup "BUILD LOI - chua commit.`nNhan cho Claude: 'ci bao loi' de sua." 'Error'
    exit 1
}
$w = (Get-Content $Log | Select-String 'Warning\(s\)' | Select-Object -Last 1).Line
Res "BUILD: OK $w"

git add -A
git -c core.quotepath=false commit -q -F $MsgF
if ($LASTEXITCODE -eq 0) { Res "COMMIT: OK $(git log --oneline -1)" } else { Res 'COMMIT: nothing' }
Remove-Item $MsgF -Force
Res "END $(Get-Date -Format s)"
Popup "BUILD THANH CONG - da commit.`n$(git log --oneline -1)" 'Information'
