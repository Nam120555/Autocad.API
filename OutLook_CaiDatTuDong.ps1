# =====================================================================
#  OutLook_CaiDatTuDong.ps1
#  1) Build + cai OutLookTool moi vao C:\OutLook
#  2) Dat lich tu xuat Excel hang ngay (07:45 va 16:45)
#  3) Xuat ngay 1 lan, tao loi tat tren Desktop, mo file Excel
#  Ket qua ghi vao _ci_ketqua.txt (Claude doc lai)
# =====================================================================
$ErrorActionPreference = 'Continue'
$Root   = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root
$Result = Join-Path $Root '_ci_ketqua.txt'
$Log    = Join-Path $Root '_build_check.log'
$App    = 'C:\OutLook'
$Exe    = Join-Path $App 'OutLookTool.exe'
$Latest = Join-Path $App 'ThongKe_Email_MoiNhat.xlsx'
function Res($t) { Add-Content -Path $Result -Value $t -Encoding UTF8; Write-Host $t }
function Popup($text, $icon) { try { Add-Type -AssemblyName System.Windows.Forms; [void][System.Windows.Forms.MessageBox]::Show($text, 'OutLookTool - Cai dat tu dong', 'OK', $icon) } catch { } }
Set-Content -Path $Result -Value "START OUTLOOK $(Get-Date -Format s)" -Encoding UTF8

# 1. Build + publish
if (Get-Process OutLookTool -ErrorAction SilentlyContinue) { Stop-Process -Name OutLookTool -Force; Start-Sleep 1 }
New-Item -ItemType Directory -Force -Path $App | Out-Null
dotnet publish (Join-Path $Root 'OutLookTool\OutLookTool.csproj') -c Release -r win-x64 --self-contained false -p:PublishSingleFile=true -o $App -nologo -v minimal *> $Log
if ($LASTEXITCODE -ne 0) {
    Res 'BUILD: FAIL'
    Get-Content $Log | Select-String ' error ' | Select-Object -First 30 | ForEach-Object { Res $_.Line }
    Popup "BUILD LOI - chua cai dat gi.`nNhan cho Claude: 'outlook bao loi'." 'Error'
    exit 1
}
Res "BUILD: OK -> $Exe"

# 2. Lich tu dong
$tr = "`"$Exe`" --export"
schtasks /Create /TN 'NXSOFT\OutLookTool Xuat Excel Sang' /TR $tr /SC DAILY /ST 07:45 /F | Out-Null
$a = $LASTEXITCODE
schtasks /Create /TN 'NXSOFT\OutLookTool Xuat Excel Chieu' /TR $tr /SC DAILY /ST 16:45 /F | Out-Null
$b = $LASTEXITCODE
if ($a -eq 0 -and $b -eq 0) { Res 'SCHEDULE: OK 07:45 + 16:45 hang ngay' } else { Res "SCHEDULE: FAIL ($a/$b)" }

# 3. Xuat ngay
Res 'EXPORT: dang chay...'
$p = Start-Process -FilePath $Exe -ArgumentList '--export' -PassThru -Wait
Res "EXPORT: exit=$($p.ExitCode)"
$tl = Join-Path $App 'TuDong_log.txt'
if (Test-Path $tl) { Get-Content $tl -Encoding UTF8 | Select-Object -Last 8 | ForEach-Object { Res "  $_" } }

# Loi tat Desktop
try {
    $desk = [Environment]::GetFolderPath('Desktop')
    $ws = New-Object -ComObject WScript.Shell
    $s1 = $ws.CreateShortcut((Join-Path $desk 'Email - Thong ke moi nhat.lnk')); $s1.TargetPath = $Latest; $s1.Save()
    $s2 = $ws.CreateShortcut((Join-Path $desk 'Phan mem loc Email Outlook.lnk')); $s2.TargetPath = $Exe; $s2.WorkingDirectory = $App; $s2.Save()
    Res 'SHORTCUT: OK'
} catch { Res "SHORTCUT: FAIL $_" }

# 4. Commit code
git add -A OutLookTool OutLook_CaiDatTuDong.ps1 OutLook_CaiDatTuDong.bat OutLook_GoLichTuDong.bat 2>$null
git -c core.quotepath=false commit -q -m "feat(OutLookTool): xuat Excel tu dong hang ngay (--export), loc + tim kiem tren sheet Chi tiet, ngay dang date, mo email trong Outlook" 2>$null
if ($LASTEXITCODE -eq 0) { Res "COMMIT: OK $(git log --oneline -1)" } else { Res 'COMMIT: nothing' }
Res "END $(Get-Date -Format s)"

if ($p.ExitCode -eq 0 -and (Test-Path $Latest)) {
    Popup "XONG!`n`n- File tra cuu: $Latest`n- Tu cap nhat 07:45 va 16:45 moi ngay`n- Loi tat da tao tren Desktop" 'Information'
    Start-Process $Latest
} else {
    Popup "Da cai lich, nhung lan xuat dau tien chua thanh cong.`nXem: $tl`nNhan cho Claude: 'outlook bao loi'." 'Warning'
}
