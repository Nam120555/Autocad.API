# Day code len GitHub (commit + push) - khong can go lenh
$ErrorActionPreference = 'Continue'
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root
$git = Get-Command git -ErrorAction SilentlyContinue
if (-not $git) {
  Write-Host 'Khong tim thay Git. Tai va cai tai: https://git-scm.com/download/win (bam Next het), roi chay lai file nay.' -ForegroundColor Red
  Read-Host 'Bam Enter de dong'; exit 1
}
Write-Host "Git: $($git.Source)" -ForegroundColor DarkGray
$wf = Join-Path $Root '.github\workflows\build.yml'
if (Test-Path $wf) { Write-Host '[OK] Da co .github\workflows\build.yml' -ForegroundColor Green }
else { Write-Host '[!] Chua co .github\workflows\build.yml - GitHub se KHONG tu build (van day code len binh thuong).' -ForegroundColor Yellow }

git add -A
git commit -q -m "ci: build tu dong + cau hinh Antigravity (AGENTS.md, workflows)" 2>$null
if ($LASTEXITCODE -eq 0) { Write-Host '[OK] Da commit' -ForegroundColor Green } else { Write-Host '[i] Khong co thay doi moi de commit' }

Write-Host "`nDang day len GitHub (lan dau co the hien cua so dang nhap GitHub - hay dang nhap)..." -ForegroundColor Cyan
git push origin master --tags
if ($LASTEXITCODE -eq 0) {
  Write-Host "`n[OK] Da day len GitHub: https://github.com/Nam120555/Autocad.API" -ForegroundColor Green
  Write-Host 'Mo tab Actions tren trang do de xem build tu dong.' -ForegroundColor Green
} else {
  Write-Host "`n[LOI] Push that bai - chup man hinh cua so nay gui cho Claude." -ForegroundColor Red
}
Read-Host "`nBam Enter de dong"
