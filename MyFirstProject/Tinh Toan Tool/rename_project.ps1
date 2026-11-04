# =============================================================
# Script đổi tên dự án: MyFirstProject → Project_Nam
# Chạy từ thư mục: D:\Autocad 2026_API\Autocad 2026_API\
# =============================================================

$Root       = "D:\Autocad 2026_API\Autocad 2026_API"
$OldName    = "MyFirstProject"
$NewName    = "Project_Nam"
$OldFolder  = Join-Path $Root $OldName
$NewFolder  = Join-Path $Root $NewName

Write-Host "========================================" -ForegroundColor Cyan
Write-Host " RENAME PROJECT: $OldName → $NewName" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# ── BƯỚC 1: Thay thế nội dung trong tất cả .cs, .csproj, .sln ──
Write-Host "`n[1/5] Thay thế nội dung text trong các file..." -ForegroundColor Yellow

$extensions = @("*.cs", "*.csproj", "*.sln", "*.md", "*.txt")
$count = 0

foreach ($ext in $extensions) {
    $files = Get-ChildItem -Path $Root -Recurse -Filter $ext -File |
             Where-Object { $_.FullName -notmatch "\\obj\\" -and $_.FullName -notmatch "\\bin\\" }

    foreach ($file in $files) {
        $content = Get-Content $file.FullName -Raw -Encoding UTF8
        if ($content -match [regex]::Escape($OldName)) {
            $newContent = $content -replace [regex]::Escape($OldName), $NewName
            [System.IO.File]::WriteAllText($file.FullName, $newContent, [System.Text.Encoding]::UTF8)
            Write-Host "  ✓ $($file.Name)" -ForegroundColor Green
            $count++
        }
    }
}
Write-Host "  → Đã cập nhật $count file" -ForegroundColor Green

# ── BƯỚC 2: Đổi tên file .csproj bên trong folder ──
Write-Host "`n[2/5] Đổi tên file .csproj..." -ForegroundColor Yellow

$oldCsproj = Join-Path $OldFolder "$OldName.csproj"
$newCsproj = Join-Path $OldFolder "$NewName.csproj"

if (Test-Path $oldCsproj) {
    Rename-Item -Path $oldCsproj -NewName "$NewName.csproj"
    Write-Host "  ✓ $OldName.csproj → $NewName.csproj" -ForegroundColor Green
} else {
    Write-Host "  ⚠ Không tìm thấy $oldCsproj" -ForegroundColor Red
}

# ── BƯỚC 3: Đổi tên folder chính ──
Write-Host "`n[3/5] Đổi tên thư mục dự án..." -ForegroundColor Yellow

if (Test-Path $OldFolder) {
    Rename-Item -Path $OldFolder -NewName $NewName
    Write-Host "  ✓ $OldName\ → $NewName\" -ForegroundColor Green
} else {
    Write-Host "  ⚠ Thư mục không tồn tại: $OldFolder" -ForegroundColor Red
}

# ── BƯỚC 4: Đổi tên file .sln ──
Write-Host "`n[4/5] Đổi tên file .sln..." -ForegroundColor Yellow

$oldSln = Join-Path $Root "$OldName.sln"
$newSln  = Join-Path $Root "$NewName.sln"

if (Test-Path $oldSln) {
    Rename-Item -Path $oldSln -NewName "$NewName.sln"
    Write-Host "  ✓ $OldName.sln → $NewName.sln" -ForegroundColor Green
} else {
    Write-Host "  ⚠ Không tìm thấy $oldSln" -ForegroundColor Red
}

# ── BƯỚC 5: Cập nhật lại file .sln với path mới ──
Write-Host "`n[5/5] Cập nhật đường dẫn trong .sln..." -ForegroundColor Yellow

$slnPath = Join-Path $Root "$NewName.sln"
if (Test-Path $slnPath) {
    $slnContent = Get-Content $slnPath -Raw -Encoding UTF8
    # Path trong .sln đã được thay ở bước 1 (MyFirstProject\ → Project_Nam\)
    Write-Host "  ✓ Path trong .sln đã được cập nhật" -ForegroundColor Green
}

# ── HOÀN THÀNH ──
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host " ✅ XONG! Dự án đã được đổi tên." -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host " Tiếp theo:" -ForegroundColor White
Write-Host "  1. Mở Visual Studio → mở file: $NewName.sln" -ForegroundColor White
Write-Host "  2. Build lại dự án (Ctrl+Shift+B)" -ForegroundColor White
Write-Host "  3. File DLL mới: $NewName.dll" -ForegroundColor White
Write-Host "  4. Trong AutoCAD dùng NETLOAD để load lại file DLL mới" -ForegroundColor White
Write-Host ""
