# Live Update Script: Refreshes Outlook data directly into the OPEN Excel workbook
# Path: C:\OutLook\update_data_live.ps1
# Called by VBA macro button inside the .xlsm file

param(
    [string]$TargetExcel = "C:\OutLook\ThongKe_Email_QLTKDS01_V3_Linked.xlsm",
    [string]$pstPath = "C:\Users\NAM\AppData\Local\Microsoft\Outlook\sgc.qltkds01@sgc.net.vn.ost"
)

$msgFolder = "C:\OutLook\SavedEmails_Archive"

if (-not (Test-Path $msgFolder)) { New-Item -ItemType Directory -Path $msgFolder | Out-Null }

Write-Host "============================================="
Write-Host "  CAP NHAT TRUC TIEP VAO EXCEL DANG MO"
Write-Host "============================================="
Write-Host ""

# ===== HELPER FUNCTIONS =====

function Get-OutlookInstance {
    for ($attempt = 1; $attempt -le 10; $attempt++) {
        try {
            return [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
        } catch {}
        try {
            return New-Object -ComObject Outlook.Application
        } catch {}
        Start-Sleep -Seconds 2
    }
    throw "Could not connect to Outlook"
}

function Get-UnitCategory([string]$senderName, [string]$senderEmail) {
    $text = "$senderName $senderEmail"
    if ($text -match "SGC_KKH-KT|SGC_KKH|SGC_PKT|PKT") { return "SGC - Khoi Ke hoach Ky thuat" }
    if ($text -match "SGC_KĐSTĐC|SGC_KDSTDC|SGC_KĐSCT|SGC_DSCT") { return "SGC - Khoi Duong sat Toc do cao" }
    if ($text -match "SGC_KQLDA|SGC_QLDA") { return "SGC - Khoi Quan ly Du an" }
    if ($text -match "SGC_KQLXD|SGC_CTCT|BQLXD|Ban QLXD") { return "SGC - Ban QLXD & Cong trinh" }
    if ($text -match "SGC_KTCH|SGC_KTCXD|SGC_KTCCKN") { return "SGC - Khoi Than toc / Khoi Ky thuat" }
    if ($text -match "SGC_KQLKT|SGC_KHT|SGC_VP|SGC_Phòng") { return "SGC - Khoi Van phong & Hanh chinh" }
    if ($text -match "VH-VTK|VH-KCUDT|VH-KXD|VH-PTDA|Vinhomes|vinhomes\.vn") { return "Vinhomes - Vien Thiet ke & Phat trien DA" }
    if ($text -match "VinCons|vincons\.net") { return "Nha thau VinCons" }
    if ($text -match "SYSTRA|systra\.com") { return "Tu van SYSTRA" }
    if ($text -match "CONINCO|ckjvn\.vn|coninco") { return "Tu van Giam sat CONINCO" }
    if ($text -match "Trung Chính|trungchinhtcc") { return "Nha thau Trung Chinh" }
    if ($text -match "Đại Dũng|daidung\.vn") { return "Nha thau Dai Dung" }
    if ($text -match "Tecco1|tecco1\.vn") { return "Nha thau Tecco 1" }
    if ($text -match "Thiên Quốc An|TQA") { return "Nha thau Thien Quoc An" }
    if ($text -match "620idic") { return "Nha thau 620 IDIC" }
    if ($text -match "SGC") { return "SGC - Don vi khac" }
    return "Doi tac / Don vi ngoai"
}

function Get-TopicCategory([string]$subject) {
    if (-not $subject) { return "Chua xac dinh" }
    if ($subject -match "Bãi đúc|đầm|dầm|đốm|gói dầm") { return "Bai duc & Dam cau" }
    if ($subject -match "cọc|ép cọc|khoan nhồi|Bê tông|be tong|TRIAL MIX|C25|C30") { return "Thi cong Coc & Be tong" }
    if ($subject -match "Thiết kế|TK|Shop|SHOP|VMTK|MEP|Track|Design|bản vẽ|BIM") { return "Thiet ke & Shopdrawing" }
    if ($subject -match "cao độ|đường đỏ|ranh|rừng|mặt bằng|mốc|trắc địa") { return "Trac dia, Cao do & Mat bang" }
    if ($subject -match "Tiến độ|báo cáo|vướng mắc|ảnh hưởng|Đề xuất|de xuat") { return "Tien do & Vuong mac Cong truong" }
    if ($subject -match "Văn bản|Công văn|NNMT|phạt|pháp lý|Ký|phê duyệt|Quyết định") { return "Phap ly & Cong van Hanh chinh" }
    if ($subject -match "máy phát|điện|hàng rào|cầu rửa xe|thiết bị|vật tư|cáp") { return "Ha tang Bai duc & Thiet bi" }
    return "Cong viec / Trao doi khac"
}

# ===== STEP 1: EXTRACT EMAILS FROM OUTLOOK =====

try {
    Write-Host "Ket noi Outlook..."
    $outlook = Get-OutlookInstance
    $namespace = $outlook.GetNamespace("MAPI")
    
    $pstStore = $null
    foreach ($store in $namespace.Stores) {
        try {
            if ($store.FilePath -eq $pstPath -or $store.FilePath -like "*sgc.qltkds01*" -or $store.DisplayName -like "*sgc.qltkds01*" -or $store.FilePath -like "*QLTKDS01*") {
                $pstStore = $store; break
            }
        } catch {}
    }
    
    if (-not $pstStore) {
        Write-Host "Ket noi PST/OST file..."
        try {
            $namespace.AddStore($pstPath)
            Start-Sleep -Seconds 3
        } catch {}
        foreach ($store in $namespace.Stores) {
            try {
                if ($store.FilePath -eq $pstPath -or $store.FilePath -like "*sgc.qltkds01*" -or $store.DisplayName -like "*sgc.qltkds01*" -or $store.FilePath -like "*QLTKDS01*") { $pstStore = $store; break }
            } catch {}
        }
    }
    
    if (-not $pstStore) {
        # Fallback to default store if available
        try {
            $pstStore = $namespace.DefaultStore
        } catch {}
    }
    
    if (-not $pstStore) { throw "Cannot access Outlook/PST/OST store" }
    $rootFolder = $pstStore.GetRootFolder()
    
    Write-Host "Dang quet email..."
    $allEmails = [System.Collections.ArrayList]::new()
    
    function Extract-Emails {
        param([object]$folder, [string]$path)
        $currentPath = if ($path) { "$path\$($folder.Name)" } else { $folder.Name }
        $itemCount = 0
        try { $itemCount = $folder.Items.Count } catch {}
        
        if ($itemCount -gt 0) {
            $items = $folder.Items
            for ($i = 1; $i -le $itemCount; $i++) {
                try {
                    $item = $items.Item($i)
                    if ($item.Class -eq 43) {
                        $sn = ""; $se = ""; $subj = ""; $rt = $null; $so = $null
                        $imp = "Binh thuong"; $ha = $false; $sz = 0; $entryID = ""
                        
                        try { $sn = $item.SenderName } catch {}
                        try { $se = $item.SenderEmailAddress } catch {}
                        try { $entryID = $item.EntryID } catch {}
                        
                        $set = ""
                        try { $set = $item.SenderEmailType } catch {}
                        if ($set -eq "EX" -or $se -like "/*") {
                            $smtpResolved = ""
                            try { $smtpResolved = $item.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x5D01001E") } catch {}
                            if (-not $smtpResolved) {
                                try {
                                    $exUser = $item.Sender.GetExchangeUser()
                                    if ($exUser) { $smtpResolved = $exUser.PrimarySmtpAddress }
                                } catch {}
                            }
                            if ($smtpResolved) { $se = $smtpResolved }
                        }
                        
                        try { $subj = $item.Subject } catch {}
                        try { $rt = $item.ReceivedTime } catch {}
                        try { $so = $item.SentOn } catch {}
                        try { $impV = $item.Importance; $imp = switch ($impV) { 0 {"Thap"} 1 {"Binh thuong"} 2 {"Cao (High)"} default {"Binh thuong"} } } catch {}
                        try { $ha = ($item.Attachments.Count -gt 0) } catch {}
                        try { $sz = $item.Size } catch {}
                        
                        $tl = [System.Collections.ArrayList]::new()
                        $cl = [System.Collections.ArrayList]::new()
                        try {
                            $recips = $item.Recipients
                            for ($rx = 1; $rx -le $recips.Count; $rx++) {
                                try {
                                    $rcp = $recips.Item($rx)
                                    if ($rcp.Type -eq 1) { $null = $tl.Add($rcp.Name) }
                                    elseif ($rcp.Type -eq 2) { $null = $cl.Add($rcp.Name) }
                                } catch {}
                            }
                        } catch {}
                        
                        $dobj = $null; $ds = ""; $ts = ""; $weekStr = ""
                        if ($rt) {
                            $dobj = $rt; $ds = $rt.ToString("yyyy-MM-dd"); $ts = $rt.ToString("HH:mm:ss")
                            $culture = [System.Globalization.CultureInfo]::InvariantCulture
                            $weekNum = $culture.Calendar.GetWeekOfYear($rt, [System.Globalization.CalendarWeekRule]::FirstFourDayWeek, [DayOfWeek]::Monday)
                            $weekStr = "Tuan " + $weekNum
                        } elseif ($so) {
                            $dobj = $so; $ds = $so.ToString("yyyy-MM-dd"); $ts = $so.ToString("HH:mm:ss")
                        }
                        
                        # Save MSG File
                        $safeSubj = $subj -replace '[\\/:*?"<>|\r\n]', '_'
                        if ($safeSubj.Length -gt 60) { $safeSubj = $safeSubj.Substring(0,60) }
                        $msgFileName = ("{0:D4}_{1}_{2}.msg" -f ($script:allEmails.Count + 1), ($ds -replace '-',''), $safeSubj)
                        $msgPath = Join-Path $script:msgFolder $msgFileName
                        
                        if (-not (Test-Path $msgPath)) {
                            try { $item.SaveAs($msgPath, 3) } catch {}
                        }
                        
                        $unitCategory = Get-UnitCategory $sn $se
                        $topicCategory = Get-TopicCategory $subj
                        
                        $null = $script:allEmails.Add([PSCustomObject]@{
                            STT=0; Ngay=$ds; Gio=$ts; Tuan=$weekStr; DateObj=$dobj
                            NguoiGui=$sn; EmailNguoiGui=$se; DonVi=$unitCategory
                            NguoiNhan_To=($tl -join "; "); NguoiNhan_CC=($cl -join "; ")
                            ChuDe=$subj; HangMuc=$topicCategory; MucDo=$imp
                            DinhKem=$(if($ha){"Co"}else{"Khong"})
                            DungLuong=[math]::Round($sz/1024, 1)
                            ThuMuc=$currentPath
                            MsgPath=$msgPath
                            EntryID=$entryID
                        })
                    }
                    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($item) | Out-Null } catch {}
                } catch {}
            }
        }
        try { foreach ($sf in $folder.Folders) { Extract-Emails -folder $sf -path $currentPath } } catch {}
    }
    
    Extract-Emails -folder $rootFolder -path ""
    Write-Host "Tong so email quet duoc: $($allEmails.Count)"
    
    $allEmails = [System.Collections.ArrayList]::new(@($allEmails | Sort-Object { $_.DateObj } -Descending))
    for ($i = 0; $i -lt $allEmails.Count; $i++) { $allEmails[$i].STT = $i + 1 }

    # ===== STEP 2: CONNECT TO OPEN EXCEL =====
    
    Write-Host "Ket noi vao Excel dang mo..."
    $excel = $null
    try {
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    } catch {
        Write-Host "Khong tim thay Excel dang mo. Mo file moi..."
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true
    }
    $excel.DisplayAlerts = $false
    
    # Find or open the target workbook
    $wb = $null
    foreach ($wbItem in $excel.Workbooks) {
        $wbName = ""
        try { $wbName = $wbItem.FullName } catch {}
        if ($wbName -eq $TargetExcel -or $wbName -like "*ThongKe_Email_QLTKDS01_V3*") {
            $wb = $wbItem
            break
        }
    }
    
    if (-not $wb) {
        if (Test-Path $TargetExcel) {
            $wb = $excel.Workbooks.Open($TargetExcel)
        } else {
            Write-Host "[LOI] Khong tim thay file $TargetExcel"
            Write-Host "Hay chay sync_emails.ps1 truoc de tao file ban dau."
            exit 1
        }
    }
    
    Write-Host "Da ket noi workbook: $($wb.Name)"
    
    # ===== STEP 3: UPDATE DATA IN SHEETS =====
    
    # Theme Colors
    $cNavyH  = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(24, 43, 73))
    $cBlueH  = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(0, 112, 192))
    $cGreenH = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(38, 128, 0))
    $cRedH   = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(192, 0, 0))
    $cPurpH  = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(112, 48, 160))
    $cDkBlH  = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(41, 86, 129))
    $cW      = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::White)
    $cCardBg = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(240, 244, 248))
    $cLtBl   = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(235, 243, 250))
    $cLtGr   = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(235, 247, 235))
    $cLtRd   = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(253, 237, 237))
    $cLtPr   = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(245, 238, 248))
    $cYel    = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(255, 242, 204))
    $cWE     = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(255, 220, 220))
    $cHR     = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(255, 120, 120))
    $cHO     = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(255, 200, 100))
    $cHG     = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(160, 225, 120))
    
    # --- Helper to get sheet by name ---
    function Get-Sheet([string]$name) {
        foreach ($ws in $wb.Worksheets) {
            if ($ws.Name -eq $name) { return $ws }
        }
        return $null
    }
    
    # ========== UPDATE DASHBOARD ==========
    Write-Host "Cap nhat Dashboard..."
    $wsDash = Get-Sheet "Trang Chu - Dashboard"
    if ($wsDash) {
        # Update sync time
        $wsDash.Range("A2").Value = "Dong bo luc: " + (Get-Date -Format "dd/MM/yyyy HH:mm:ss") + " | Thu muc Mail: C:\OutLook\SavedEmails_Archive"
        
        # Update KPI values (row 6 = KPI value cells)
        $wsDash.Range("A6").Value = [string]$allEmails.Count
        $wsDash.Range("C6").Value = [string]($allEmails | Select-Object -ExpandProperty NguoiGui -Unique).Count
        $highCount = @($allEmails | Where-Object { $_.MucDo -like "*Cao*" }).Count
        $wsDash.Range("E6").Value = [string]$highCount
        $attachCount = @($allEmails | Where-Object { $_.DinhKem -eq "Co" }).Count
        $wsDash.Range("F6").Value = [string]$attachCount
        
        # Update Unit statistics (starting row 11)
        $unitGroups = $allEmails | Group-Object DonVi | Sort-Object Count -Descending
        # Clear old unit data (rows 11 to 60)
        $clearRange = $wsDash.Range("A11:D60")
        $clearRange.ClearContents()
        $clearRange.Interior.ColorIndex = 0
        
        $row = 11; $uStt = 1
        foreach ($ug in $unitGroups) {
            $pct = [math]::Round(($ug.Count / $allEmails.Count) * 100, 1)
            $wsDash.Cells.Item($row, 1) = [int]$uStt
            $wsDash.Cells.Item($row, 2) = [string]$ug.Name
            $wsDash.Cells.Item($row, 3) = [int]$ug.Count
            $wsDash.Cells.Item($row, 4) = [string]"$pct %"
            if ($row % 2 -eq 0) { $wsDash.Range("A${row}:D${row}").Interior.Color = $cLtBl }
            $row++; $uStt++
        }
        $wsDash.Range("A10:D" + ($row-1)).Borders.LineStyle = 1
        
        # Update Topic statistics (starting row 11, columns E-H)
        $topicGroups = $allEmails | Group-Object HangMuc | Sort-Object Count -Descending
        $clearRange2 = $wsDash.Range("E11:H60")
        $clearRange2.ClearContents()
        $clearRange2.Interior.ColorIndex = 0
        
        $tRow = 11; $tStt = 1
        foreach ($tg in $topicGroups) {
            $pct = [math]::Round(($tg.Count / $allEmails.Count) * 100, 1)
            $wsDash.Cells.Item($tRow, 5) = [int]$tStt
            $wsDash.Cells.Item($tRow, 6) = [string]$tg.Name
            $wsDash.Cells.Item($tRow, 7) = [int]$tg.Count
            $wsDash.Cells.Item($tRow, 8) = [string]"$pct %"
            if ($tRow % 2 -eq 0) { $wsDash.Range("E${tRow}:H${tRow}").Interior.Color = $cLtGr }
            $tRow++; $tStt++
        }
        $wsDash.Range("E10:H" + ($tRow-1)).Borders.LineStyle = 1
    }
    
    # ========== UPDATE SHEET: Chi tiet Email ==========
    Write-Host "Cap nhat sheet Chi tiet Email..."
    $ws1 = Get-Sheet "Chi tiet Email"
    if ($ws1) {
        # Clear old data (keep header row 1)
        $lastUsedRow = $ws1.UsedRange.Rows.Count + $ws1.UsedRange.Row - 1
        if ($lastUsedRow -gt 1) {
            $ws1.Range($ws1.Cells.Item(2, 1), $ws1.Cells.Item($lastUsedRow + 10, 16)).ClearContents()
            $ws1.Range($ws1.Cells.Item(2, 1), $ws1.Cells.Item($lastUsedRow + 10, 16)).Interior.ColorIndex = 0
        }
        
        $row = 2
        foreach ($em in $allEmails) {
            $ws1.Cells.Item($row, 1) = [int]$em.STT
            $ws1.Cells.Item($row, 2) = [string]$em.Ngay
            $ws1.Cells.Item($row, 3) = [string]$em.Gio
            $ws1.Cells.Item($row, 4) = [string]$em.Tuan
            $ws1.Cells.Item($row, 5) = [string]$em.NguoiGui
            $ws1.Cells.Item($row, 6) = [string]$em.EmailNguoiGui
            $ws1.Cells.Item($row, 7) = [string]$em.DonVi
            $ws1.Cells.Item($row, 8) = [string]$em.NguoiNhan_To
            $ws1.Cells.Item($row, 9) = [string]$em.NguoiNhan_CC
            $ws1.Cells.Item($row, 10) = [string]$em.ChuDe
            $ws1.Cells.Item($row, 11) = [string]$em.HangMuc
            $ws1.Cells.Item($row, 12) = [string]$em.MucDo
            $ws1.Cells.Item($row, 13) = [string]$em.DinhKem
            $ws1.Cells.Item($row, 14) = [double]$em.DungLuong
            
            if (Test-Path $em.MsgPath) {
                $ws1.Hyperlinks.Add($ws1.Cells.Item($row, 15), $em.MsgPath, "", "Bam de mo file email .msg", "Mo Mail (.msg)") | Out-Null
            } else {
                $ws1.Cells.Item($row, 15) = [string]$em.ThuMuc
            }
            
            if ($em.EntryID) {
                $outlookUri = "outlook:" + $em.EntryID
                $ws1.Hyperlinks.Add($ws1.Cells.Item($row, 16), $outlookUri, "", "Bam de mo truc tiep trong Outlook", "Mo trong Outlook") | Out-Null
            } else {
                $ws1.Cells.Item($row, 16) = "-"
            }
            
            if ($row % 2 -eq 0) { $ws1.Range($ws1.Cells.Item($row,1),$ws1.Cells.Item($row,16)).Interior.Color = $cLtBl }
            if ($em.MucDo -like "*Cao*") {
                $ws1.Cells.Item($row, 12).Font.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::Red)
                $ws1.Cells.Item($row, 12).Font.Bold = $true
            }
            $row++
        }
        # Re-apply borders
        $ws1.Range($ws1.Cells.Item(1,1), $ws1.Cells.Item($row-1, 16)).Borders.LineStyle = 1
        $ws1.Range($ws1.Cells.Item(1,1), $ws1.Cells.Item($row-1, 16)).Borders.Weight = 2
    }
    
    # ========== UPDATE SHEET: Thong ke theo Ngay ==========
    Write-Host "Cap nhat sheet Thong ke theo Ngay..."
    $ws2 = Get-Sheet "Thong ke theo Ngay"
    if ($ws2) {
        $lastUsedRow = $ws2.UsedRange.Rows.Count + $ws2.UsedRange.Row - 1
        if ($lastUsedRow -gt 1) {
            $ws2.Range($ws2.Cells.Item(2, 1), $ws2.Cells.Item($lastUsedRow + 10, 9)).ClearContents()
            $ws2.Range($ws2.Cells.Item(2, 1), $ws2.Cells.Item($lastUsedRow + 10, 9)).Interior.ColorIndex = 0
        }
        
        $dateGroups = $allEmails | Where-Object { $_.Ngay } | Group-Object Ngay | Sort-Object Name -Descending
        $row = 2; $stt = 1
        foreach ($dg in $dateGroups) {
            $dd = [datetime]::Parse($dg.Name)
            $dow = switch ($dd.DayOfWeek) { "Monday"{"Thu 2"} "Tuesday"{"Thu 3"} "Wednesday"{"Thu 4"} "Thursday"{"Thu 5"} "Friday"{"Thu 6"} "Saturday"{"Thu 7"} "Sunday"{"Chu Nhat"} }
            $weekStr = ($dg.Group | Select-Object -First 1).Tuan
            $us = ($dg.Group | Select-Object -ExpandProperty NguoiGui -Unique).Count
            $ar = @(); foreach ($e in $dg.Group) { if ($e.NguoiNhan_To) { $ar += $e.NguoiNhan_To.Split(";") | ForEach-Object { $_.Trim() } | Where-Object { $_ } } }
            $ur = ($ar | Select-Object -Unique).Count
            $hc = @($dg.Group | Where-Object { $_.MucDo -like "*Cao*" }).Count
            $ac = @($dg.Group | Where-Object { $_.DinhKem -eq "Co" }).Count
            $ws2.Cells.Item($row,1) = $stt; $ws2.Cells.Item($row,2) = [string]$dg.Name; $ws2.Cells.Item($row,3) = [string]$dow
            $ws2.Cells.Item($row,4) = [string]$weekStr; $ws2.Cells.Item($row,5) = [int]$dg.Count; $ws2.Cells.Item($row,6) = [int]$us
            $ws2.Cells.Item($row,7) = [int]$ur; $ws2.Cells.Item($row,8) = [int]$hc; $ws2.Cells.Item($row,9) = [int]$ac
            if ($dow -eq "Thu 7" -or $dow -eq "Chu Nhat") { $ws2.Range($ws2.Cells.Item($row,1),$ws2.Cells.Item($row,9)).Interior.Color = $cWE }
            elseif ($row % 2 -eq 0) { $ws2.Range($ws2.Cells.Item($row,1),$ws2.Cells.Item($row,9)).Interior.Color = $cLtGr }
            $row++; $stt++
        }
        $ws2.Cells.Item($row,2) = "TONG CONG"; $ws2.Cells.Item($row,5) = [int]$allEmails.Count
        $ws2.Range($ws2.Cells.Item($row,1),$ws2.Cells.Item($row,9)).Font.Bold = $true
        $ws2.Range($ws2.Cells.Item($row,1),$ws2.Cells.Item($row,9)).Interior.Color = $cYel
        $ws2.Range($ws2.Cells.Item(1,1), $ws2.Cells.Item($row, 9)).Borders.LineStyle = 1
        $ws2.Range($ws2.Cells.Item(1,1), $ws2.Cells.Item($row, 9)).Borders.Weight = 2
    }
    
    # ========== UPDATE SHEET: Thong ke Nguoi gui ==========
    Write-Host "Cap nhat sheet Thong ke Nguoi gui..."
    $ws3 = Get-Sheet "Thong ke Nguoi gui"
    if ($ws3) {
        $lastUsedRow = $ws3.UsedRange.Rows.Count + $ws3.UsedRange.Row - 1
        if ($lastUsedRow -gt 1) {
            $ws3.Range($ws3.Cells.Item(2, 1), $ws3.Cells.Item($lastUsedRow + 10, 8)).ClearContents()
            $ws3.Range($ws3.Cells.Item(2, 1), $ws3.Cells.Item($lastUsedRow + 10, 8)).Interior.ColorIndex = 0
        }
        
        $senderGroups = $allEmails | Where-Object { $_.NguoiGui } | Group-Object NguoiGui | Sort-Object Count -Descending
        $row = 2; $stt = 1
        foreach ($sg in $senderGroups) {
            $ems = @($sg.Group | Where-Object { $_.DateObj })
            $f = ($ems | Sort-Object DateObj | Select-Object -First 1).Ngay
            $l = ($ems | Sort-Object DateObj -Descending | Select-Object -First 1).Ngay
            $d = ($ems | Select-Object -ExpandProperty Ngay -Unique).Count
            $sObj = $sg.Group | Select-Object -First 1
            $ws3.Cells.Item($row,1) = $stt; $ws3.Cells.Item($row,2) = [string]$sg.Name; $ws3.Cells.Item($row,3) = [string]$sObj.EmailNguoiGui
            $ws3.Cells.Item($row,4) = [string]$sObj.DonVi; $ws3.Cells.Item($row,5) = [int]$sg.Count; $ws3.Cells.Item($row,6) = [string]$f
            $ws3.Cells.Item($row,7) = [string]$l; $ws3.Cells.Item($row,8) = [int]$d
            if ($row % 2 -eq 0) { $ws3.Range($ws3.Cells.Item($row,1),$ws3.Cells.Item($row,8)).Interior.Color = $cLtRd }
            $row++; $stt++
        }
        $ws3.Range($ws3.Cells.Item(1,1), $ws3.Cells.Item($row-1, 8)).Borders.LineStyle = 1
        $ws3.Range($ws3.Cells.Item(1,1), $ws3.Cells.Item($row-1, 8)).Borders.Weight = 2
    }
    
    # ========== UPDATE SHEET: Thong ke Nguoi nhan ==========
    Write-Host "Cap nhat sheet Thong ke Nguoi nhan..."
    $ws4 = Get-Sheet "Thong ke Nguoi nhan"
    if ($ws4) {
        $lastUsedRow = $ws4.UsedRange.Rows.Count + $ws4.UsedRange.Row - 1
        if ($lastUsedRow -gt 1) {
            $ws4.Range($ws4.Cells.Item(2, 1), $ws4.Cells.Item($lastUsedRow + 10, 6)).ClearContents()
            $ws4.Range($ws4.Cells.Item(2, 1), $ws4.Cells.Item($lastUsedRow + 10, 6)).Interior.ColorIndex = 0
        }
        
        $rMap = @{}
        foreach ($em in $allEmails) {
            $ar = @()
            if ($em.NguoiNhan_To) { $ar += $em.NguoiNhan_To.Split(";") | ForEach-Object { $_.Trim() } | Where-Object { $_ } }
            if ($em.NguoiNhan_CC) { $ar += $em.NguoiNhan_CC.Split(";") | ForEach-Object { $_.Trim() } | Where-Object { $_ } }
            foreach ($rn in $ar) {
                if (-not $rMap.ContainsKey($rn)) { $rMap[$rn] = @{N=$rn;C=0;T=0;X=0;D=[System.Collections.ArrayList]::new()} }
                $rMap[$rn].C++
                if ($em.NguoiNhan_To -and $em.NguoiNhan_To.Contains($rn)) { $rMap[$rn].T++ }
                if ($em.NguoiNhan_CC -and $em.NguoiNhan_CC.Contains($rn)) { $rMap[$rn].X++ }
                if ($em.Ngay -and $rMap[$rn].D -notcontains $em.Ngay) { $null=$rMap[$rn].D.Add($em.Ngay) }
            }
        }
        $rList = $rMap.Values | Sort-Object { $_.C } -Descending | Select-Object -First 500
        $row = 2; $stt = 1
        foreach ($r in $rList) {
            $ws4.Cells.Item($row,1) = $stt; $ws4.Cells.Item($row,2) = [string]$r.N; $ws4.Cells.Item($row,3) = [int]$r.C
            $ws4.Cells.Item($row,4) = [int]$r.T; $ws4.Cells.Item($row,5) = [int]$r.X; $ws4.Cells.Item($row,6) = [int]$r.D.Count
            if ($row % 2 -eq 0) { $ws4.Range($ws4.Cells.Item($row,1),$ws4.Cells.Item($row,6)).Interior.Color = $cLtPr }
            $row++; $stt++
        }
        $ws4.Range($ws4.Cells.Item(1,1), $ws4.Cells.Item($row-1, 6)).Borders.LineStyle = 1
        $ws4.Range($ws4.Cells.Item(1,1), $ws4.Cells.Item($row-1, 6)).Borders.Weight = 2
    }
    
    # ========== UPDATE SHEET: Ma tran Gui x Ngay ==========
    Write-Host "Cap nhat sheet Ma tran..."
    $ws5 = Get-Sheet "Ma tran Gui x Ngay"
    if ($ws5) {
        # Clear everything - we'll rewrite headers too for matrix
        $lastUsedRow = $ws5.UsedRange.Rows.Count + $ws5.UsedRange.Row - 1
        $lastUsedCol = $ws5.UsedRange.Columns.Count + $ws5.UsedRange.Column - 1
        if ($lastUsedCol -lt 4) { $lastUsedCol = 50 }
        $ws5.Range($ws5.Cells.Item(1, 1), $ws5.Cells.Item($lastUsedRow + 10, $lastUsedCol + 10)).ClearContents()
        $ws5.Range($ws5.Cells.Item(1, 1), $ws5.Cells.Item($lastUsedRow + 10, $lastUsedCol + 10)).Interior.ColorIndex = 0
        
        $senderGroups = $allEmails | Where-Object { $_.NguoiGui } | Group-Object NguoiGui | Sort-Object Count -Descending
        $uDates = $allEmails | Where-Object { $_.Ngay } | Select-Object -ExpandProperty Ngay -Unique | Sort-Object
        $top30 = $senderGroups | Select-Object -First 30
        
        $ws5.Cells.Item(1,1) = "STT"; $ws5.Cells.Item(1,2) = "Nguoi gui"; $ws5.Cells.Item(1,3) = "Don vi"
        for ($d = 0; $d -lt $uDates.Count; $d++) { $ws5.Cells.Item(1, $d+4) = [string]$uDates[$d] }
        $tc = $uDates.Count + 4
        $ws5.Cells.Item(1, $tc) = "TONG"
        $hr5 = $ws5.Range($ws5.Cells.Item(1,1),$ws5.Cells.Item(1,$tc))
        $hr5.Font.Bold = $true; $hr5.Font.Size = 10; $hr5.Font.Color = $cW; $hr5.Interior.Color = $cDkBlH; $hr5.HorizontalAlignment = -4108
        
        $row = 2; $stt = 1
        foreach ($s in $top30) {
            $sObj = $s.Group | Select-Object -First 1
            $ws5.Cells.Item($row,1) = $stt; $ws5.Cells.Item($row,2) = [string]$s.Name; $ws5.Cells.Item($row,3) = [string]$sObj.DonVi
            $st = 0
            for ($d = 0; $d -lt $uDates.Count; $d++) {
                $cnt = @($s.Group | Where-Object { $_.Ngay -eq $uDates[$d] }).Count
                if ($cnt -gt 0) {
                    $ws5.Cells.Item($row, $d+4) = [int]$cnt; $st += $cnt
                    if ($cnt -ge 5) { $ws5.Cells.Item($row, $d+4).Interior.Color = $cHR }
                    elseif ($cnt -ge 3) { $ws5.Cells.Item($row, $d+4).Interior.Color = $cHO }
                    else { $ws5.Cells.Item($row, $d+4).Interior.Color = $cHG }
                }
            }
            $ws5.Cells.Item($row, $tc) = [int]$st; $row++; $stt++
        }
        $ws5.Cells.Item($row,2) = "TONG / NGAY"; $ws5.Cells.Item($row,2).Font.Bold = $true
        for ($d = 0; $d -lt $uDates.Count; $d++) {
            $ws5.Cells.Item($row, $d+4) = [int]@($allEmails | Where-Object { $_.Ngay -eq $uDates[$d] }).Count
        }
        $ws5.Cells.Item($row, $tc) = [int]$allEmails.Count
        $ws5.Range($ws5.Cells.Item($row,1),$ws5.Cells.Item($row,$tc)).Font.Bold = $true
        $ws5.Range($ws5.Cells.Item($row,1),$ws5.Cells.Item($row,$tc)).Interior.Color = $cYel
        $ws5.Range($ws5.Cells.Item(1,1), $ws5.Cells.Item($row, $tc)).Borders.LineStyle = 1
        $ws5.Range($ws5.Cells.Item(1,1), $ws5.Cells.Item($row, $tc)).Borders.Weight = 2
        $ws5.Columns.Item(1).ColumnWidth = 5; $ws5.Columns.Item(2).ColumnWidth = 32; $ws5.Columns.Item(3).ColumnWidth = 30
        for ($d = 0; $d -lt $uDates.Count; $d++) { $ws5.Columns.Item($d+4).ColumnWidth = 11 }
        $ws5.Columns.Item($tc).ColumnWidth = 8
    }
    
    # ===== STEP 4: SAVE =====
    Write-Host "Dang luu file..."
    $wb.Save()
    
    # Also save a copy as ALL
    $allExcelPath = "C:\OutLook\ThongKe_Email_QLTKDS01_ALL.xlsx"
    try { $wb.SaveCopyAs($allExcelPath) } catch {}
    
    # Switch to Dashboard
    try { $wsDash.Select() | Out-Null } catch {}
    
    Write-Host ""
    Write-Host "============================================="
    Write-Host "  CAP NHAT THANH CONG!"
    Write-Host "============================================="
    Write-Host "Tong email: $($allEmails.Count)"
    Write-Host "File: $TargetExcel"
    Write-Host "Thoi gian: $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')"
    
} catch {
    Write-Host "LOI: $($_.Exception.Message)"
    Write-Host $_.ScriptStackTrace
}
