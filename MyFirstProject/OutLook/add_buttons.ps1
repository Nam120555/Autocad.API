# add_buttons.ps1 - Them 3 nut bam va VBA Macro vao file Excel dang mo
# Chay khi file ThongKe_Email_QLTKDS01_V3_Linked.xlsm dang mo trong Excel

Write-Host "============================================="
Write-Host "  THEM NUT BAM VA VBA MACRO VAO EXCEL"
Write-Host "============================================="
Write-Host ""

try {
    # Ket noi vao Excel dang mo
    Write-Host "Ket noi vao Excel dang mo..."
    $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    
    if (-not $excel) {
        throw "Khong tim thay Excel dang mo. Hay mo file Excel truoc."
    }
    
    # Tim workbook
    $wb = $null
    foreach ($wbItem in $excel.Workbooks) {
        $wbName = ""
        try { $wbName = $wbItem.Name } catch {}
        if ($wbName -like "*ThongKe_Email_QLTKDS01*") {
            $wb = $wbItem
            break
        }
    }
    
    if (-not $wb) {
        throw "Khong tim thay workbook ThongKe_Email. Hay mo file Excel truoc."
    }
    
    Write-Host "Da tim thay workbook: $($wb.Name)"
    
    # Tim sheet Dashboard
    $wsDash = $null
    foreach ($ws in $wb.Worksheets) {
        if ($ws.Name -like "*Dashboard*" -or $ws.Name -like "*Trang*") {
            $wsDash = $ws
            break
        }
    }
    
    if (-not $wsDash) {
        $wsDash = $wb.Worksheets.Item(1)
    }
    
    Write-Host "Sheet Dashboard: $($wsDash.Name)"
    
    # === BUOC 1: Chen dong huong dan (Row 3) ===
    Write-Host "Chen dong huong dan..."
    $instrRange = $wsDash.Range("A3:H3")
    $instrRange.Merge()
    $instrRange.Value = "Bam cac nut ben duoi de cap nhat du lieu email truc tiep (khong can dong file Excel)"
    $instrRange.Font.Italic = $true
    $instrRange.Font.Size = 10
    $instrRange.Font.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(24, 43, 73))
    $instrRange.HorizontalAlignment = -4108
    $instrRange.RowHeight = 20
    
    # Tang chieu cao row 4 de chua nut
    $wsDash.Rows.Item(4).RowHeight = 42
    
    # === BUOC 2: Xoa cac Shape cu neu co ===
    Write-Host "Xoa cac nut cu (neu co)..."
    $shapesToRemove = @("btnCapNhatNgay", "btnTuDong30Phut", "btnDungTuDong")
    foreach ($sName in $shapesToRemove) {
        try {
            $wsDash.Shapes.Item($sName).Delete()
        } catch {}
    }
    
    # === BUOC 3: Tao 3 nut Shape moi ===
    Write-Host "Tao 3 nut bam..."
    
    # Nut 1: CAP NHAT NGAY (Xanh duong) - dat o Row 4
    $btn1 = $wsDash.Shapes.AddShape(5, 5, 62, 160, 35)  # msoShapeRoundedRectangle=5
    $btn1.Name = "btnCapNhatNgay"
    $btn1.Fill.ForeColor.RGB = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(0, 112, 192))
    $btn1.Line.Visible = $false
    $btn1.TextFrame2.TextRange.Text = [char]0x1F504 + " CAP NHAT NGAY"
    $btn1.TextFrame2.TextRange.Font.Size = 11
    $btn1.TextFrame2.TextRange.Font.Bold = $true
    $btn1.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::White)
    $btn1.TextFrame2.HorizontalAnchor = 2  # msoAnchorCenter
    $btn1.TextFrame2.VerticalAnchor = 3    # msoAnchorMiddle
    $btn1.Shadow.Visible = $true
    $btn1.Shadow.Type = 1
    $btn1.Shadow.OffsetX = 2; $btn1.Shadow.OffsetY = 2
    Write-Host "  [OK] Nut 1: CAP NHAT NGAY (xanh duong)"
    
    # Nut 2: TU DONG MOI 30 PHUT (Xanh la)
    $btn2 = $wsDash.Shapes.AddShape(5, 175, 62, 180, 35)
    $btn2.Name = "btnTuDong30Phut"
    $btn2.Fill.ForeColor.RGB = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(38, 128, 0))
    $btn2.Line.Visible = $false
    $btn2.TextFrame2.TextRange.Text = [char]0x23F0 + " TU DONG MOI 30 PHUT"
    $btn2.TextFrame2.TextRange.Font.Size = 11
    $btn2.TextFrame2.TextRange.Font.Bold = $true
    $btn2.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::White)
    $btn2.TextFrame2.HorizontalAnchor = 2
    $btn2.TextFrame2.VerticalAnchor = 3
    $btn2.Shadow.Visible = $true
    $btn2.Shadow.Type = 1
    $btn2.Shadow.OffsetX = 2; $btn2.Shadow.OffsetY = 2
    Write-Host "  [OK] Nut 2: TU DONG MOI 30 PHUT (xanh la)"
    
    # Nut 3: DUNG TU DONG (Do)
    $btn3 = $wsDash.Shapes.AddShape(5, 365, 62, 160, 35)
    $btn3.Name = "btnDungTuDong"
    $btn3.Fill.ForeColor.RGB = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(192, 0, 0))
    $btn3.Line.Visible = $false
    $btn3.TextFrame2.TextRange.Text = [char]0x23F9 + " DUNG TU DONG"
    $btn3.TextFrame2.TextRange.Font.Size = 11
    $btn3.TextFrame2.TextRange.Font.Bold = $true
    $btn3.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::White)
    $btn3.TextFrame2.HorizontalAnchor = 2
    $btn3.TextFrame2.VerticalAnchor = 3
    $btn3.Shadow.Visible = $true
    $btn3.Shadow.Type = 1
    $btn3.Shadow.OffsetX = 2; $btn3.Shadow.OffsetY = 2
    Write-Host "  [OK] Nut 3: DUNG TU DONG (do)"
    
    # === BUOC 4: Nhung VBA Module ===
    Write-Host ""
    Write-Host "Nhung VBA Macro..."
    
    # Xoa module cu neu co
    try {
        foreach ($comp in $wb.VBProject.VBComponents) {
            if ($comp.Name -eq "ModCapNhatEmail") {
                $wb.VBProject.VBComponents.Remove($comp)
                Write-Host "  Da xoa module VBA cu."
                break
            }
        }
    } catch {}
    
    $vbaCode = @"
Option Explicit

Private Declare PtrSafe Function SetTimer Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr, _
    ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
Private Declare PtrSafe Function KillTimer Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr) As Long

Private TimerID As LongPtr
Private Const INTERVAL_MS As Long = 1800000  ' 30 phut = 1800000ms

' === MACRO 1: Cap nhat ngay lap tuc ===
Public Sub CapNhatEmail()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(1)
    
    On Error Resume Next
    ws.Range("A2").Value = "Dang cap nhat... Vui long cho..."
    Application.StatusBar = "Dang quet va cap nhat du lieu email tu Outlook..."
    DoEvents
    On Error GoTo 0
    
    Dim scriptPath As String
    scriptPath = "C:\OutLook\update_data_live.ps1"
    
    If Dir(scriptPath) = "" Then
        MsgBox "Khong tim thay file script:" & vbCrLf & scriptPath & vbCrLf & vbCrLf & _
               "Hay chay lenh XUATLINK_OUTLOOK trong AutoCAD de tao lai.", vbExclamation, "Loi"
        Exit Sub
    End If
    
    Dim cmd As String
    cmd = "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & scriptPath & """"
    
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run cmd, 0, True
    Set wsh = Nothing
    
    Application.StatusBar = "Cap nhat hoan tat luc " & Format(Now, "dd/MM/yyyy HH:mm:ss")
    
    MsgBox "Da cap nhat xong du lieu email tu Outlook!" & vbCrLf & _
           "Tong so email: Xem Dashboard" & vbCrLf & _
           "Thoi gian: " & Format(Now, "dd/MM/yyyy HH:mm:ss"), vbInformation, "Thanh cong"
End Sub

' === MACRO 2: Bat tu dong cap nhat moi 30 phut ===
Public Sub CapNhatTuDong()
    If TimerID <> 0 Then
        KillTimer 0, TimerID
        TimerID = 0
    End If
    
    TimerID = SetTimer(0, 0, INTERVAL_MS, AddressOf TimerCallback)
    
    If TimerID <> 0 Then
        Application.StatusBar = "Da bat tu dong cap nhat moi 30 phut. Timer ID: " & TimerID
        MsgBox "Da bat che do TU DONG CAP NHAT moi 30 phut!" & vbCrLf & vbCrLf & _
               "Email se duoc quet va cap nhat lien tuc." & vbCrLf & _
               "Bam nut 'DUNG TU DONG' de tat.", vbInformation, "Tu dong cap nhat"
    Else
        MsgBox "Khong the bat timer tu dong.", vbExclamation, "Loi"
    End If
End Sub

' === Timer Callback (goi khi het 30 phut) ===
Private Sub TimerCallback(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal nIDEvent As LongPtr, ByVal dwTime As Long)
    On Error Resume Next
    KillTimer 0, nIDEvent
    
    Application.StatusBar = "Dang tu dong cap nhat email... " & Format(Now, "HH:mm:ss")
    CapNhatEmail
    
    ' Re-arm timer
    TimerID = SetTimer(0, 0, INTERVAL_MS, AddressOf TimerCallback)
    On Error GoTo 0
End Sub

' === MACRO 3: Dung tu dong cap nhat ===
Public Sub DungCapNhat()
    If TimerID <> 0 Then
        KillTimer 0, TimerID
        TimerID = 0
        Application.StatusBar = "Da tat che do tu dong cap nhat."
        MsgBox "Da TAT che do tu dong cap nhat.", vbInformation, "Dung tu dong"
    Else
        MsgBox "Hien khong co timer tu dong nao dang chay.", vbInformation, "Thong bao"
    End If
End Sub
"@
    
    $vbModule = $wb.VBProject.VBComponents.Add(1)  # vbext_ct_StdModule = 1
    $vbModule.Name = "ModCapNhatEmail"
    $vbModule.CodeModule.AddFromString($vbaCode)
    Write-Host "  [OK] Da nhung VBA Module: ModCapNhatEmail"
    
    # === BUOC 5: Gan macro vao cac nut ===
    Write-Host "Gan macro vao cac nut bam..."
    $wsDash.Shapes.Item("btnCapNhatNgay").OnAction = "ModCapNhatEmail.CapNhatEmail"
    $wsDash.Shapes.Item("btnTuDong30Phut").OnAction = "ModCapNhatEmail.CapNhatTuDong"
    $wsDash.Shapes.Item("btnDungTuDong").OnAction = "ModCapNhatEmail.DungCapNhat"
    Write-Host "  [OK] Da gan macro cho 3 nut bam."
    
    # === BUOC 6: Luu file ===
    Write-Host ""
    Write-Host "Luu file..."
    $excel.DisplayAlerts = $false
    
    # Neu file hien tai la .xlsx, can SaveAs thanh .xlsm
    $currentPath = $wb.FullName
    if ($currentPath -like "*.xlsx") {
        $newPath = $currentPath -replace '\.xlsx$', '.xlsm'
        $wb.SaveAs($newPath, 52)  # 52 = xlOpenXMLWorkbookMacroEnabled
        Write-Host "  Da luu lai thanh file .xlsm: $newPath"
    } else {
        $wb.Save()
        Write-Host "  Da luu file: $currentPath"
    }
    
    $excel.DisplayAlerts = $true
    
    Write-Host ""
    Write-Host "============================================="
    Write-Host "  HOAN TAT! DA TAO 3 NUT BAM THANH CONG"
    Write-Host "============================================="
    Write-Host ""
    Write-Host "  [XANH DUONG] CAP NHAT NGAY    - Quet email va cap nhat ngay"
    Write-Host "  [XANH LA]    TU DONG 30 PHUT  - Tu dong cap nhat moi 30 phut"  
    Write-Host "  [DO]          DUNG TU DONG     - Tat che do tu dong"
    Write-Host ""
    Write-Host "Hay quay lai Excel de su dung cac nut!"
    
} catch {
    Write-Host ""
    Write-Host "[LOI] $($_.Exception.Message)"
    Write-Host ""
    if ($_.Exception.Message -like "*programmatic access*" -or $_.Exception.Message -like "*VBProject*") {
        Write-Host "=========================================="
        Write-Host "  BAN CAN BAT QUYEN TRUY CAP VBA:"
        Write-Host "=========================================="
        Write-Host "  1. Trong Excel: File > Options > Trust Center"
        Write-Host "  2. Bam 'Trust Center Settings'"
        Write-Host "  3. Chon 'Macro Settings'"
        Write-Host "  4. Tick vao 'Trust access to the VBA project object model'"
        Write-Host "  5. Bam OK, roi chay lai script nay."
    }
    Write-Host $_.ScriptStackTrace
}
