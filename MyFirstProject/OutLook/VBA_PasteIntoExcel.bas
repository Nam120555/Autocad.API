Option Explicit

Private Declare PtrSafe Function SetTimer Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr, _
    ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
Private Declare PtrSafe Function KillTimer Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr) As Long

Private TimerID As LongPtr
Private Const INTERVAL_MS As Long = 1800000  ' 30 phut = 1800000ms

' ============================================================
' MACRO 0: CHAY 1 LAN DUY NHAT DE TAO 3 NUT BAM
' Sau khi chay xong, ban co the xoa Sub nay di
' ============================================================
Public Sub TaoNutBam()
    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Worksheets(1)
    
    ' Xoa nut cu neu co
    Dim shp As Shape
    For Each shp In wsDash.Shapes
        If shp.Name = "btnCapNhatNgay" Or shp.Name = "btnTuDong30Phut" Or shp.Name = "btnDungTuDong" Then
            shp.Delete
        End If
    Next shp
    
    ' Chen dong huong dan Row 3
    With wsDash.Range("A3:H3")
        .Merge
        .Value = "Bam cac nut ben duoi de cap nhat du lieu email truc tiep (khong can dong file Excel)"
        .Font.Italic = True
        .Font.Size = 10
        .Font.Color = RGB(24, 43, 73)
        .HorizontalAlignment = xlCenter
    End With
    wsDash.Rows(3).RowHeight = 20
    wsDash.Rows(4).RowHeight = 42
    
    ' === NUT 1: CAP NHAT NGAY (Xanh duong) ===
    Dim btn1 As Shape
    Set btn1 = wsDash.Shapes.AddShape(msoShapeRoundedRectangle, 5, 50, 160, 35)
    With btn1
        .Name = "btnCapNhatNgay"
        .Fill.ForeColor.RGB = RGB(0, 112, 192)
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Text = ChrW(&H21BB) & " CAP NHAT NGAY"
        .TextFrame2.TextRange.Font.Size = 11
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.HorizontalAnchor = msoAnchorCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .Shadow.Visible = msoTrue
        .Shadow.OffsetX = 2
        .Shadow.OffsetY = 2
        .OnAction = "CapNhatEmail"
    End With
    
    ' === NUT 2: TU DONG 30 PHUT (Xanh la) ===
    Dim btn2 As Shape
    Set btn2 = wsDash.Shapes.AddShape(msoShapeRoundedRectangle, 175, 50, 180, 35)
    With btn2
        .Name = "btnTuDong30Phut"
        .Fill.ForeColor.RGB = RGB(38, 128, 0)
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Text = ChrW(&H23F0) & " TU DONG MOI 30 PHUT"
        .TextFrame2.TextRange.Font.Size = 11
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.HorizontalAnchor = msoAnchorCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .Shadow.Visible = msoTrue
        .Shadow.OffsetX = 2
        .Shadow.OffsetY = 2
        .OnAction = "CapNhatTuDong"
    End With
    
    ' === NUT 3: DUNG TU DONG (Do) ===
    Dim btn3 As Shape
    Set btn3 = wsDash.Shapes.AddShape(msoShapeRoundedRectangle, 365, 50, 160, 35)
    With btn3
        .Name = "btnDungTuDong"
        .Fill.ForeColor.RGB = RGB(192, 0, 0)
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Text = ChrW(&H23F9) & " DUNG TU DONG"
        .TextFrame2.TextRange.Font.Size = 11
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.HorizontalAnchor = msoAnchorCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .Shadow.Visible = msoTrue
        .Shadow.OffsetX = 2
        .Shadow.OffsetY = 2
        .OnAction = "DungCapNhat"
    End With
    
    MsgBox "Da tao thanh cong 3 nut bam!" & vbCrLf & vbCrLf & _
           "  [XANH DUONG] CAP NHAT NGAY" & vbCrLf & _
           "  [XANH LA] TU DONG 30 PHUT" & vbCrLf & _
           "  [DO] DUNG TU DONG" & vbCrLf & vbCrLf & _
           "Hay dong VBA Editor (Alt+F11) va thu bam nut!", vbInformation, "Thanh cong"
End Sub

' ============================================================
' MACRO 1: CAP NHAT NGAY LAP TUC
' Quet toan bo email tu Outlook va cap nhat vao Excel
' ============================================================
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

' ============================================================
' MACRO 2: BAT TU DONG CAP NHAT MOI 30 PHUT
' ============================================================
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

' ============================================================
' MACRO 3: DUNG TU DONG CAP NHAT
' ============================================================
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
