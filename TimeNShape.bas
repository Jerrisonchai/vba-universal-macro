Option Explicit
Sub capturetime()   'Log macro start time
ThisWorkbook.Worksheets("LOG").Unprotect Password:="admin"
Dim timeWS As Worksheet, timeRange As Range
Dim vUser
vUser = Environ("Username")
Set timeWS = ThisWorkbook.Sheets("LOG")
Set timeRange = timeWS.Range("A" & timeWS.Rows.Count).End(xlUp).Offset(1)
    timeRange.Value = Application.Caller
    timeRange.Offset(, 1).Value = Now()
    timeRange.Offset(, 2).Value = vUser
    timeRange.Offset(, 3).Value = "Run Macro Started"
    timeRange.Offset(, 4).Value = ThisWorkbook.Name
Set timeWS = Nothing
Set timeRange = Nothing
Set vUser = Nothing
ThisWorkbook.Worksheets("LOG").Protect Password:="admin"
End Sub
Sub captureendtime()    'Log macro end time
ThisWorkbook.Worksheets("LOG").Unprotect Password:="admin"
Dim timeWS As Worksheet, timeRange As Range
Dim vUser
vUser = Environ("Username")
Set timeWS = ThisWorkbook.Sheets("LOG")
Set timeRange = timeWS.Range("A" & timeWS.Rows.Count).End(xlUp).Offset(1)
    timeRange.Value = Application.Caller
    timeRange.Offset(, 1).Value = Now()
    timeRange.Offset(, 2).Value = vUser
    timeRange.Offset(, 3).Value = "Run Macro Ended"
    timeRange.Offset(, 4).Value = ThisWorkbook.Name
Set timeWS = Nothing
Set timeRange = Nothing
Set vUser = Nothing
ThisWorkbook.Worksheets("LOG").Protect Password:="admin"
End Sub
Sub MyShape_Click()     'Auto change button color when clicked
Dim sh As Shape
Set sh = ActiveSheet.Shapes(Application.Caller)
If sh.Fill.ForeColor.RGB = RGB(0, 255, 127) Then
    sh.Fill.ForeColor.RGB = RGB(0, 0, 255)
Else
    sh.Fill.ForeColor.RGB = RGB(0, 255, 127)
End If
Call MyFont_Click
Set sh = Nothing
End Sub
Sub MyFont_Click()      'Auto change button color when clicked
Dim sh As Shape
Set sh = ActiveSheet.Shapes(Application.Caller)
If sh.TextFrame.Characters.Font.Color = RGB(255, 255, 255) Then
    sh.TextFrame.Characters.Font.Color = RGB(0, 100, 0)
Else
    sh.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
End If
Set sh = Nothing
End Sub
Public Sub OptimizedMode(ByVal enable As Boolean)       'Allow macro to run faster
     Application.EnableEvents = Not enable
     Application.ScreenUpdating = Not enable
     Application.EnableAnimations = Not enable
     Application.DisplayStatusBar = Not enable
     Application.PrintCommunication = Not enable
End Sub
Sub ProtectAllSheets()      'This code will protect all the sheets at one go
    Dim ws As Worksheet, Password As String
    Password = "admin"
    For Each ws In ThisWorkbook.Worksheets
        ws.Protect Password:=Password
    Next ws
    Set ws = Nothing
    Password = vbNullString
End Sub
Function Col_Letter(lngCol As Long) As String
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function
Sub OptOn()
OptimizedMode True
End Sub
Sub OptOff()
OptimizedMode False
End Sub
Sub CheckEvents()
MsgBox "Events Are Currently " & IIf(Application.EnableEvents, "Enabled", "Disabled"), vbInformation, "EnableEvents Status"
MsgBox "ScreenUpdating Are Currently " & IIf(Application.ScreenUpdating, "Enabled", "Disabled"), vbInformation, "ScreenUpdating Status"
MsgBox "EnableAnimations Are Currently " & IIf(Application.EnableAnimations, "Enabled", "Disabled"), vbInformation, "EnableAnimations Status"
MsgBox "DisplayStatusBar Are Currently " & IIf(Application.DisplayStatusBar, "Enabled", "Disabled"), vbInformation, "DisplayStatusBar Status"
MsgBox "PrintCommunication Are Currently " & IIf(Application.PrintCommunication, "Enabled", "Disabled"), vbInformation, "PrintCommunication Status"
End Sub

