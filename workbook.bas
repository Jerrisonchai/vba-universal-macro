Option Explicit
Private Sub Workbook_Open()
Call capturetime
Dim sourceWorkbook As Workbook
Dim sourceSheet1 As Worksheet
Dim shape1 As Shape
Set sourceWorkbook = ThisWorkbook
Set sourceSheet1 = sourceWorkbook.Sheets("Dashboard")
OptimizedMode False
For Each shape1 In sourceSheet1.Shapes
    On Error Resume Next
    shape1.Fill.ForeColor.RGB = RGB(0, 0, 255)
    shape1.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
Next
sourceSheet1.Activate
Set sourceWorkbook = Nothing
Set sourceSheet1 = Nothing
Set shape1 = Nothing
Call captureendtime
  LoginInstance = 0
  Application.Visible = False
  frmLogin.Show
ThisWorkbook.Sheets("Dashboard").Activate
End Sub

