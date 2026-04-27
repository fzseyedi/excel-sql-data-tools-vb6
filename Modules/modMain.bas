Attribute VB_Name = "modMain"
Option Explicit

Public Sub Main()
    On Error GoTo ErrorHandler

    Load frmMain
    frmMain.Show

    Exit Sub

ErrorHandler:
    MsgBox "Program startup error:" & vbCrLf & Err.Number & " - " & Err.Description, vbCritical, "Startup Error"
End Sub

Public Sub CenterFormOnScreen(xForm As Form)
    xForm.Left = (Screen.Width - xForm.Width) \ 2
    xForm.Top = (Screen.Height - xForm.Height) \ 2
    
    If xForm.Top < 0 Then xForm.Top = 0
    If xForm.Left < 0 Then xForm.Left = 0
End Sub

