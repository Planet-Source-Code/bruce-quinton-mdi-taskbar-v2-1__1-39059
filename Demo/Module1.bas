Attribute VB_Name = "Module1"
Option Explicit

Public fMainForm As frmMain
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const WM_SYSCOMMAND = &H112
Public Const WM_COMMAND = &H111


Sub Main()
    Set fMainForm = New frmMain
    fMainForm.Show
End Sub



