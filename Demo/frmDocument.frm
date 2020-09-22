VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDocument 
   Caption         =   "frmDocument"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6540
   Icon            =   "frmDocument.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4725
   ScaleWidth      =   6540
   Begin VB.TextBox txt_Title 
      Height          =   285
      Left            =   600
      TabIndex        =   8
      Top             =   75
      Width           =   1455
   End
   Begin VB.TextBox txt_message 
      Height          =   3135
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   6255
   End
   Begin VB.CheckBox Chk_Flash 
      Caption         =   "Flash"
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   3840
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Flash Colour"
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   3840
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog GetColour 
      Left            =   3720
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Colours"
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ChangeFont"
      Height          =   315
      Left            =   2640
      TabIndex        =   2
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Selected"
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Unselected"
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   4320
      Top             =   4200
   End
   Begin VB.Label Title 
      AutoSize        =   -1  'True
      Caption         =   "Title:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   345
   End
   Begin VB.Label lbl_FlashColour 
      Caption         =   "0"
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   3840
      Width           =   975
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
GetColour.ShowColor

    fMainForm.TaskBar1.Set_Unselected_Colour Me.hWnd, GetColour.Color
    
End Sub

Private Sub Command2_Click()
GetColour.ShowColor
    
    fMainForm.TaskBar1.Set_Selected_Colour Me.hWnd, GetColour.Color

End Sub

Private Sub Command3_Click()
GetColour.ShowColor
 fMainForm.TaskBar1.Font_Color Me.hWnd, GetColour.Color
    
End Sub

Private Sub Command4_Click()
GetColour.ShowColor

    lbl_FlashColour.Caption = GetColour.Color

End Sub

Private Sub Form_GotFocus()
txt_Title = Me.Caption

End Sub

Private Sub Timer1_Timer()
If Chk_Flash.Value = 1 Then
    fMainForm.TaskBar1.FlashMe Me.hWnd, lbl_FlashColour.Caption, 5
End If

End Sub

Private Sub txt_Title_Change()
Me.Caption = txt_Title
End Sub
