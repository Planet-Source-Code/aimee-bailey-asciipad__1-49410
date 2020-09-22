VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmGoTo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   825
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   825
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&GoTo"
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Line"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmGoTo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text > UpDown1.Max Then Text1.Text = UpDown1.Max
frmMain.goto Int(Text1.Text)
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
MsgBox frmMain.GetCurrentLine
MsgBox frmMain.LineCount

UpDown1.Max = frmMain.LineCount
UpDown1.Value = frmMain.GetCurrentLine
Text1.Text = UpDown1.Value

If frmMain.LineCount = 0 Then
    MsgBox "There is only 1 line to navigate!"
    Unload Me
End If
End Sub

Public Function GetLineCount()

For i = 1 To Len(frmMain.Text1.Text)
    'DoEvents
    If Mid(frmMain.Text1.Text, i, 1) = vbLf Then
        x = x + 1
    End If
Next i
MsgBox x
frmMain.LineCount = x
End Function
Private Sub UpDown1_Change()
Text1.Text = UpDown1.Value
End Sub
