VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Find"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check1 
      Caption         =   "UnCase Sensitive"
      Height          =   255
      Left            =   1080
      TabIndex        =   8
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   285
      Left            =   4440
      TabIndex        =   7
      Top             =   480
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Replace"
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   5535
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Top             =   240
         Width           =   3015
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Replace"
         Height          =   285
         Left            =   4200
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "With What"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Find"
      Default         =   -1  'True
      Height          =   285
      Left            =   4440
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Find What:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo err
With frmMain.Text1
If Check1.Value = 0 Then
    .SelStart = InStr(1, .Text, Text1.Text) - 1
Else
    .SelStart = InStr(1, LCase(.Text), LCase(Text1.Text)) - 1
End If

frmMain.FindPos = .SelStart
frmMain.FindLength = Len(Text1.Text)
frmMain.FindString = Text1.Text

.SelLength = Len(Text1.Text)
Exit Sub
err:
MsgBox "No Matches Found!", vbInformation, "Find - AsciiPad"
End With

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()

frmMain.Text1.SelText = Text2.Text

Unload Me
End Sub

Private Sub Form_Load()
Text1.Text = frmMain.FindString
    
    If frmMain.Text1.SelLength > 0 Then
        If Len(Trim(frmMain.FindString)) > 0 Then
            Command3.Enabled = True
        Else
            Command3.Enabled = False
        End If
    Else
            Command3.Enabled = False
    End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1_Click
End If
End Sub
