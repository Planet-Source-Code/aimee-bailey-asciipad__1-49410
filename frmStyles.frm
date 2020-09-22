VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmStyles 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Editor Style Options"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CMD 
      Left            =   240
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   13
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   12
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Editor Style"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.CheckBox Check1 
         Caption         =   "Bold"
         Height          =   255
         Left            =   720
         TabIndex        =   11
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   255
         Left            =   4200
         TabIndex        =   10
         Top             =   1320
         Width           =   255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   255
         Left            =   4200
         TabIndex        =   9
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3120
         ScaleHeight     =   195
         ScaleWidth      =   1035
         TabIndex        =   8
         Top             =   1320
         Width           =   1095
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00000000&
         Height          =   255
         Left            =   3120
         ScaleHeight     =   195
         ScaleWidth      =   1035
         TabIndex        =   7
         Top             =   840
         Width           =   1095
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frmStyles.frx":0000
         Left            =   720
         List            =   "frmStyles.frx":0031
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   840
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   720
         Sorted          =   -1  'True
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label Label4 
         Caption         =   "BackColor"
         Height          =   255
         Left            =   2280
         TabIndex        =   6
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "ForeColor"
         Height          =   255
         Left            =   2280
         TabIndex        =   5
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Size"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Font"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmStyles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo err
CMD.CancelError = True
CMD.Color = Picture1.BackColor
CMD.ShowColor
Picture1.BackColor = CMD.Color
err:
End Sub

Private Sub Command2_Click()
On Error GoTo err
CMD.CancelError = True
CMD.Color = Picture2.BackColor
CMD.ShowColor
Picture2.BackColor = CMD.Color
err:
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
frmMain.Text1.font = Combo1.Text
frmMain.Text1.font.size = Combo2.Text
frmMain.Text1.forecolor = Picture1.BackColor
frmMain.Text1.BackColor = CLng(Picture2.BackColor)

If Check1.Value = 1 Then
    frmMain.Text1.font.Bold = True
Else
    frmMain.Text1.font.Bold = False
End If

Unload Me
End Sub

Private Sub Command5_Click()
frmMain.DoStyle Combo1.Text, Combo2.Text, Picture2.BackColor, Picture1.BackColor
End Sub

Private Sub Form_Load()
For i = 0 To Screen.FontCount - 1
Combo1.AddItem Screen.Fonts(i)
Next i
Combo1.Text = frmMain.Text1.font
Combo2.Text = frmMain.Text1.font.size
Picture1.BackColor = frmMain.Text1.forecolor
Picture2.BackColor = frmMain.Text1.BackColor

If frmMain.Text1.font.Bold = True Then
    Check1.Value = 1
Else
    Check1.Value = 0
End If
End Sub
