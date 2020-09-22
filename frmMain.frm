VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Untitled - AsciiPad"
   ClientHeight    =   4545
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   6015
   ClipControls    =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   6015
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Interval        =   300
      Left            =   3120
      Top             =   2280
   End
   Begin VB.Timer Timer2 
      Interval        =   2000
      Left            =   3960
      Top             =   2280
   End
   Begin MSComDlg.CommonDialog cmd2 
      Left            =   1200
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dPrint 
      Left            =   360
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   2520
      Top             =   2280
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print"
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete Selection"
            ImageKey        =   "Delete"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            Object.ToolTipText     =   "Find"
            ImageKey        =   "Find"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "About"
            Object.ToolTipText     =   "About AsciiPad"
            ImageKey        =   "Help"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   4275
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4948
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "10/23/2003"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   360
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   240
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0554
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0666
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0778
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":088A
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":099C
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0AAE
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0BC0
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CD2
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1024
            Key             =   "Delete"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      Height          =   3615
      HideSelection   =   0   'False
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   360
      Width           =   5895
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu new_mnu 
         Caption         =   "New"
         Begin VB.Menu newwindow_mnu 
            Caption         =   "&Window"
            Shortcut        =   {F2}
         End
         Begin VB.Menu mnuFileNew 
            Caption         =   "&New File"
            Shortcut        =   ^N
         End
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrintPreview 
         Caption         =   "Print Setup..."
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu LastFile 
         Caption         =   "none"
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu undo_mnu 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu line2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu delete_mnu 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu find_mnu 
         Caption         =   "Find..."
         Shortcut        =   ^F
      End
      Begin VB.Menu find_next 
         Caption         =   "Find Next"
         Shortcut        =   {F3}
      End
      Begin VB.Menu replace_mnu 
         Caption         =   "Replace..."
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
      End
      Begin VB.Menu selectall_mnu 
         Caption         =   "Select All"
      End
      Begin VB.Menu mnuEditdatetime_mnu 
         Caption         =   "Date && Time"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Status &Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu styles_mnu 
         Caption         =   "&Styles..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuHelpSearchForHelpOn 
         Caption         =   "&Search For Help On..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About "
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Public CurrentFile As String
Public Saved As Boolean
Public TopHinge As Integer
Public BottomHinge As Integer
Const DefaultHeight = 3615
Public linelen As Integer
Public FindPos As Integer
Public FindLength As Integer
Public FindString As String
Public LineCount As Integer
Public xLineCount As Integer

Private Sub Command1_Click()
cmd2.ShowColor
Text1.BackColor = cmd2.Color
End Sub

Private Sub Command2_Click()
Text1.BackColor = CLng(&H8000000F)
End Sub

Public Function DoStyle(font As String, size As Long, bgcolor As Long, forecolor As Long)
Text1.font = font
Text1.font.size = size
Text1.BackColor = bgcolor
Text1.forecolor = forecolor
End Function


Private Sub delete_mnu_Click()
Text1.SetFocus
SendKeys (" ")
End Sub

Private Sub find_mnu_Click()
frmFind.Show vbModal, Me
End Sub

Private Sub find_next_Click()
On Error GoTo err
With Text1
If Len(Trim(FindString)) = 0 Then find_mnu_Click: GoTo err1
    .SelStart = InStr(.SelStart + .SelLength, LCase(.Text), LCase(FindString)) - 1


frmMain.FindPos = .SelStart
frmMain.FindLength = Len(FindString)

.SelLength = Len(FindString)
Exit Sub
err:
MsgBox "No More Matches Found!", vbInformation, "Find - AsciiPad"
err1:
End With

End Sub

Private Sub Form_Initialize()
If Trim(Command$) = "-N" Then
        Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000) + 1000
        Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000) + 1000
Else
        Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
        Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
        If Len(Trim(Command$)) > 0 Then
            OpenFile Replace(Command$, Chr(34), "")
        End If
End If
GetCurrentPos
End Sub

Private Sub Form_Load()
    
    'Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    'Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    CurrentFile = "Untitled"
    Saved = True
    DoLastFile
    
    
End Sub


Private Sub Form_Resize()
Dim a, b As Boolean

    On Error Resume Next
    Text1.Width = Me.Width - 140
       
    ResizeTextBox
    
End Sub

Public Function ResizeTextBox()
Dim a, b As Boolean
Dim ah, bh, ch As Integer

    On Error Resume Next
    
    a = tbToolBar.Visible
    b = sbStatusBar.Visible
    ah = tbToolBar.Height
    bh = sbStatusBar.Height
    ch = Me.Height - tbToolBar.Height - 1080
    
    If a = False And b = False Then
        Text1.Height = ch + ah + bh
        Text1.Top = 0
    ElseIf a = False And b = True Then
        Text1.Height = ch + ah
        Text1.Top = 0
    ElseIf a = True And b = False Then
        Text1.Height = ch + bh
        Text1.Top = 360
    ElseIf a = True And b = True Then
        Text1.Height = ch
        Text1.Top = 360
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        'SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        'SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
    
    If Saved = False Then
            x = MsgBox("Do you want to save your work before you exit?" & vbCrLf & vbCrLf & GetFile(CurrentFile), vbYesNoCancel, "AsciiPad - Exit")
            If x = vbYes Then
                mnuFileSave_Click
            ElseIf x = vbCancel Then
                Cancel = 1
                Exit Sub
            End If
            End
    Else
aa:
    End
    End If
End Sub

Private Sub goto_mnu_Click()
On Error Resume Next
GetCurrentPos
frmGoTo.Show vbModal, Me
End Sub

Private Sub LastFile_Click()
OpenFile GetSetting(App.Title, "Settings", "LastFile", "Untitled")
End Sub

Public Function OpenFile(file As String)
    h = FreeFile
    Open file For Input As #h
    Text1.Text = Input(LOF(h), h)
    Close #h
    CurrentFile = file
    Saved = True
End Function

Private Sub mnuFileSave_Click()
On Error GoTo err
If CurrentFile = "Untitled" Then
    h = FreeFile
    
    With dlgCommonDialog
        .DialogTitle = "Save As..."
        .CancelError = True
        .Filter = "ASCII Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        .ShowSave
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
        CurrentFile = sFile
    End With
    
    Open CurrentFile For Output As #h
    Print #h, Text1.Text
    Close #h
    Saved = True
    SaveSetting App.Title, "Settings", "LastFile", CurrentFile
Else
    h = FreeFile
    Open CurrentFile For Output As #h
    Print #h, Text1.Text
    Close #h
    Saved = True
    SaveSetting App.Title, "Settings", "LastFile", CurrentFile
End If
err:
End Sub

Private Sub mnuFileSaveAs_Click()
On Error GoTo err
    Dim sFile As String
    
    With dlgCommonDialog
        .DialogTitle = "Save As..."
        .CancelError = trie
        .Filter = "ASCII Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        .ShowSave
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With

    h = FreeFile
    Open sFile For Output As #h
    Print #h, Text1.Text
    Close #h
    Saved = True
    SaveSetting App.Title, "Settings", "LastFile", CurrentFile
err:
End Sub

Private Sub newwindow_mnu_Click()
'Shell App.Path & "\" & App.EXEName & " -N", vbNormalFocus
End Sub

Private Sub replace_mnu_Click()
frmFind.Show vbModal, Me
End Sub

Private Sub selectall_mnu_Click()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub styles_mnu_Click()
frmStyles.Show vbModal, Me
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComCtlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "New"
            mnuFileNew_Click
        Case "Open"
            mnuFileOpen_Click
        Case "Save"
            mnuFileSave_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Cut"
            mnuEditCut_Click
        Case "Copy"
            mnuEditCopy_Click
        Case "Paste"
            mnuEditPaste_Click
        Case "About"
            mnuHelpAbout_Click
        Case "Delete"
            delete_mnu_Click
        Case "Find"
            find_mnu_Click
    End Select
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuHelpSearchForHelpOn_Click()
    Dim nRet As Integer

    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
        If err Then
            MsgBox err.Description
        End If
    End If

End Sub

Private Sub mnuHelpContents_Click()
    Dim nRet As Integer

    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
        If err Then
            MsgBox err.Description
        End If
    End If

End Sub




Private Sub mnuViewOptions_Click()
    frmOptions.Show vbModal, Me
End Sub

Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sbStatusBar.Visible = mnuViewStatusBar.Checked
    ResizeTextBox
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    tbToolBar.Visible = mnuViewToolbar.Checked
    ResizeTextBox
End Sub

Private Sub mnuEditdatetime_mnu_Click()
    Text1.SelText = FormatDateTime(Time$, vbLongTime) & " " & FormatDateTime(Date$, vbGeneralDate)
End Sub

Private Sub mnuEditPaste_Click()
    Dim x As String
    x = Clipboard.GetText
    Text1.SelText = x
End Sub

Private Sub mnuEditCopy_Click()
    Clipboard.SetText Text1.SelText
End Sub

Private Sub mnuEditCut_Click()
    Clipboard.SetText Text1.SelText
    Text1.SelText = ""
End Sub

Private Sub mnuEditUndo_Click()
    MsgBox "Add 'mnuEditUndo_Click' code."
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFilePrint_Click()
On Error GoTo err
dPrint.Flags = 0
dPrint.CancelError = True
dPrint.ShowPrinter
Printer.Print Text1.Text
err:
End Sub

Private Sub mnuFilePrintPreview_Click()
On Error GoTo err
dPrint.Flags = &H40
dPrint.CancelError = True
dPrint.ShowPrinter
err:
End Sub

Private Sub mnuFileproperties_mnu_Click()
    MsgBox "Add 'mnuFileproperties_mnu_Click' code."
End Sub

Private Sub mnuFileOpen_Click()
On Error GoTo err
    Dim sFile As String
    
    With dlgCommonDialog
        .DialogTitle = "Open"
        .CancelError = trie
        .Filter = "ASCII Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    
    h = FreeFile
    
    Open sFile For Input As #h
    Text1.Text = Input(LOF(h) - 1, h)
    Close #h
    CurrentFile = sFile
    
    DoLastFile
    Exit Sub
err:
TryNormalOpen sFile
End Sub

Public Function TryNormalOpen(file As String)
Dim x
x = FreeFile
Close x
Text1.Text = ""
Open file For Input As #x
Do Until EOF(x)
DoEvents
    Input #x, a$
    Text1.Text = Text1.Text & a$ & vbCrLf
Loop
Close #x
End Function

Private Sub mnuFileNew_Click()
    Text1.Text = ""
    CurrentFile = "Untitled"
    'If Saved = False Then
    '    mnuFileSave_Click
    'End If
    
End Sub

Private Sub Text1_Change()
Saved = False
GetCurrentPos
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
GetCurrentPos
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
GetCurrentPos
End Sub

Private Sub Timer1_Timer()
Me.Caption = GetFile(CurrentFile) & " - AsciiPad"
End Sub

Public Function GetFile(file As String) As String
Dim x, i As Integer
If file = "Untitled" Then GetFile = "Untitled": Exit Function

For i = 1 To Len(file)
    If Mid(file, i, 1) = "\" Then x = i
Next i
GetFile = Mid(file, x + 1, 255)
End Function

Public Function DoLastFile()
On Error GoTo err
Dim x As String
x = GetSetting(App.Title, "Settings", "LastFile", "none")
    
    If x = "none" Then
        LastFile.Caption = x
        LastFile.Enabled = False
    Else
        LastFile.Caption = GetFile(x)
        LastFile.Enabled = True
    End If
    
    Exit Function
err:
    LastFile.Enabled = False
End Function

Public Function GetCurrentPos()

If InStr(1, Text1.Text, vbLf) > 0 Then
    xx = (Text1.SelStart - GetCurrentCol)
Else
    xx = GetCurrentCol
End If

sbStatusBar.Panels(3).Text = "Line " & GetCurrentLine '& " Col " & xx
End Function

Public Function GetCurrentCol() As Integer
Dim i As Integer
On Error Resume Next
With Text1

    xx = Int(.SelStart - 220)
    If xx < 0 Then xx = 0

    For i = .SelStart + 1 To xx Step -1
        If Mid(.Text, i, 1) = vbCr Then
            GetCurrentCol = .SelStart - i
            Exit Function
        ElseIf Mid(.Text, i, 1) = vbLf Then
            GetCurrentCol = i
            Exit Function
        End If
    Next i
    
err:
End With
End Function

Public Function GetCurrentLine() As Integer
Dim linelen As Long
LineCount = 0
If Len(Text1.Text) > 0 Then
For i = 1 To Text1.SelStart
If Mid(Text1.Text, i, 2) = CStr(vbCrLf) Then LineCount = LineCount + 1
linelen = i
Next i
End If
GetCurrentLine = LineCount
End Function

Private Sub undo_mnu_Click()
Text1.SetFocus
SendKeys ("^Z")
End Sub
