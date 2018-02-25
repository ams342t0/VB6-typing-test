VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C00000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "THE TYPING TEST"
   ClientHeight    =   9045
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10650
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   10650
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.StatusBar sbStatBar 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   8550
      Width           =   10650
      _ExtentX        =   18785
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H000000FF&
      Height          =   3045
      Left            =   75
      TabIndex        =   2
      Top             =   5445
      Width           =   10470
      Begin RichTextLib.RichTextBox txtUserInput 
         Height          =   2790
         Left            =   120
         TabIndex        =   4
         Top             =   150
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   4921
         _Version        =   393217
         BackColor       =   16777152
         BorderStyle     =   0
         ScrollBars      =   3
         Appearance      =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         TextRTF         =   $"frmMain.frx":374C2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5265
      Left            =   75
      TabIndex        =   0
      Top             =   90
      Width           =   10470
      Begin VB.Timer Timer1 
         Left            =   8220
         Top             =   390
      End
      Begin RichTextLib.RichTextBox txtSourceText 
         Height          =   5010
         Left            =   120
         TabIndex        =   1
         Top             =   150
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   8837
         _Version        =   393217
         BackColor       =   12648384
         BorderStyle     =   0
         HideSelection   =   0   'False
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         Appearance      =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         TextRTF         =   $"frmMain.frx":3754F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "E&XIT"
   End
   Begin VB.Menu mnuNewGame 
      Caption         =   "N&EW GAME"
   End
   Begin VB.Menu mnuCheck 
      Caption         =   "CHECKING"
      Begin VB.Menu mnuActiveChecking 
         Caption         =   "CHECK WHILE TYPING"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuCheckCompos 
         Caption         =   "CHECK COMPOSITION"
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu mnuEndGame 
      Caption         =   "EN&D GAME"
   End
   Begin VB.Menu mnuScores 
      Caption         =   "&SCORES"
   End
   Begin VB.Menu mnuAdmin 
      Caption         =   "&ADMIN"
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "ABOUT"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetTick Lib "kernel32.dll" Alias "GetTickCount" () As Long

Private Sub Form_Load()
    EndGame
    InitControls
    Me.Show
End Sub

Private Sub mnuEditNew_Click()
    frmSourceText.Show vbModal, Me
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuActiveChecking_Click()
    mnuActiveChecking.Checked = Not mnuActiveChecking.Checked
End Sub

Private Sub mnuAdmin_Click()
    frmPassword.Show vbModal, Me
End Sub

Private Sub mnuCheckCompos_Click()
    HighlightMistakes
End Sub

Private Sub mnuEndGame_Click()
    EndGame
    frmResult.Show vbModal, Me
End Sub

Private Sub mnuExit_Click()
    r = MsgBox("Exit typing test?", vbYesNo, "EXIT TYPING TEST")
    
    If r = vbYes Then
        End
    End If
    
    Cancel = 1
End Sub

Private Sub mnuNewGame_Click()
   frmChooseLesson.Show vbModal, Me
   frmBegin.Show vbModal, Me
   BeginGame
End Sub

Private Sub mnuScores_Click()
    frmScores.Show vbModal, Me
End Sub

Private Sub Timer1_Timer()
    sbStatBar.Panels(5).Text = "TIME: " & Format$(TimeSerial(0, 0, (GetTick - lTimeStart) / 1000), "hh:nn:ss")
End Sub


Private Sub txtUserInput_Change()
    CounterCheck
    updateStatbar
End Sub

Private Sub txtUserInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyV Then
        If Shift = 2 Then
            Clipboard.Clear
        End If
    End If
End Sub


Sub BeginGame()
    'LoadSource
    mnuNewGame.Enabled = False
    mnuEndGame.Enabled = True
    mnuCheck.Enabled = True
    mnuActiveChecking.Checked = True
    InitVariables
    txtSourceText.SelStart = 0
    txtSourceText.SelLength = lTypeLen
    txtSourceText.SelColor = RGB(0, 0, 0)
    txtUserInput.Enabled = True
    txtSourceText.Enabled = True
    txtUserInput.SetFocus
    txtUserInput.Text = ""
    lTimeStart = GetTick
    Timer1.Enabled = True
    updateStatbar
End Sub


Sub EndGame()
    mnuEndGame.Enabled = False
    mnuNewGame.Enabled = True
    mnuCheck.Enabled = False
    txtUserInput.Enabled = False
    txtSourceText.Enabled = False
    Timer1.Enabled = False
End Sub


Sub InitControls()
    sbStatBar.Panels.Add 2, , ""
    sbStatBar.Panels.Add 3, , ""
    sbStatBar.Panels.Add 4, , ""
    sbStatBar.Panels.Add 4, , ""
    
    sbStatBar.Panels(1).Width = 2000
    sbStatBar.Panels(2).Width = 2000
    sbStatBar.Panels(3).Width = 2000
    sbStatBar.Panels(4).Width = 2000
    sbStatBar.Panels(5).Width = 2000
    
    sbStatBar.Panels(1).Text = "COMPLETED:"
    sbStatBar.Panels(2).Text = "MISTAKES:"
    sbStatBar.Panels(3).Text = "ACCURACY:"
    sbStatBar.Panels(4).Text = "SPEED:"
    sbStatBar.Panels(5).Text = "TIME:"
    
    txtUserInput.Text = ""
End Sub


Function CounterCheck()
    Dim i As Long
    Dim a As String
    Dim b As String
    
    
    lTypeLen = Len(txtUserInput.Text)
    
    If lTypeLen > Len(txtSourceText.Text) Then
        lTypeLen = Len(txtSourceText.Text)
    End If
    
    If mnuActiveChecking.Checked Then
        txtSourceText.Visible = False
        txtSourceText.SelStart = 0
        txtSourceText.SelLength = lTypeLen
        txtSourceText.SelColor = RGB(0, 0, 0)
        txtSourceText.SelUnderline = False
    End If
    
    lMistakes = 0
    
    For i = 1 To lTypeLen
        a = Mid$(txtUserInput.Text, i, 1)
        b = Mid$(txtSourceText.Text, i, 1)
        
        If StrComp(a, b, vbBinaryCompare) <> 0 Then
            
            If mnuActiveChecking.Checked Then
                txtSourceText.SelStart = i - 1
                txtSourceText.SelLength = 1
            
                If b = " " Then
                    txtSourceText.SelUnderline = True
                End If
                txtSourceText.SelColor = RGB(255, 0, 0)
            End If
            
            lMistakes = lMistakes + 1
        End If
    Next
    
    If mnuActiveChecking.Checked Then
        txtSourceText.Visible = True
    End If
    
    txtSourceText.SelStart = i - 1
    txtSourceText.SelLength = 1
    
End Function

Sub HighlightMistakes()
    Dim i As Long
    
    lTypeLen = Len(txtUserInput.Text)
    
    txtSourceText.Visible = False
    txtSourceText.SelStart = 0
    txtSourceText.SelLength = lTypeLen
    txtSourceText.SelColor = RGB(0, 0, 0)
    txtSourceText.SelUnderline = False
    txtSourceText.SelLength = 0
    
    For i = 1 To lTypeLen
        If StrComp(Mid$(txtUserInput.Text, i, 1), Mid$(txtSourceText.Text, i, 1), vbBinaryCompare) <> 0 Then
            txtSourceText.SelStart = i - 1
            txtSourceText.SelLength = 1
            
            If b = " " Then
               txtSourceText.SelUnderline = True
            End If
            
            txtSourceText.SelColor = RGB(255, 0, 0)
        End If
    Next
    
    txtSourceText.Visible = True
End Sub

Sub InitVariables()
    lScore = 0
    lCompleted = 0
    lSpeed = 0
    lTypeLen = 0
    Timer1.Interval = 1000
    txtUserInput.Text = ""
End Sub

Sub updateStatbar()
    On Error Resume Next
    
    lCompleted = 100 * Round(lTypeLen / Len(txtSourceText.Text), 2)
    sbStatBar.Panels(1).Text = "COMPLETED: " & lCompleted & "%"
    sbStatBar.Panels(2).Text = "MISTAKES: " & lMistakes
    lAccuracy = 100 * Round(1 - (lMistakes / lTypeLen), 2)
    sbStatBar.Panels(3).Text = "ACCURACY: " & lAccuracy & "%"
    lTotalTime = (GetTick - lTimeStart) / 1000
    lSpeed = Round(60 * (lTypeLen - lMistakes) / lTotalTime, 0)
    sbStatBar.Panels(4).Text = "SPEED: " & lSpeed & " CPM"
End Sub

