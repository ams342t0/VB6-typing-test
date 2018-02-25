VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmResult 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RESULTS"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4815
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&CLOSE"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   4575
   End
   Begin MSComctlLib.ListView lvStat 
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   4048
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "frmResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub INITSTAT()
    With lvStat
        .View = lvwReport
        .FullRowSelect = True
        .GridLines = True
        .HideColumnHeaders = True
        .LabelEdit = lvwManual
        .HideSelection = True
        .ColumnHeaders.Add 1, , , 3000
        .ColumnHeaders.Add 2, , , 1500
    End With
    
    lvStat.ListItems.Add 1, , "Characters Typed"
    lvStat.ListItems.Add 2, , "Mistakes"
    lvStat.ListItems.Add 3, , "Accuracy"
    lvStat.ListItems.Add 4, , "Total Time"
    lvStat.ListItems.Add 5, , "Typing Speed (CPM)"
    lvStat.ListItems.Add 6, , "Overall Score"
    
End Sub

Private Sub cmdOK_Click()
    strname = InputBox("NAME", "ENTER NAME", "USER")
    
    ConnectScores
    SaveScore
    CloseScores
    
    Unload Me
    
    frmScores.Show vbModal, Form1
End Sub

Private Sub Form_Load()
    INITSTAT
    SHOWSCORES
End Sub

Sub SHOWSCORES()
With lvStat
    .ListItems(1).SubItems(1) = lTypeLen
    .ListItems(2).SubItems(1) = lMistakes
    .ListItems(3).SubItems(1) = lAccuracy & "%"
    .ListItems(4).SubItems(1) = Format$(TimeSerial(0, 0, lTotalTime), "hh:nn:ss")
    .ListItems(5).SubItems(1) = lSpeed
    lScore = GetTotalScore
    .ListItems(6).SubItems(1) = lScore
End With
End Sub

