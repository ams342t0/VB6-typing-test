VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmScores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HIGH SCORES"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   7335
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "CLOSE"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   120
      TabIndex        =   1
      Top             =   4680
      Width           =   7095
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4455
      Left            =   60
      TabIndex        =   0
      Top             =   105
      Width           =   7230
      _ExtentX        =   12753
      _ExtentY        =   7858
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "NAME"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "TIME"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "COMPLETED"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ACCURACY"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "SPEED"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "SCORE"
         Object.Width           =   1235
      EndProperty
   End
End
Attribute VB_Name = "frmScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rs As Recordset
    Dim lv As ListItem
    
    ConnectScores
    
    Set rs = cnscores.Execute("SELECT * FROM TBLSCORE ORDER BY  SCORE DESC")
    
    With rs
        If Not (.EOF And .BOF) Then
            ListView1.ListItems.Clear
            
            While Not .EOF
                Set lv = ListView1.ListItems.Add(, , rs.Fields("USERNAME"))
                lv.SubItems(1) = Format$(TimeSerial(0, 0, rs.Fields("TOTALTIME")), "HH:NN:SS")
                lv.SubItems(2) = rs.Fields("COMPLETED") & "%"
                lv.SubItems(3) = rs.Fields("ACCURACY") & "%"
                lv.SubItems(4) = rs.Fields("SPEED")
                lv.SubItems(5) = rs.Fields("SCORE")
                .MoveNext
            Wend
        End If
    End With
    
    CloseScores
    
        
End Sub



