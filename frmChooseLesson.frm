VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChooseLesson 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SOURCE TEXT"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7185
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   7185
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "CLOSE"
      Height          =   585
      Left            =   2730
      TabIndex        =   1
      Top             =   4830
      Width           =   1680
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4680
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   7020
      _ExtentX        =   12383
      _ExtentY        =   8255
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "SOURCE TEXT"
         Object.Width           =   5292
      EndProperty
   End
End
Attribute VB_Name = "frmChooseLesson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset


Private Sub cmdOK_Click()
    UserSelectLesson
    Unload Me
End Sub

Private Sub Form_Load()
    LoadSourceTexts
End Sub

Sub LoadSourceTexts()
    Dim lv As ListItem
    
    ConnectScores
    
    Set rs = cnscores.Execute("SELECT * FROM TBLSOURCE ORDER BY ID")
    
    With ListView1
        .ListItems.Clear
        
        While Not rs.EOF
            Set lv = .ListItems.Add(, , Format$(rs.Fields("ID"), "0000"))
            lv.SubItems(1) = Mid$(rs.Fields("sourcetext"), 1, 30) & "..."
            lv.Checked = rs.Fields("isactive")
            
            If lv.Checked Then
                lID = rs.Fields("id")
            End If
            
            rs.MoveNext
        Wend
        
    End With
    
    
    CloseScores
End Sub


Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    lID = CLng(Item.Text)
End Sub


