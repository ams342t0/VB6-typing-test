VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSourceText 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SOURCE TEXT"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   7215
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClearScores 
      Caption         =   "CLEAR SCORES"
      Height          =   585
      Left            =   5430
      TabIndex        =   5
      Top             =   4140
      Width           =   1680
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "CLOSE"
      Height          =   585
      Left            =   5430
      TabIndex        =   4
      Top             =   2235
      Width           =   1680
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "DELETE"
      Height          =   585
      Left            =   5430
      TabIndex        =   3
      Top             =   1515
      Width           =   1680
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "EDIT"
      Height          =   585
      Left            =   5415
      TabIndex        =   2
      Top             =   795
      Width           =   1680
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "NEW"
      Height          =   585
      Left            =   5415
      TabIndex        =   1
      Top             =   90
      Width           =   1680
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4680
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   8255
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
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
Attribute VB_Name = "frmSourceText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset


Private Sub cmdClearScores_Click()
    Dim r
    
    On Error Resume Next
    
    r = MsgBox("Clear scores?", vbYesNo + vbExclamation, "CLEAR SCORES")
    
    If r = vbYes Then
        cnscores.Execute "DELETE FROM tblscore"
        MsgBox "Source deleted.", vbOKOnly, "CLEAR SCORES"
    End If
End Sub

Private Sub cmdDelete_Click()
    Dim r
    
    On Error Resume Next
    
    r = MsgBox("Delete this source?", vbYesNo + vbExclamation, "DELETE SOURCE TEXT")

    If r = vbYes Then
        cnscores.Execute "DELETE FROM tblSource WHERE Id = " & lID
        MsgBox "Source deleted.", vbOKOnly, "DELETE SOURCE"
    End If
    
    Set rs = cnscores.Execute("SELECT ID FROM TBLSOURCE WHERE ISACTIVE")
    
    If rs.EOF And rs.BOF Then
        Set rs = Nothing
        
        Set rs = CreateObject("adodb.recordset")
        rs.Open "SELECT * FROM tblSource", cnscores, adOpenDynamic, adLockOptimistic
        rs.Fields("isactive") = True
        rs.Update
        rs.Close
        
        Set rs = Nothing
    End If
    
    LoadSourceTexts
End Sub

Private Sub cmdEdit_Click()
    Dim rs As ADODB.Recordset
    
    Set rs = cnscores.Execute("SELECT SOURCETEXT FROM TBLSOURCE WHERE id=" & lID)
    
    frmEdit.txtEdit.Text = rs.Fields("sourcetext")
    
    Set rs = Nothing

    isnewsource = False
    
    frmEdit.Show vbModal, Me
    LoadSourceTexts
End Sub

Private Sub cmdNew_Click()
    isnewsource = True
    frmEdit.txtEdit.Text = ""
    frmEdit.Show vbModal, Me
    LoadSourceTexts
End Sub

Private Sub Form_Load()
    ConnectScores
    LoadSourceTexts
End Sub

Sub LoadSourceTexts()
    Dim lv As ListItem
    
    Set rs = cnscores.Execute("SELECT * FROM TBLSOURCE ORDER BY ID")
    
    With ListView1
        .ListItems.Clear
        
        While Not rs.EOF
            Set lv = .ListItems.Add(, , Format$(rs.Fields("ID"), "0000"))
            lv.SubItems(1) = Mid$(rs.Fields("sourcetext"), 1, 30) & "..."
            lv.Checked = rs.Fields("isactive")
            rs.MoveNext
        Wend
        
    End With
    
    lID = 1
    
    If ListView1.ListItems.Count > 1 Then
        cmdDelete.Enabled = True
    Else
        cmdDelete.Enabled = False
    End If
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CloseScores
End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim lv As ListItem
    
    For Each lv In ListView1.ListItems
        If lv.Index <> Item.Index Then
            lv.Checked = False
        End If
    Next
    
    Item.Selected = True
    
    lID = CLng(Item.Text)
    
    cnscores.Execute "UPDATE tblSource SET isactive = false"
    cnscores.Execute "UPDATE tblsource set isactive =" & Item.Checked & " where id = " & lID
    
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    lID = CLng(Item.Text)
End Sub


