VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EDIT SOURCE"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10335
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   10335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "SAVE"
      Height          =   555
      Left            =   4140
      TabIndex        =   1
      Top             =   5190
      Width           =   1950
   End
   Begin RichTextLib.RichTextBox txtEdit 
      Height          =   5010
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   8837
      _Version        =   393217
      BackColor       =   12648384
      BorderStyle     =   0
      HideSelection   =   0   'False
      ScrollBars      =   3
      Appearance      =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      TextRTF         =   $"frmEdit.frx":0000
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
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSave_Click()
    If isnewsource Then
        cnscores.Execute "INSERT INTO tblSource (sourcetext) VALUES (""" & txtEdit.Text & """)"
    Else
        cnscores.Execute "UPDATE tblSource SET sourcetext=""" & txtEdit.Text & """ WHERE id = " & lID
    End If
    MsgBox "Saved"
End Sub
