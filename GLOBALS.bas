Attribute VB_Name = "Module1"
Public strname As String
Public lScore As Long
Public lSpeed As Double
Public lTimeStart As Long
Public lMistakes As Long
Public lCompleted As Double
Public lTypeLen As Long
Public lAccuracy As Double
Public lTotalTime As Long
Public lactivesource As Long
Public isnewsource As Boolean

Public cnscores As ADODB.Connection
Public lID As Long


Public Function GetTotalScore() As Long
    GetTotalScore = (lTypeLen - lMistakes) + Round((lCompleted * 5 * 0.3) + (lSpeed * 0.4) + (lAccuracy * 0.3), 0)
End Function


Public Sub ConnectScores()
    Set cnscores = CreateObject("ADODB.CONNECTION")
    
    With cnscores
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source= " & App.Path & "\scores.mdb"
        .Mode = adModeReadWrite
        .CursorLocation = adUseServer
        .Open
    End With
End Sub

Public Sub CloseScores()
    cnscores.Close
    Set cnscores = Nothing
End Sub

Public Sub SaveScore()
    Dim STRSQL As String
    
    STRSQL = "INSERT INTO TBLSCORE ( USERNAME, SCORE, TOTALTIME, CHARSTYPED, MISTAKES, ACCURACY, SPEED,COMPLETED,LESSONID,DATETIMESTAMP)" & _
             " VALUES (" & _
             """" & strname & """" & "," & _
             lScore & "," & _
             lTotalTime & "," & _
             lTypeLen & "," & _
             lMistakes & "," & _
             lAccuracy & "," & _
             lSpeed & "," & _
             lCompleted & "," & _
             lID & "," & _
             "#" & Now() & "#" & ")"
    
    cnscores.Execute STRSQL
End Sub


Public Sub LoadSource()
    Dim rs As ADODB.Recordset
    
    ConnectScores
    
    Set rs = cnscores.Execute("SELECT SOURCETEXT FROM TBLSOURCE WHERE ISACTIVE")
    
    Form1.txtSourceText.Text = rs.Fields("sourcetext")
    
    Set rs = Nothing
    
    CloseScores
End Sub


Public Sub UserSelectLesson()
    Dim rs As ADODB.Recordset
    
    ConnectScores
    
    Set rs = cnscores.Execute("SELECT SOURCETEXT FROM TBLSOURCE WHERE ID=" & lID)
    
    Form1.txtSourceText.Text = rs.Fields("sourcetext")
    
    Set rs = Nothing
    
    CloseScores
End Sub


Public Function GetPassword() As String
    Dim rs As ADODB.Recordset
    
    ConnectScores
            
    Set rs = cnscores.Execute("select password from tblglobals")
    
    GetPassword = rs.Fields("password")
    
    Set rs = Nothing
        
    CloseScores
End Function
