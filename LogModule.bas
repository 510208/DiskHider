Attribute VB_Name = "LogModule"
Public Sub LogWrite(ByVal Txt As String, Optional ByVal StringTai As Long)
    Dim LogWriteTxtPause As String
    Dim ErrorStatus As String
    Dim SpaceSize As Long
SearchStringTai:
    Select Case StringTai
        Case 1
            ErrorStatus = "Error"
            SpaceSize = 0
        Case 2
            ErrorStatus = "Run"
            SpaceSize = 2
        Case 3
            ErrorStatus = "Info"
            SpaceSize = 1
        Case Else
            StringTai = 2
            GoTo SearchStringTai
    End Select
    LogWriteTxtPause = Main.LogLbl.Text + vbNewLine + "[" + Str(Now) + "]" + "[" + ErrorStatus + "]: " + Space(SpaceSize) + Txt
    Main.LogLbl.Text = LogWriteTxtPause
    Debug.Print LogWriteTxtPause
End Sub

Public Function AscCodePassWord(Txt)
    Out = ""
    For i = 1 To Len(Txt)
        Out = Out & Format(Asc(Mid(Txt, i, 1)), "00000000")
    Next i
    AscCodePassWord = Out
End Function

Public Function ChAscCodePassWord(Txt)
    Out = ""
    For i = 1 To Len(Txt) Step 8
        Out = Out & Chr(Val(Mid(Txt, i, 8)))
    Next i
    ChAscCodePassWord = Out
End Function

