Attribute VB_Name = "LogModule"
Public Sub LogWrite(ByVal Txt As String, Optional ByVal StringTai As Long)
    Dim LogWriteTxtPause As String
    Dim ErrorStatus As String
    'Dim SpaceSize As Long
SearchStringTai:
    Select Case StringTai
        Case 1
            ErrorStatus = "Error"
            'SpaceSize = 0
        Case 2
            ErrorStatus = "Run"
            'SpaceSize = 2
        Case 3
            ErrorStatus = "Info"
            'SpaceSize = 1
        Case Else
            StringTai = 2
            GoTo SearchStringTai
    End Select
    LogWriteTxtPause = Logfrm.LogLbl.Text + vbNewLine + "[" + Str(Now) + "]" + "[" + ErrorStatus + "]: " + Txt
    Logfrm.LogLbl.Text = LogWriteTxtPause
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

Public Sub TxtSave()
    Dim filepath, filenum, oldcontent
    filepath = App.Path & "\Log\Log.txt"
    '在畫面上放一個TextBox，命名為txtTempForLog，將Visible設定為False
    '-----↓把Log檔的舊內容讀出來，暫存在畫面上的txtTempForLog裡-----------
    filenum = FreeFile
    Main.TxtTempForLog.Text = ""
    Open filepath For Input As #filenum ' 開啟文字檔,開始讀出記錄
    ' 若不是空檔案,一行一行把txt讀出來放在txtTempForLog
    If EOF(filenum) = False Then ' 判斷 Test.txt 是不是空的檔案
        Do ' TextBox容量只有32KB大檔案請用RichTextBox
            Line Input #filenum, oldcontent
            Main.TxtTempForLog.SelText = oldcontent
        Loop Until EOF(filenum)
        Close #filenum
    End If
    Close #filenum
    '-----↑把Log檔的舊內容讀出來，暫存在畫面上的txtTempForLog裡-----------
    filenum = FreeFile ' 先把新now寫進去再把剛讀出來的txt從txtTempForLog寫進去
    'Open filepath For Append As #FileNum  '用Append會把新內容加在後面。我要把新內容加在最前面，所以需要把舊內容暫存在畫面上的txtTempForLog裡再貼進來
    'Print #FileNum, Now & "：" & message
    'Close #filenum
    Open filepath For Output As #filenum ' 開啟文字檔,準備寫入檔案
    Print #filenum, Logfrm.LogLbl.Text '把舊內容貼在新內容後面，寫入檔案
    Close #filenum
End Sub
