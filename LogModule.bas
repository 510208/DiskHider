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
    '�b�e���W��@��TextBox�A�R�W��txtTempForLog�A�NVisible�]�w��False
    '-----����Log�ɪ��¤��eŪ�X�ӡA�Ȧs�b�e���W��txtTempForLog��-----------
    filenum = FreeFile
    Main.TxtTempForLog.Text = ""
    Open filepath For Input As #filenum ' �}�Ҥ�r��,�}�lŪ�X�O��
    ' �Y���O���ɮ�,�@��@���txtŪ�X�ө�btxtTempForLog
    If EOF(filenum) = False Then ' �P�_ Test.txt �O���O�Ū��ɮ�
        Do ' TextBox�e�q�u��32KB�j�ɮ׽Х�RichTextBox
            Line Input #filenum, oldcontent
            Main.TxtTempForLog.SelText = oldcontent
        Loop Until EOF(filenum)
        Close #filenum
    End If
    Close #filenum
    '-----����Log�ɪ��¤��eŪ�X�ӡA�Ȧs�b�e���W��txtTempForLog��-----------
    filenum = FreeFile ' ����snow�g�i�h�A���Ū�X�Ӫ�txt�qtxtTempForLog�g�i�h
    'Open filepath For Append As #FileNum  '��Append�|��s���e�[�b�᭱�C�ڭn��s���e�[�b�̫e���A�ҥH�ݭn���¤��e�Ȧs�b�e���W��txtTempForLog�̦A�K�i��
    'Print #FileNum, Now & "�G" & message
    'Close #filenum
    Open filepath For Output As #filenum ' �}�Ҥ�r��,�ǳƼg�J�ɮ�
    Print #filenum, Logfrm.LogLbl.Text '���¤��e�K�b�s���e�᭱�A�g�J�ɮ�
    Close #filenum
End Sub
