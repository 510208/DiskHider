VERSION 5.00
Begin VB.Form Main 
   Appearance      =   0  '����
   BackColor       =   &H80000005&
   BorderStyle     =   1  '��u�T�w
   Caption         =   "DiskHider"
   ClientHeight    =   2460
   ClientLeft      =   14040
   ClientTop       =   2925
   ClientWidth     =   4935
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   4935
   StartUpPosition =   2  '�ù�����
   Begin VB.Frame Frame4 
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      Caption         =   "����"
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   2160
      TabIndex        =   11
      Top             =   1680
      Width           =   2655
      Begin VB.CommandButton Command3 
         Caption         =   "����(&A)"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.TextBox TxtTempForLog 
      Appearance      =   0  '����
      Height          =   270
      Left            =   4080
      TabIndex        =   10
      Text            =   "Text2"
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      Caption         =   "�Ϻо��Ыت��A(&I)"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   2160
      TabIndex        =   8
      Top             =   960
      Width           =   2655
      Begin VB.Label DiskInfo 
         Alignment       =   2  '�m�����
         Appearance      =   0  '����
         BackColor       =   &H80000005&
         BorderStyle     =   1  '��u�T�w
         Caption         =   "�L�k�d��"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "�I���H���s���J"
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.TextBox TextForCheckSpace 
      Appearance      =   0  '����
      Height          =   270
      Left            =   4320
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      Caption         =   "�ѱK(&E)"
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4695
      Begin VB.CommandButton Command1 
         Caption         =   "�ѱK(&R)"
         Height          =   375
         Left            =   3720
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '����
         BeginProperty Font 
            Name            =   "YaHei Consolas Hybrid"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  '�Ȥ�
         Left            =   1200
         PasswordChar    =   "?"
         TabIndex        =   3
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label1 
         Appearance      =   0  '����
         BackColor       =   &H80000005&
         Caption         =   "�K�X(&P)�G"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "���B�ݿ�J�K�X�ӫD���ͮɪ��K�_�I"
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      Caption         =   "��L�ʧ@(&D)"
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   1935
      Begin VB.CommandButton Command2 
         Caption         =   "�ק�K�X(&C)"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton MkPwdTxt 
         Caption         =   "���ͷs�Ϻо�(&P)"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "���ͱa�K�X���Ϻо�"
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Menu key 
      Caption         =   "�K�_(&K)"
      Begin VB.Menu MakeKey 
         Caption         =   "���ͱK�_(&M)"
         Shortcut        =   ^M
      End
      Begin VB.Menu ChangePWD 
         Caption         =   "���K�X(&C)"
      End
      Begin VB.Menu dash 
         Caption         =   "-"
      End
      Begin VB.Menu Return 
         Caption         =   "���sŪ���K�_(&R)"
         Shortcut        =   {F5}
      End
      Begin VB.Menu ForgotPWD 
         Caption         =   "�ѰO�K�X(&F)"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu Helper 
      Caption         =   "����(&H)"
      Begin VB.Menu LogFile 
         Caption         =   "Log�O����(&L)"
      End
      Begin VB.Menu dash1 
         Caption         =   "-"
      End
      Begin VB.Menu About 
         Caption         =   "����(&A)"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fso As FileSystemObject
Dim UserPWD As String
Option Explicit
Dim ExtractInfo As Boolean
Dim MainWidthAndHeight(1) As Long
Dim pwdIsNA As Boolean

Private Sub About_Click()
    frmAbout.Show
End Sub

Private Sub ChangePWD_Click()
    If ExtractInfo Then
        Dim OldPWD As String
        OldPWD = InputBox("��J�±K�X", "���K�X")
        If OldPWD = UserPWD Then
            Dim newPWD, checkNewPWD, userKey
            newPWD = InputBox("��J�s�K�X", "���K�X")
            checkNewPWD = InputBox("�T�{�s�K�X", "���K�X")
            If newPWD = checkNewPWD Then
                MsgBox "�ק粒���I", vbInformation
                UserPWD = newPWD
                newPWD = ""
                checkNewPWD = ""
                OldPWD = ""
                userKey = AscCodePassWord(UserPWD)
                InputBox "�ƻs�z���᪺�s�K�X�öK�W��pwd.txt�ɮפ�", "���K�X", UserPWD
                MsgBox "���\���ͱK�_�I" & vbNewLine & "�бN�ݷ|���X���T�����K�_����r�߶K�쥻�n��ڥؿ��upwd.txt�v���A�p���@�ӡA" & App.Title _
                & "�~�i�H���`�s���øѱK�A�ö}�ұz���K�X�C" & vbNewLine & "���¡C" & vbNewLine & vbNewLine & "�Ƶ��G�p�G�ڥؿ��U�S��pwd.txt�A�Цۦ�ЫبöK�W�K�_�C", vbInformation
            End If
        End If
    Else
        MsgBox "���~�I" & vbNewLine & "�z�|���ѱK�Ϋإߵ����ϺСA�ѱK�Ϋإߵ����ϺЫ�A�ոաC", vbCritical
    End If
End Sub

Function ReadPWD()
Retry:
    LogWrite "ReadPWD"
    Dim Str1 As String
    TextForCheckSpace = ""
    Open App.Path + "\pwd.txt" For Input As #1
    On Error GoTo fileIsSpace
    Line Input #1, Str1
    TextForCheckSpace.Text = Str1
    If TextForCheckSpace.Text = "" Or TextForCheckSpace.Text = " " Then
        MsgBox "��󬰪šC", vbInformation
        pwdIsNA = True
        Close #1
    Else
        On Error GoTo Error
        Const ForReading = 1
        Dim fid As TextStream
        Set fso = CreateObject("Scripting.FileSystemObject")
        LogWrite "Dim Vars:" & vbNewLine & "fso As FileSystemObject" & vbNewLine & "fid As TextStream" & vbNewLine & "fso = CreateObject('Scripting.FileSystemObject')"
        Set fid = fso.OpenTextFile(App.Path + "\pwd.txt", ForReading)
        ReadPWD = fid.ReadLine
        fid.Close
        Close #1
        pwdIsNA = False
    End If
    DiskInfo.Caption = pwdIsNA
    If pwdIsNA Then
        DiskInfo.BackColor = QBColor(10)
        DiskInfo.ForeColor = QBColor(15)
    Else
        DiskInfo.BackColor = QBColor(12)
        DiskInfo.ForeColor = QBColor(15)
    End If
    Exit Function
    Close #1
Error:
    Dim MsgBoxReturnValue
    LogWrite "ReadPWD:Not Found pwd.txt", 1
    MsgBoxReturnValue = MsgBox("���~�I�L�kŪ���gASCII�[�K�ᤧ�K�X�奻�A�Х��i��K�X�奻�[�K�ýT�w�x�s��" + App.Path + "\pwd.txt��A��Ū��", vbCritical + vbAbortRetryIgnore)
    Select Case MsgBoxReturnValue
        Case vbAbort
            End
            LogWrite "ReadPWD-Ans:Abort", 1
        Case vbRetry
            GoTo Retry
            LogWrite "ReadPWD-Ans:Retry", 1
        Case vbIgnore
            LogWrite "ReadPWD-Ans:Ignore", 1
            Exit Function
    End Select
    LogWrite "Exit Select", 1
    Exit Function
fileIsSpace:
    pwdIsNA = True
    LogWrite "Text is Space"
End Function

Private Sub Command1_Click()
    'On Error Resume Next
    If Not pwdIsNA Then
        LogWrite pwdIsNA
        If Text1.Text = UserPWD Then
            Shell "cmd.exe /c start " & App.Path & "\About.bat", vbNormalFocus
            LogWrite "Shell 'cmd.exe /c start ' & app.path & ' \ About.bat ', vbNormalFocus"
            MsgBox "�����I", vbInformation
            ExtractInfo = True
        Else
            MsgBox "�K�X�����T�I", vbCritical
        End If
    Else
        Dim MsgBoxClickVal
        MsgBoxClickVal = MsgBox("ĵ�i�I" & vbNewLine & "�|���إߵ����Ϻо��αK�_�ɤw�l�a�A�O�_���إߵ����Ϻо��H", vbExclamation + vbYesNoCancel)
        Select Case MsgBoxClickVal
            Case vbYes
                MkPwdTxt_Click
            Case Else
                Exit Sub
        End Select
        LogWrite pwdIsNA
    End If
End Sub

Private Sub Command2_Click()
    ChangePWD_Click
End Sub

Private Sub Command3_Click()
    frmAbout.Show
End Sub

Private Sub DiskInfo_Click()
    ReadPWD
End Sub

Private Sub ForgotPWD_Click()
    Dim RecoverPWD
    RecoverPWD = MsgBox("��p�I" & vbNewLine & "���F��T�w���A�ڭ̵L�k���ұz�������w�о֦��̡A�z�O�_�@�N�R�������w�Шí��s�إߡH", vbYesNo + vbExclamation)
    LogWrite "'��p�I' & vbNewLine & '���F��T�w���A�ڭ̵L�k���ұz�������w�о֦��̡A�z�O�_�@�N�R�������w�Шí��s�إߡH', vbYesNo + vbExclamation"
    Select Case RecoverPWD
        Case 7
            LogWrite RecoverPWD, 3
            Exit Sub
        Case Else
            LogWrite RecoverPWD, 3
    End Select
    RecoverPWD = MsgBox("�Ъ`�N�I" & vbNewLine & "���s�إߵ����w�Ы�A��w�Ф�����T�N�|�Q�����I" & vbNewLine & "�O�_�n���s�w�СH", vbYesNo + vbExclamation)
    Select Case RecoverPWD
        Case 7
            Exit Sub
        Case Else
            LogWrite RecoverPWD, 3
    End Select
    Shell "cmd.exe /c " & "rmdir /S /Q D:\RECYCLED\UDrives", vbNormalFocus
    MsgBox "���~�I" & vbNewLine & "�L�k���s�}�l�s�w�СA�]�ӥ������ʧ@", vbCritical
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub LogFile_Click()
    Logfrm.Show
End Sub

Private Sub MakeKey_Click()
    MkPwdTxt_Click
End Sub

Private Sub MkPwdTxt_Click()
    If pwdIsNA Then
        frmLogin.Show
    Else
        MsgBox "���~�I" & vbNewLine & "���q���w�ЫعL�����Ϻо��A�T���", vbCritical
    End If
End Sub

Private Sub Form_Load()
    Me.Show
    Logfrm.LogLbl.Text = "[LogWrite List]"
    UserPWD = ReadPWD()
    If UserPWD = "" Then
        UserPWD = "N/A"
    End If
    LogWrite "ReadPWD=" & UserPWD, 3
    UserPWD = ChAscCodePassWord(UserPWD)
    Debug.Print UserPWD
    ExtractInfo = False
    MainWidthAndHeight(0) = Me.Width
    MainWidthAndHeight(1) = Me.Height
End Sub

Private Sub Picture1_Click()
    Logfrm.LogLbl.Text = "[LogWrite List]"
End Sub

Private Sub Picture2_Click()
    TxtSave
End Sub

Private Sub Return_Click()
    Form_Load
End Sub
