VERSION 5.00
Begin VB.Form Main 
   Appearance      =   0  '����
   BackColor       =   &H80000005&
   BorderStyle     =   1  '��u�T�w
   Caption         =   "DiskHider"
   ClientHeight    =   3015
   ClientLeft      =   14040
   ClientTop       =   2925
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   6960
   Begin VB.Frame Frame2 
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      Caption         =   "�ѱK(&E)"
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4215
      Begin VB.CommandButton Command1 
         Caption         =   "�ѱK(&R)"
         Height          =   375
         Left            =   3240
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '����
         Height          =   270
         Left            =   1200
         TabIndex        =   4
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label1 
         Appearance      =   0  '����
         BackColor       =   &H80000005&
         Caption         =   "�K�X(&P)�G"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   6
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
      Height          =   1215
      Left            =   4440
      TabIndex        =   1
      Top             =   120
      Width           =   2415
      Begin VB.CommandButton Command2 
         Caption         =   "�ק�K�X(&C)"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   2175
      End
      Begin VB.CommandButton MkPwdTxt 
         Caption         =   "���ͱa�K�X���Ϻо�(&P)"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.TextBox LogLbl 
      Appearance      =   0  '����
      BeginProperty Font 
         Name            =   "YaHei Consolas Hybrid"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  '�������b
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "Main.frx":0000
      Top             =   960
      Width           =   4215
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

Function ReadPWD()
Retry:
    LogWrite "ReadPWD"
On Error GoTo Error
    Const ForReading = 1
    Dim fso As FileSystemObject
    Dim fid As TextStream
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fid = fso.OpenTextFile(App.Path + "\pwd.txt", ForReading)
    ReadPWD = fid.ReadLine
    fid.Close
    Exit Function
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
End Function

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
        MsgBox "���~�I" & vbNewLine & "�z�|���ѱK�A�ѱK��A�ոաC", vbCritical
    End If
End Sub

Private Sub Command1_Click()
    On Error Resume Next
    If Text1.Text = UserPWD Then
        Shell "cmd.exe /c start " & App.Path & "\About.bat", vbNormalFocus
        LogWrite "Shell 'cmd.exe /c start ' & app.path & ' \ About.bat ', vbNormalFocus"
        MsgBox "�����I", vbInformation
        ExtractInfo = True
    Else
        MsgBox "�K�X�����T�I", vbCritical
    End If
End Sub

Private Sub Command2_Click()
    ChangePWD_Click
End Sub

Private Sub ForgotPWD_Click()
    Dim RecoverPWD
    RecoverPWD = MsgBox("��p�I" & vbNewLine & "���F��T�w���A�ڭ̵L�k���ұz�������w�о֦��̡A�z�O�_�@�N�R�������w�Шí��s�إߡH", vbYesNo + vbExclamation)
    Select Case RecoverPWD
        Case 7
            Exit Sub
    End Select
    RecoverPWD = MsgBox("�Ъ`�N�I" & vbNewLine & "���s�إߵ����w�Ы�A��w�Ф�����T�N�|�Q�����I" & vbNewLine & "�O�_�n���s�w�СH", vbYesNo + vbExclamation)
    Select Case RecoverPWD
        Case 7
            Exit Sub
    End Select
    Shell "cmd.exe /c " & "rmdir /S /Q D:\RECYCLED\UDrives", vbNormalFocus
    MsgBox "���~�I" & vbNewLine & "�L�k���s�}�l�s�w�СA�]�ӥ������ʧ@", vbCritical
End Sub

Private Sub MakeKey_Click()
    MkPwdTxt_Click
End Sub

Private Sub MkPwdTxt_Click()
    frmLogin.Show
End Sub

Private Sub Form_Load()
    Me.Show
    LogLbl.Text = "[LogWrite List]"
    UserPWD = ReadPWD()
    If UserPWD = "" Then
        UserPWD = "N/A"
    End If
    LogWrite "ReadPWD=" & UserPWD, 3
    UserPWD = ChAscCodePassWord(UserPWD)
    Debug.Print UserPWD
    ExtractInfo = False
End Sub

Private Sub Return_Click()
    Form_Load
End Sub
