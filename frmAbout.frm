VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  '���u�T�w��ܤ��
   Caption         =   "����ڪ����ε{��"
   ClientHeight    =   3555
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2453.724
   ScaleMode       =   0  '�ϥΪ̦ۭq
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "frmAbout.frx":10CA
      ScaleHeight     =   337.12
      ScaleMode       =   0  '�ϥΪ̦ۭq
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "�T�w"
      Default         =   -1  'True
      Height          =   345
      Left            =   4245
      TabIndex        =   0
      Top             =   2625
      Width           =   1260
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "�t�θ�T(&S)..."
      Height          =   345
      Left            =   4260
      TabIndex        =   2
      Top             =   3075
      Width           =   1245
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '����u
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1687.583
      Y2              =   1687.583
   End
   Begin VB.Label lblDescription 
      Caption         =   "���ε{������"
      ForeColor       =   &H00000000&
      Height          =   1170
      Left            =   1050
      TabIndex        =   3
      Top             =   1125
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "���ε{�����D"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1050
      TabIndex        =   5
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1697.936
      Y2              =   1697.936
   End
   Begin VB.Label lblVersion 
      Caption         =   "����"
      Height          =   225
      Left            =   1050
      TabIndex        =   6
      Top             =   780
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "ĵ�i: ..."
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   255
      TabIndex        =   4
      Top             =   2625
      Width           =   3870
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ���U���X�w���ʿﶵ...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' ���U���X ROOT ����...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' �H Unicode nul ���������r��
Const REG_DWORD = 4                      ' 32-�줸�ƭ�

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "���� " & App.Title
    lblVersion.Caption = "���� " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    picIcon.Picture = Main.Icon
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' ���ձq���U�Ϩ��o�t�θ�T�{�����|\�W��...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' ���ձq���U�Ϩ��o�t�θ�T�{�����|...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' �ˬd�w���� 32 �줸�ɮת����O�_�s�b
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' ���~ - �䤣���ɮ�...
        Else
            GoTo SysInfoErr
        End If
    ' ���~ - �䤣����U����...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "�ثe�L�k���Ѩt�θ�T", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' �j��p�ƾ�
    Dim rc As Long                                          ' �Ǧ^�N�X
    Dim hKey As Long                                        ' �}�Ҫ����U���X������N�X
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' ���U���X����ƫ��A
    Dim tmpVal As String                                    ' ���U���X�Ȫ��Ȧs�Ŷ�
    Dim KeyValSize As Long                                  ' ���U���X�ܼƪ��j�p
    '------------------------------------------------------------
    ' �}�� KeyRoot {HKEY_LOCAL_MACHINE...} ���U�����U���X (RegKey)
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' �}�ҵ��U���X
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' �B�z���~...
    
    tmpVal = String$(1024, 0)                               ' �t�m�ܼƪŶ�
    KeyValSize = 1024                                       ' �Х��ܼƤj�p
    
    '------------------------------------------------------------
    ' �^�����U���X��...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' ���o/�إ߾��X��
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' �B�z���~
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 �|�[�J�H Null ���������r��...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' ��� Null�A�q�r�ꤤ���X
    Else                                                    ' WinNT ���|�[�J�H Null ���������r��...
        tmpVal = Left(tmpVal, KeyValSize)                   ' �䤣�� Null�A���X�r��
    End If
    '------------------------------------------------------------
    ' �M�w���X�Ȫ��ഫ���A...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' �j�M��ƫ��A...
    Case REG_SZ                                             ' String ���U���X��ƫ��A
        KeyVal = tmpVal                                     ' �ƻs�r���
    Case REG_DWORD                                          ' Double Word ���U���X��ƫ��A
        For i = Len(tmpVal) To 1 Step -1                    ' �ഫ�C�@�Ӧ줸
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' �v�r�إ߭�
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' �N Double Word �ഫ�� String
    End Select
    
    GetKeyValue = True                                      ' �Ǧ^���\���T��
    rc = RegCloseKey(hKey)                                  ' �������U���X
    Exit Function                                           ' ���}
    
GetKeyError:      ' ���~�o�ͫ�M��...
    KeyVal = ""                                             ' �]�w�Ǧ^�Ȭ��Ŧr��
    GetKeyValue = False                                     ' �Ǧ^���Ѫ��T��
    rc = RegCloseKey(hKey)                                  ' �������U���X
End Function
