VERSION 5.00
Begin VB.Form frmLogin 
   Appearance      =   0  '����
   BackColor       =   &H80000005&
   BorderStyle     =   3  '���u�T�w��ܤ��
   Caption         =   "�n�J"
   ClientHeight    =   1545
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   6510
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  '�ϥΪ̦ۭq
   ScaleWidth      =   6112.537
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '�ù�����
   Begin VB.TextBox Text1 
      Appearance      =   0  '����
      BorderStyle     =   0  '�S���ؽu
      Height          =   855
      Left            =   3720
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "frmLogin.frx":10CA
      Top             =   120
      Width           =   2655
   End
   Begin VB.TextBox KeyTxt 
      Appearance      =   0  '����
      Height          =   345
      IMEMode         =   3  '�Ȥ�
      Left            =   1320
      TabIndex        =   4
      Top             =   600
      Width           =   1845
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  '����
      Caption         =   "�T�w"
      Default         =   -1  'True
      Height          =   390
      Left            =   480
      TabIndex        =   2
      Top             =   1080
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  '����
      Cancel          =   -1  'True
      Caption         =   "����"
      Height          =   390
      Left            =   2220
      TabIndex        =   3
      Top             =   1080
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  '����
      Height          =   345
      IMEMode         =   3  '�Ȥ�
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   165
      Width           =   2325
   End
   Begin VB.Image Image1 
      Height          =   330
      Left            =   3240
      Picture         =   "frmLogin.frx":1121
      Top             =   600
      Width           =   330
   End
   Begin VB.Label lblLabels 
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      Caption         =   "�K�_(&K):"
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      Caption         =   "�K�X(&P):"
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   0
      Top             =   180
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fso As FileSystemObject
Public LogWriteinSucceeded As Boolean

Sub mkpwd()
    Dim PWDAsc As String
    PWDAsc = AscCodePassWord(txtPassword.Text)
    KeyTxt.Text = PWDAsc
    MsgBox "���\���ͱK�_�I" & vbNewLine & "�бN�u�K�_�v�奻�ؤ�����r�߶K�쥻�n��ڥؿ��upwd.txt�v���A�p���@�ӡA" & App.Title _
    & "�~�i�H���`�s���øѱK�A�ö}�ұz���K�X�C" & vbNewLine & "���¡C" & vbNewLine & vbNewLine & "�Ƶ��G�p�G�ڥؿ��U�S��pwd.txt�A�Цۦ�ЫبöK�W�K�_�C", vbInformation
End Sub

Private Sub cmdCancel_Click()
    '�]�w�����ܼƬ� false �ӥN��
    '���Ѫ��n�J
    LogWriteinSucceeded = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    LogWriteinSucceeded = True
    If txtPassword.Text = "" Then
        MsgBox "���~�I�A�٥���J�K�X�I", vbCritical
        LogWrite "cmdOK_Click:PWD N/A"
        txtPassword.Text = ""
        txtPassword.SetFocus
    Else
        mkpwd
    End If
End Sub

Private Sub Image1_Click()
    Clipboard.SetText KeyTxt.Text
    MsgBox "�����ƻs�I", vbInformation
    frmLogin.Hide
End Sub
