VERSION 5.00
Begin VB.Form LogFrm 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   5835
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.PictureBox Picture2 
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4680
      Picture         =   "LogFrm.frx":0000
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   2
      ToolTipText     =   "�M�Ŭ�����r(&C)"
      Top             =   360
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4680
      Picture         =   "LogFrm.frx":09AA
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   1
      ToolTipText     =   "�M�Ŭ�����r(&C)"
      Top             =   0
      Width           =   375
   End
   Begin VB.TextBox LogLbl 
      Appearance      =   0  '����
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  '�������b
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "LogFrm.frx":1354
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "LogFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LogLbl_Change()
    TxtSave
End Sub