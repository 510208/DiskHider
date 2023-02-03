VERSION 5.00
Begin VB.Form LogFrm 
   Appearance      =   0  '平面
   BackColor       =   &H80000005&
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8685
   Icon            =   "LogFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   8685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '系統預設值
   Begin VB.PictureBox Picture2 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8280
      Picture         =   "LogFrm.frx":10CA
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   2
      ToolTipText     =   "清空紀錄文字(&C)"
      Top             =   360
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8280
      Picture         =   "LogFrm.frx":1A74
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   1
      ToolTipText     =   "清空紀錄文字(&C)"
      Top             =   0
      Width           =   375
   End
   Begin VB.TextBox LogLbl 
      Appearance      =   0  '平面
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
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "LogFrm.frx":241E
      Top             =   0
      Width           =   8295
   End
End
Attribute VB_Name = "LogFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Width = LogLbl.Width + Picture1.Width
End Sub

Private Sub LogLbl_Change()
    TxtSave
End Sub
