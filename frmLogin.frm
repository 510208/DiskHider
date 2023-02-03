VERSION 5.00
Begin VB.Form frmLogin 
   Appearance      =   0  '平面
   BackColor       =   &H80000005&
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "登入"
   ClientHeight    =   1545
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   6510
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   6112.537
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.TextBox Text1 
      Appearance      =   0  '平面
      BorderStyle     =   0  '沒有框線
      Height          =   855
      Left            =   3720
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "frmLogin.frx":10CA
      Top             =   120
      Width           =   2655
   End
   Begin VB.TextBox KeyTxt 
      Appearance      =   0  '平面
      Height          =   345
      IMEMode         =   3  '暫止
      Left            =   1320
      TabIndex        =   4
      Top             =   600
      Width           =   1845
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  '平面
      Caption         =   "確定"
      Default         =   -1  'True
      Height          =   390
      Left            =   480
      TabIndex        =   2
      Top             =   1080
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  '平面
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   390
      Left            =   2220
      TabIndex        =   3
      Top             =   1080
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  '平面
      Height          =   345
      IMEMode         =   3  '暫止
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
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      Caption         =   "密鑰(&K):"
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      Caption         =   "密碼(&P):"
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

'
'                        _oo0oo_
'                       o8888888o
'                       88" . "88
'                       (| -_- |)
'                       0\  =  /0
'                     ___/`---'\___
'                   .' \\|     |// '.
'                  / \\|||  :  |||// \
'                 / _||||| -:- |||||- \
'                |   | \\\  - /// |   |
'                | \_|  ''\---/''  |_/ |
'                \  .-\__  '-'  ___/-. /
'              ___'. .'  /--.--\  `. .'___
'           ."" '<  `.___\_<|>_/___.' >' "".
'          | | :  `- \`.;`\ _ /`;.`/ - ` : | |
'          \  \ `_.   \_ __\ /__ _/   .-` /  /
'      =====`-.____`.___ \_____/___.-`___.-'=====
'                        `=---='
' 
' 
'      ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' 
'            佛祖保佑       永不當機     永無BUG
' 
'        佛曰:  
'                寫字樓里寫字間，寫字間里程式員；  
'                程式人員寫程式，又拿程式換酒錢。  
'                酒醒只在網上坐，酒醉還來網下眠；  
'                酒醉酒醒日復日，網上網下年復年。  
'                但願老死電腦間，不願鞠躬老闆前；  
'                奔馳寶馬貴者趣，公交自行程式員。  
'                別人笑我忒瘋癲，我笑自己命太賤；  
'                不見滿街漂亮妹，哪個歸得程式員？
'

Sub mkpwd()
    Dim PWDAsc As String
    PWDAsc = AscCodePassWord(txtPassword.Text)
    KeyTxt.Text = PWDAsc
    MsgBox "成功產生密鑰！" & vbNewLine & "請將「密鑰」文本框中的文字粘貼到本軟體根目錄「pwd.txt」中，如此一來，" & App.Title _
    & "才可以正常存取並解密，並開啟您的密碼。" & vbNewLine & "謝謝。" & vbNewLine & vbNewLine & "備註：如果根目錄下沒有pwd.txt，請自行創建並貼上密鑰。", vbInformation
End Sub

Private Sub cmdCancel_Click()
    '設定全域變數為 false 來代表
    '失敗的登入
    LogWriteinSucceeded = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    LogWriteinSucceeded = True
    If txtPassword.Text = "" Then
        MsgBox "錯誤！你還未輸入密碼！", vbCritical
        LogWrite "cmdOK_Click:PWD N/A"
        txtPassword.Text = ""
        txtPassword.SetFocus
    Else
        mkpwd
    End If
End Sub

Private Sub Image1_Click()
    Clipboard.SetText KeyTxt.Text
    MsgBox "完成複製！", vbInformation
    frmLogin.Hide
End Sub
