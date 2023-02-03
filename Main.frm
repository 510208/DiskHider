VERSION 5.00
Begin VB.Form Main 
   Appearance      =   0  '平面
   BackColor       =   &H80000005&
   BorderStyle     =   1  '單線固定
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
   StartUpPosition =   2  '螢幕中央
   Begin VB.Frame Frame4 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      Caption         =   "關於"
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   2160
      TabIndex        =   11
      Top             =   1680
      Width           =   2655
      Begin VB.CommandButton Command3 
         Caption         =   "關於(&A)"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.TextBox TxtTempForLog 
      Appearance      =   0  '平面
      Height          =   270
      Left            =   4080
      TabIndex        =   10
      Text            =   "Text2"
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      Caption         =   "磁碟機創建狀態(&I)"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   2160
      TabIndex        =   8
      Top             =   960
      Width           =   2655
      Begin VB.Label DiskInfo 
         Alignment       =   2  '置中對齊
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         BorderStyle     =   1  '單線固定
         Caption         =   "無法查詢"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "點擊以重新載入"
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.TextBox TextForCheckSpace 
      Appearance      =   0  '平面
      Height          =   270
      Left            =   4320
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      Caption         =   "解密(&E)"
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4695
      Begin VB.CommandButton Command1 
         Caption         =   "解密(&R)"
         Height          =   375
         Left            =   3720
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '平面
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
         IMEMode         =   3  '暫止
         Left            =   1200
         PasswordChar    =   "?"
         TabIndex        =   3
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label1 
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         Caption         =   "密碼(&P)："
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "此處需輸入密碼而非產生時的密鑰！"
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      Caption         =   "其他動作(&D)"
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   1935
      Begin VB.CommandButton Command2 
         Caption         =   "修改密碼(&C)"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton MkPwdTxt 
         Caption         =   "產生新磁碟機(&P)"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "產生帶密碼的磁碟機"
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Menu key 
      Caption         =   "密鑰(&K)"
      Begin VB.Menu MakeKey 
         Caption         =   "產生密鑰(&M)"
         Shortcut        =   ^M
      End
      Begin VB.Menu ChangePWD 
         Caption         =   "更改密碼(&C)"
      End
      Begin VB.Menu dash 
         Caption         =   "-"
      End
      Begin VB.Menu Return 
         Caption         =   "重新讀取密鑰(&R)"
         Shortcut        =   {F5}
      End
      Begin VB.Menu ForgotPWD 
         Caption         =   "忘記密碼(&F)"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu Helper 
      Caption         =   "說明(&H)"
      Begin VB.Menu LogFile 
         Caption         =   "Log記錄檔(&L)"
      End
      Begin VB.Menu dash1 
         Caption         =   "-"
      End
      Begin VB.Menu About 
         Caption         =   "關於(&A)"
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
        OldPWD = InputBox("輸入舊密碼", "更改密碼")
        If OldPWD = UserPWD Then
            Dim newPWD, checkNewPWD, userKey
            newPWD = InputBox("輸入新密碼", "更改密碼")
            checkNewPWD = InputBox("確認新密碼", "更改密碼")
            If newPWD = checkNewPWD Then
                MsgBox "修改完成！", vbInformation
                UserPWD = newPWD
                newPWD = ""
                checkNewPWD = ""
                OldPWD = ""
                userKey = AscCodePassWord(UserPWD)
                InputBox "複製您更改後的新密碼並貼上至pwd.txt檔案內", "更改密碼", UserPWD
                MsgBox "成功產生密鑰！" & vbNewLine & "請將待會跳出的訊息中密鑰的文字粘貼到本軟體根目錄「pwd.txt」中，如此一來，" & App.Title _
                & "才可以正常存取並解密，並開啟您的密碼。" & vbNewLine & "謝謝。" & vbNewLine & vbNewLine & "備註：如果根目錄下沒有pwd.txt，請自行創建並貼上密鑰。", vbInformation
            End If
        End If
    Else
        MsgBox "錯誤！" & vbNewLine & "您尚未解密或建立虛擬磁碟，解密或建立虛擬磁碟後再試試。", vbCritical
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
        MsgBox "文件為空。", vbInformation
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
    MsgBoxReturnValue = MsgBox("錯誤！無法讀取經ASCII加密後之密碼文本，請先進行密碼文本加密並確定儲存於" + App.Path + "\pwd.txt後再行讀檔", vbCritical + vbAbortRetryIgnore)
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
            MsgBox "完成！", vbInformation
            ExtractInfo = True
        Else
            MsgBox "密碼不正確！", vbCritical
        End If
    Else
        Dim MsgBoxClickVal
        MsgBoxClickVal = MsgBox("警告！" & vbNewLine & "尚未建立虛擬磁碟機或密鑰檔已損壞，是否重建立虛擬磁碟機？", vbExclamation + vbYesNoCancel)
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
    RecoverPWD = MsgBox("抱歉！" & vbNewLine & "為了資訊安全，我們無法驗證您為虛擬硬碟擁有者，您是否願意刪除虛擬硬碟並重新建立？", vbYesNo + vbExclamation)
    LogWrite "'抱歉！' & vbNewLine & '為了資訊安全，我們無法驗證您為虛擬硬碟擁有者，您是否願意刪除虛擬硬碟並重新建立？', vbYesNo + vbExclamation"
    Select Case RecoverPWD
        Case 7
            LogWrite RecoverPWD, 3
            Exit Sub
        Case Else
            LogWrite RecoverPWD, 3
    End Select
    RecoverPWD = MsgBox("請注意！" & vbNewLine & "重新建立虛擬硬碟後，原硬碟內之資訊將會被移除！" & vbNewLine & "是否要重製硬碟？", vbYesNo + vbExclamation)
    Select Case RecoverPWD
        Case 7
            Exit Sub
        Case Else
            LogWrite RecoverPWD, 3
    End Select
    Shell "cmd.exe /c " & "rmdir /S /Q D:\RECYCLED\UDrives", vbNormalFocus
    MsgBox "錯誤！" & vbNewLine & "無法重新開始新硬碟，因而未完成動作", vbCritical
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
        MsgBox "錯誤！" & vbNewLine & "此電腦已創建過虛擬磁碟機，禁止重創", vbCritical
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
