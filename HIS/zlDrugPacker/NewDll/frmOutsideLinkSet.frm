VERSION 5.00
Begin VB.Form frmOutsideLinkSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "药品分包机数据库设置"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5415
   Icon            =   "frmOutsideLinkSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtDBName 
      Height          =   270
      Left            =   1680
      TabIndex        =   3
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "测试(&T)"
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "连接(&S)"
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   720
      Width           =   1095
   End
   Begin VB.Frame fraLine 
      Caption         =   "Frame1"
      Height          =   2175
      Left            =   3840
      TabIndex        =   11
      Top             =   -240
      Width           =   38
   End
   Begin VB.TextBox txtServer 
      Height          =   270
      Left            =   1680
      TabIndex        =   2
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox txtPWD 
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox txtUser 
      Height          =   270
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label lblServer 
      AutoSize        =   -1  'True
      Caption         =   "服务器(&E)"
      Height          =   180
      Left            =   600
      TabIndex        =   10
      Top             =   960
      Width           =   810
   End
   Begin VB.Label lblPWD 
      AutoSize        =   -1  'True
      Caption         =   "密  码(&P)"
      Height          =   180
      Left            =   600
      TabIndex        =   9
      Top             =   600
      Width           =   810
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      Caption         =   "用  户(&U)"
      Height          =   180
      Left            =   600
      TabIndex        =   8
      Top             =   240
      Width           =   810
   End
   Begin VB.Label lblType 
      AutoSize        =   -1  'True
      Caption         =   "数据库(&N)"
      Height          =   180
      Left            =   600
      TabIndex        =   7
      Top             =   1320
      Width           =   810
   End
End
Attribute VB_Name = "frmOutsideLinkSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public gblnSetupFinish As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim strUser As String, strServer As String, strPWD As String, strDBName As String
    Dim blnConnect As Boolean
    
    strUser = Trim(txtUser.Text)
    strServer = Trim(txtServer.Text)
    strPWD = Trim(txtPWD.Text)
    strDBName = Trim(txtDBName.Text)
    gblnSetupFinish = False
    
    '数据审核
    If Len(strUser) = 0 Then
        MsgBox "请输入用户信息！", vbInformation, GSTR_MESSAGE
        Exit Sub
    End If
    If Len(strServer) = 0 Then
        MsgBox "请输入外联服务器信息！", vbInformation, GSTR_MESSAGE
        Exit Sub
    End If
    '连接
    Screen.MousePointer = vbHourglass
    blnConnect = MSSQLServerOpen(strServer, strDBName, strUser, strPWD)
    Screen.MousePointer = vbDefault
    If blnConnect = True Then
        '加密密码
        strPWD = StringEnDeCodecn(strPWD, 68)
        '保存设置
        SaveSetting "ZLSOFT", GSTR_REGEDIT_PATH, "SERVER", strServer
        SaveSetting "ZLSOFT", GSTR_REGEDIT_PATH, "DBNAME", strDBName
        SaveSetting "ZLSOFT", GSTR_REGEDIT_PATH, "USER", strUser
        SaveSetting "ZLSOFT", GSTR_REGEDIT_PATH, "PASSWORD", strPWD
        gblnSetupFinish = True
        Unload Me
    End If
End Sub

Private Sub cmdTest_Click()
    Dim blnTest As Boolean
    Screen.MousePointer = vbHourglass
    blnTest = MSSQLServerOpen(txtServer.Text, txtDBName.Text, txtUser.Text, txtPWD.Text)
    Screen.MousePointer = vbDefault
    If blnTest Then
        gcnOutside.Close
        MsgBox "测试连接成功！", vbInformation, GSTR_MESSAGE
    End If
    Set gcnOutside = Nothing
End Sub

Private Sub Form_Activate()
    txtUser.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{Tab}"
    End If
End Sub

Private Sub Form_Load()
    Dim strPWD As String, strUser As String, strServer As String, strDBName As String
    '初始化
    strUser = GetSetting(appName:="ZLSOFT", Section:=GSTR_REGEDIT_PATH, Key:="USER", Default:="")
    If Trim(strUser) = "" Then
        strServer = MSTR_SERVER
        strUser = MSTR_USER
        strDBName = MSTR_DBNAME
        strPWD = MSTR_PASSWORD
    Else
        strServer = GetSetting(appName:="ZLSOFT", Section:=GSTR_REGEDIT_PATH, Key:="SERVER", Default:="")
        strDBName = GetSetting(appName:="ZLSOFT", Section:=GSTR_REGEDIT_PATH, Key:="DBNAME", Default:="")
        strPWD = GetSetting(appName:="ZLSOFT", Section:=GSTR_REGEDIT_PATH, Key:="PASSWORD", Default:="")
        strPWD = StringEnDeCodecn(strPWD, 68)       '解密密码
    End If
    txtPWD.Text = strPWD
    txtUser.Text = strUser
    txtServer.Text = strServer
    txtDBName.Text = strDBName
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    If Not gcnOutside Is Nothing Then
'        If gcnOutside.State = 1 Then gcnOutside.Close
'    End If
End Sub

Private Sub txtDBName_GotFocus()
    SelText txtDBName
End Sub

Private Sub txtPWD_GotFocus()
    SelText txtPWD
End Sub

Private Sub txtServer_GotFocus()
    SelText txtServer
End Sub

Private Sub txtUser_GotFocus()
    SelText txtUser
End Sub

