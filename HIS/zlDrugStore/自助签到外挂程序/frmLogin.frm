VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "用户登录"
   ClientHeight    =   2640
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4320
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2640
   ScaleWidth      =   4320
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   1800
      TabIndex        =   10
      Top             =   2160
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3060
      TabIndex        =   9
      Top             =   2160
      Width           =   1100
   End
   Begin VB.Frame Frame2 
      Height          =   120
      Left            =   0
      TabIndex        =   8
      Top             =   1920
      Width           =   4965
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   0
      TabIndex        =   7
      Top             =   480
      Width           =   4965
   End
   Begin VB.TextBox txtUser 
      Height          =   300
      Left            =   2055
      TabIndex        =   1
      ToolTipText     =   "系统所有者"
      Top             =   720
      Width           =   1890
   End
   Begin VB.TextBox txtPass 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   2055
      PasswordChar    =   "*"
      TabIndex        =   3
      ToolTipText     =   "未转换的系统所有者密码"
      Top             =   1140
      Width           =   1890
   End
   Begin VB.TextBox txtSvr 
      Height          =   300
      Left            =   2055
      TabIndex        =   4
      Top             =   1575
      Width           =   1890
   End
   Begin VB.Image imgFlag 
      Height          =   720
      Left            =   360
      Picture         =   "frmLogin.frx":030A
      Top             =   840
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "请使用数据库用户名和密码进行登录"
      Height          =   180
      Left            =   1200
      TabIndex        =   6
      Top             =   240
      Width           =   2880
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   480
      Picture         =   "frmLogin.frx":0994
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "用户名"
      Height          =   180
      Left            =   1440
      TabIndex        =   5
      Top             =   780
      Width           =   540
   End
   Begin VB.Label lblPass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "密码"
      Height          =   180
      Left            =   1620
      TabIndex        =   2
      Top             =   1200
      Width           =   360
   End
   Begin VB.Label lblSvr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "服务器"
      Height          =   180
      Left            =   1440
      TabIndex        =   0
      Top             =   1635
      Width           =   540
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Set gcnOracle = Nothing
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If txtUser.Text = "" Or txtPass.Text = "" Then Exit Sub
    
    On Error Resume Next
    
    Set gcnOracle = New ADODB.Connection
    gcnOracle.Provider = "MSDataShape"
    gcnOracle.Open "Driver={Microsoft ODBC for Oracle};Server=" & txtSvr.Text, txtUser.Text, txtPass.Text
    
    If Err.Number <> 0 Then
        MsgBox "连接失败！", vbInformation, "错误"
        txtPass.SetFocus
        Exit Sub
    End If
   
    '修改注册表
    SaveSetting "ZLSOFT", "自助签到\登录信息", "USER", txtUser.Text
    SaveSetting "ZLSOFT", "自助签到\登录信息", "SERVER", txtSvr.Text
    
    Call InitCommon(gcnOracle)
   
    frmSetStock.Show
    
    Unload Me
    Exit Sub
End Sub

Private Sub Form_Load()
    txtUser.Text = GetSetting("ZLSOFT", "自助签到\登录信息", "USER", "")
    txtSvr.Text = GetSetting("ZLSOFT", "自助签到\登录信息", "SERVER", "")
End Sub
