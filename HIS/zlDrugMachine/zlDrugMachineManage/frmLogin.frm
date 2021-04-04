VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "操作员登录"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5055
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   5055
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txtUser 
      Height          =   300
      Left            =   2670
      TabIndex        =   1
      Top             =   240
      Width           =   1920
   End
   Begin VB.TextBox txtPass 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   2670
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   675
      Width           =   1920
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确认(&O)"
      Default         =   -1  'True
      Height          =   345
      Left            =   2280
      TabIndex        =   6
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   345
      Left            =   3480
      TabIndex        =   7
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   0
      TabIndex        =   8
      Top             =   1500
      Width           =   5025
   End
   Begin VB.ComboBox cmbDatabase 
      Height          =   300
      Left            =   2670
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1095
      Width           =   1920
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "用户名(&U)"
      Height          =   180
      Left            =   1800
      TabIndex        =   0
      Top             =   300
      Width           =   810
   End
   Begin VB.Label lblPass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "密码(&P)"
      Height          =   180
      Left            =   1800
      TabIndex        =   2
      Top             =   735
      Width           =   630
   End
   Begin VB.Label lblDatabase 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "服务器(&S)"
      Height          =   180
      Left            =   1800
      TabIndex        =   4
      Top             =   1155
      Width           =   810
   End
   Begin VB.Image imgKey 
      Appearance      =   0  'Flat
      Height          =   1200
      Left            =   240
      Picture         =   "frmLogin.frx":179A
      Top             =   240
      Width           =   1410
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnShow As Boolean                         '显示状态（Load事件后）
Private mcolServer As New Collection                '保存服务器串列表
Private mblnReturn As Boolean                       '返回值； True确定；False取消
Private mbytCount As Byte

Public Property Get ReturnStatus()
    ReturnStatus = mblnReturn
End Property

Private Sub cmbDatabase_Change()
    If Me.Visible Then
        ClearComponent
        cmdOK.Enabled = cmbDatabase.ListIndex >= 0
    End If
End Sub

Private Sub cmdCancel_Click()
    mblnReturn = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strPass As String
    
    strPass = txtPass.Text
    gstrUser = UCase(Trim(txtUser.Text))
    mbytCount = mbytCount + 1

    '验证用户与密码
    If gobjRegister Is Nothing Then
        '旧加密
        If Not OraDataOpen(Trim(cmbDatabase.Text), gstrUser, IIf(gstrUser = "SYS" Or gstrUser = "SYSTEM", strPass, TranPasswd(strPass))) Then
makReLogin:
            txtPass.SetFocus
            gobjComLib.zlControl.TXTSelAll txtPass
            If mbytCount >= 3 Then
                Unload Me
            End If
            Exit Sub
        End If
    Else
        '新加密
        Set gcnOracle = gobjRegister.GetConnection(Trim(cmbDatabase.Text), gstrUser, strPass, True)
        If gcnOracle.State = adStateClosed Then
            GoTo makReLogin
        End If
    End If

    SaveSetting "ZLSOFT", "注册信息\登陆信息", "USER", Trim(txtUser.Text)
    SaveSetting "ZLSOFT", "注册信息\登陆信息", "SERVER", cmbDatabase.Text
    
    mblnReturn = True
    Unload Me
End Sub

Private Sub Form_Activate()
    Dim strDB As String
    Dim i As Integer
    
    If mblnShow = False Then Exit Sub
    
    txtUser.Text = GetSetting("ZLSOFT", "注册信息\登陆信息", "USER", "")
    strDB = Trim(GetSetting("ZLSOFT", "注册信息\登陆信息", "SERVER", ""))
    For i = 0 To cmbDatabase.ListCount - 1
        If Trim(cmbDatabase.List(i)) = strDB Then
            cmbDatabase.ListIndex = i
            Exit For
        End If
    Next
    
    If txtUser.Text <> "" Then
        If txtPass.Visible And txtPass.Enabled Then txtPass.SetFocus
    Else
        If txtUser.Visible And txtUser.Enabled Then txtUser.SetFocus
    End If
    
    mblnShow = False
End Sub

Private Sub Form_Load()
    mblnReturn = False
    
    Call LoadServer(cmbDatabase, mcolServer)
    BackColor = RGB(240, 240, 240)
    
    mblnShow = True
End Sub

Private Sub txtPass_GotFocus()
    gobjComLib.zlControl.TXTSelAll txtPass
End Sub

Private Sub txtUser_Change()
    If Me.Visible = False Then Exit Sub
    cmdOK.Default = False
    cmdOK.Enabled = Trim(txtUser.Text) <> ""
End Sub

Private Sub txtUser_GotFocus()
    If Me.Visible = False Then Exit Sub
    gobjComLib.zlControl.TXTSelAll txtUser
    mdlMain.OpenIme False
End Sub

Private Sub ClearComponent()
'功能：清空注册表[本机部件]--因为不同的数据库可能使用的系统和版本不同
    If Me.Visible = True Then  '启动时对控件的赋值不考虑在内
        SaveSetting "ZLSOFT", "注册信息", "本机部件", ""
    End If
End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtPass.SetFocus
    End If
End Sub

Private Sub txtUser_LostFocus()
    cmdOK.Default = True
End Sub
