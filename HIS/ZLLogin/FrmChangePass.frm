VERSION 5.00
Begin VB.Form FrmChangePass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "修改密码"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4860
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   4860
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraPwd 
      Caption         =   "更改密码"
      Height          =   1455
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   3165
      Begin VB.TextBox txtOldPWD 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1110
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   270
         Width           =   1590
      End
      Begin VB.TextBox txtNewPWD 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1110
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   645
         Width           =   1590
      End
      Begin VB.TextBox txtComfirmPwd 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1110
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   1005
         Width           =   1590
      End
      Begin VB.Label lblComfirmPwd 
         AutoSize        =   -1  'True
         Caption         =   "密码验证"
         Height          =   180
         Left            =   270
         TabIndex        =   4
         Top             =   1065
         Width           =   720
      End
      Begin VB.Label lblNewPwd 
         AutoSize        =   -1  'True
         Caption         =   "新密码"
         Height          =   180
         Left            =   450
         TabIndex        =   2
         Top             =   705
         Width           =   540
      End
      Begin VB.Label lblOldPwd 
         AutoSize        =   -1  'True
         Caption         =   "旧密码"
         Height          =   180
         Left            =   450
         TabIndex        =   0
         Top             =   330
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3480
      TabIndex        =   7
      Top             =   660
      Width           =   1230
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3480
      TabIndex        =   6
      Top             =   210
      Width           =   1230
   End
End
Attribute VB_Name = "FrmChangePass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'入参
Private mfrmParent As Object '父窗体
Private mstrUserName As String '原始用户名
Private mstrPwd As String '原始密码
Private mstrServer As String '原始服务器
Private mbln转换 As Boolean '是否密码要转换
'模块变量
Private mblnOk As Boolean

Public Function ShowMe(ByVal frmParent As Object, ByVal strUserName As String, ByRef strPWD As String, ByRef strServer As String, Optional ByVal blnTrans As Boolean) As Boolean
'功能：修改密码
'参数：frmParent=父窗体
'          strUserName=用户名
'          strPwd=密码
'          strServer=服务器
    Set mfrmParent = frmParent
    mstrUserName = strUserName
    mstrPwd = strPWD
    mstrServer = strServer
    mbln转换 = blnTrans
    mblnOk = False
    Me.Show vbModal
    strUserName = mstrUserName
    strPWD = mstrPwd
    strServer = mstrServer
    ShowMe = mblnOk
End Function

Private Sub cmdOK_Click()
    Dim strPassword As String
    Dim strServer As String, strError As String, strToolTip As String
    Dim intPos As Integer
    Dim cnOracle As ADODB.Connection
    Dim blnTransPassword As Boolean
    
    If Trim(txtOldPWD.Text) = "" Then
        MsgBox "请输入旧密码！", vbInformation, gstrSysName
        txtOldPWD.SetFocus
        Exit Sub
    End If
    If Trim(txtNewPWD.Text) = "" Then
        MsgBox "请输入新密码！", vbInformation, gstrSysName
        txtNewPWD.SetFocus
        Exit Sub
    End If
    If Trim(txtComfirmPwd.Text) = "" Then
        MsgBox "请输入密码验证！", vbInformation, gstrSysName
        txtComfirmPwd.SetFocus
        Exit Sub
    End If
    If txtNewPWD.Text <> txtComfirmPwd.Text Then
        MsgBox "新密码输入错误，请重新输入！", vbInformation, gstrSysName
        txtNewPWD.SetFocus
        Exit Sub
    End If
    
    If txtNewPWD.Text = Trim(txtOldPWD.Text) Then
        MsgBox "新密码和旧密码完全一样，请重新输入！", vbInformation, gstrSysName
        txtNewPWD.SetFocus
        Exit Sub
    End If
    
    strPassword = Trim(txtOldPWD.Text)
    If Trim(strPassword) <> "" And Len(strPassword) <> 1 Then
        If Mid(strPassword, Len(strPassword) - 1, 1) = "/" Or Mid(strPassword, Len(strPassword) - 1, 1) = "@" Or Mid(strPassword, 1, 1) = "/" Or Mid(strPassword, 1, 1) = "@" Then
            If txtOldPWD.Enabled Then txtOldPWD.SetFocus
            MsgBox "旧密码错误！", vbInformation, gstrSysName
            Exit Sub
        End If
    End If

    '分离字符串
    intPos = InStr(strPassword, "@")
    If intPos > 0 Then
        strServer = Mid(strPassword, intPos + 1)
        strPassword = Mid(strPassword, 1, intPos - 1)
    End If
    If strServer = "" Then
        strServer = mstrServer
    End If
    
    blnTransPassword = Not (UCase(mstrUserName) = "SYS" Or UCase(mstrUserName) = "SYSTEM") Or mbln转换
    Set cnOracle = gobjRegister.GetConnection(strServer, mstrUserName, strPassword, blnTransPassword, , strError)
    If cnOracle.State = adStateClosed Then
        If InStr(strError, "ORA-28001") > 0 Then
            strError = "密码已经过期。请联系管理员重置密码！"
        End If
        MsgBox "原始密码验证失败：" & vbCrLf & strError, vbInformation, "提示"
        Exit Sub
    Else
        strPassword = Trim(txtNewPWD.Text)
        If Not CheckPWDComplex(cnOracle, strPassword, strToolTip) Then
            txtNewPWD.ToolTipText = strToolTip
            txtComfirmPwd.ToolTipText = txtNewPWD.ToolTipText
            txtNewPWD.SetFocus
            Exit Sub
        Else
            txtNewPWD.ToolTipText = strToolTip
            txtComfirmPwd.ToolTipText = txtNewPWD.ToolTipText
        End If
        
        If gobjRegister.UpdateUserPassword(cnOracle, mstrUserName, strPassword, blnTransPassword, strError) Then
            MsgBox "密码修改成功!", vbInformation, gstrSysName
            mstrServer = strServer
            mstrPwd = strPassword
            mblnOk = True
        Else
            If strError <> "" Then
                MsgBox "密码修改失败：" & vbCrLf & strError, vbExclamation, "提示"
            End If
            Exit Sub
        End If
    End If
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    mstrUserName = ""
    mstrPwd = ""
    mstrServer = ""
    Unload Me
End Sub

Private Sub Form_Activate()
    Call SetWindowPos(Me.hwnd, HWND_TOPMOST, Me.Left / 15, Me.Top / 15, Me.Height / 15, Me.Width / 15, SWP_NOSIZE + SWP_SHOWWINDOW)
    If mstrPwd <> "" And mstrUserName = mstrPwd Then
        txtOldPWD.Enabled = False
    ElseIf txtOldPWD.Text = "" Then
        txtOldPWD.SetFocus
    Else
        txtNewPWD.SetFocus
    End If
End Sub

Private Sub Form_Load()
    txtOldPWD.Text = mstrPwd
End Sub

Private Sub txtComfirmPwd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call cmdOK_Click
End Sub

Private Sub txtNewPWD_GotFocus()
    GetFocus txtNewPWD
End Sub

Private Sub txtNewPWD_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}", 1
End Sub

Private Sub txtOldPWD_GotFocus()
    GetFocus txtOldPWD
End Sub

Private Sub txtComfirmPwd_GotFocus()
    GetFocus txtComfirmPwd
End Sub

Private Sub GetFocus(ByVal TxtBox As TextBox)
    With TxtBox
        .SelStart = 0
        .SelLength = LenB(StrConv(.Text, vbFromUnicode))
    End With
End Sub

Private Sub txtOldPWD_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}", 1
End Sub
