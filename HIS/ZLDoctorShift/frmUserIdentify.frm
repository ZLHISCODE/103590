VERSION 5.00
Begin VB.Form frmUserIdentify 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "用户验证"
   ClientHeight    =   2040
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   4170
   Icon            =   "frmUserIdentify.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   -360
      TabIndex        =   6
      Top             =   1335
      Width           =   5025
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2865
      TabIndex        =   3
      Top             =   1590
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   1755
      TabIndex        =   2
      Top             =   1590
      Width           =   1100
   End
   Begin VB.TextBox txtPass 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1950
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   960
      Width           =   1920
   End
   Begin VB.TextBox txtUser 
      Enabled         =   0   'False
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1950
      TabIndex        =   0
      Top             =   555
      Width           =   1920
   End
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "身份验证，请输入用户名与密码"
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   1335
      TabIndex        =   7
      Top             =   105
      Width           =   2520
   End
   Begin VB.Image imgFlag 
      Height          =   720
      Left            =   315
      Picture         =   "frmUserIdentify.frx":000C
      Top             =   240
      Width           =   720
   End
   Begin VB.Label lblPass 
      AutoSize        =   -1  'True
      Caption         =   "密码"
      Height          =   180
      Left            =   1500
      TabIndex        =   5
      Top             =   1020
      Width           =   360
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      Caption         =   "用户名"
      Height          =   180
      Left            =   1320
      TabIndex        =   4
      Top             =   615
      Width           =   540
   End
End
Attribute VB_Name = "frmUserIdentify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrNote As String
Private mlngSys As Long
Private mstrServer As String
Private mstrUserName As String
Private mblnOK As Boolean
Private mblnTransPass As Boolean

Public Function ShowMe(frmParent As Object, ByVal strNote As String, ByVal lngSys As Long, ByVal strUser As String, ByVal blnTransPass As Boolean) As Boolean
'参数：strNote-提示信息(简短)；lngSys-系统号；strUser-用户名；blnTransPass是否转换密码

    mstrNote = strNote
    mlngSys = lngSys
    mstrUserName = strUser
    mblnTransPass = blnTransPass
    
    Me.Show 1, frmParent
    ShowMe = mblnOK
End Function

Private Sub cmdOK_Click()
    Dim strUser As String
    Dim strPass As String
    
    strUser = Trim(txtUser.Text)
    strPass = Trim(txtPass.Text)
    
    '有效字符串效验
    If strUser = "" Then
        MsgBox "请输入用户名。", vbInformation, gstrSysName
        txtUser.SetFocus: Exit Sub
    End If
    If InStr(strUser, "/") > 0 Or InStr(strUser, "@") > 0 Then
        MsgBox "输入了无效的用户名，请重新输入。", vbInformation, gstrSysName
        txtUser.SetFocus: Exit Sub
    End If
    If strPass = "" Then
        MsgBox "请输入密码。", vbInformation, gstrSysName
        txtPass.SetFocus: Exit Sub
    End If
    If InStr(strPass, "/") > 0 Or InStr(strPass, "@") > 0 Then
        MsgBox "输入了无效的密码，请重新输入。", vbInformation, gstrSysName
        txtPass.Text = "": txtPass.SetFocus: Exit Sub
    End If

    If Not OpenOracle(strUser, strPass) Then
        Unload Me
        Exit Sub
    End If
    mblnOK = True
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub Form_Activate()
    If Trim(txtUser.Text) <> "" Then txtPass.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        If Me.ActiveControl Is txtPass Then
            Call cmdOK_Click
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub Form_Load()

    mblnOK = False
    mstrServer = GetSetting("ZLSOFT", "注册信息\登陆信息", "Server", "")
    If mstrUserName = "" Then
        Call zlControl.ControlSetFocus(txtUser)
    Else
        txtUser.Enabled = False
        txtUser.Text = mstrUserName
        Call zlControl.ControlSetFocus(txtPass)
    End If
    If mstrNote <> "" Then lblNote.Caption = mstrNote
End Sub

Private Sub txtPass_GotFocus()
    Call zlControl.TxtSelAll(txtPass)
End Sub

Private Sub SetEnabled(ByVal blnEnabled As Boolean)
    cmdCancel.Enabled = blnEnabled
    cmdOK.Enabled = blnEnabled
    Screen.MousePointer = IIf(Not blnEnabled, 11, 0)
End Sub

Private Function OpenOracle(ByVal strUser As String, ByVal strPass As String) As Boolean
'功能：验证用户,并返回用户名和连接
    Dim rsTmp As New ADODB.Recordset
    Dim cnNew As ADODB.Connection
    Dim objRegister As Object
    Dim strSQL As String
    
    Call SetEnabled(False)
    strUser = UCase(strUser)
    
    On Error GoTo errH
    
    '检查用户名
    strSQL = "Select UserName From All_Users Where UserName=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strUser)
    If rsTmp.EOF Then
        MsgBox "该用户不存在。", vbInformation, gstrSysName
        Call SetEnabled(True)
        txtPass.Text = "": txtUser.SetFocus
        Exit Function
    End If
    
    '检查连接
    Set objRegister = GetObject("", "zlRegister.clsRegister")
    Set cnNew = objRegister.GetConnection(mstrServer, strUser, strPass, mblnTransPass, , , False)
    Call SetEnabled(True)
    If cnNew.State = adStateClosed Then
        txtPass.Text = "": Call zlControl.ControlSetFocus(txtPass)
        Set cnNew = Nothing: Exit Function
    End If
    
    OpenOracle = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

