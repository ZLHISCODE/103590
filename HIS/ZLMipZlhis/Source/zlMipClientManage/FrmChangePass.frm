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
   Begin VB.CommandButton CDM确认 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3510
      TabIndex        =   3
      Top             =   240
      Width           =   1230
   End
   Begin VB.CommandButton CMD放弃 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3510
      TabIndex        =   4
      Top             =   690
      Width           =   1230
   End
   Begin VB.Frame Fra密码 
      Caption         =   "更改密码"
      Height          =   1455
      Left            =   150
      TabIndex        =   5
      Top             =   150
      Width           =   3165
      Begin VB.TextBox TXT确认密码 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1110
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1005
         Width           =   1590
      End
      Begin VB.TextBox TXT新密码 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1110
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   645
         Width           =   1590
      End
      Begin VB.TextBox TXT原密码 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1110
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   270
         Width           =   1590
      End
      Begin VB.Label Lbl旧密码 
         AutoSize        =   -1  'True
         Caption         =   "旧密码"
         Height          =   180
         Left            =   450
         TabIndex        =   8
         Top             =   330
         Width           =   540
      End
      Begin VB.Label Lbl新密码 
         AutoSize        =   -1  'True
         Caption         =   "新密码"
         Height          =   180
         Left            =   450
         TabIndex        =   7
         Top             =   705
         Width           =   540
      End
      Begin VB.Label Lbl密码验证 
         AutoSize        =   -1  'True
         Caption         =   "密码验证"
         Height          =   180
         Left            =   270
         TabIndex        =   6
         Top             =   1065
         Width           =   720
      End
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

Private Sub CDM确认_Click()
    Dim strPassword As String
    Dim strServer As String, strError As String
    Dim intPos As Integer
    Dim strSQL As String, rsData As New ADODB.Recordset
    Dim arrTmp As Variant, lngLen As Long, i As Long, intChr As Integer
    Dim blnHaveNum As Boolean, blnAlpha As Boolean, blnChar As Boolean
    Dim blnPwdLen As Boolean, intPwdMin As Integer, intPwdMax As Integer
    Dim blnComplex As Boolean, strOterChrs As String
    Dim blnTransPassword As Boolean

    If Trim(TXT原密码.Text) = "" Then
        MsgBox "请输入旧密码！", vbInformation, gstrSysName
        TXT原密码.SetFocus
        Exit Sub
    End If
    If Trim(TXT新密码.Text) = "" Then
        MsgBox "请输入新密码！", vbInformation, gstrSysName
        TXT新密码.SetFocus
        Exit Sub
    End If
    If Trim(TXT确认密码.Text) = "" Then
        MsgBox "请输入密码验证！", vbInformation, gstrSysName
        TXT确认密码.SetFocus
        Exit Sub
    End If
    If TXT新密码.Text <> TXT确认密码.Text Then
        MsgBox "新密码输入错误，请重新输入！", vbInformation, gstrSysName
        TXT新密码.SetFocus
        Exit Sub
    End If
    
    If TXT新密码.Text = Trim(TXT原密码.Text) Then
        MsgBox "新密码和旧密码完全一样，请重新输入！", vbInformation, gstrSysName
        TXT新密码.SetFocus
        Exit Sub
    End If
    
    strPassword = Trim(TXT原密码.Text)
    If Trim(strPassword) <> "" And Len(strPassword) <> 1 Then
        If Mid(strPassword, Len(strPassword) - 1, 1) = "/" Or Mid(strPassword, Len(strPassword) - 1, 1) = "@" Or Mid(strPassword, 1, 1) = "/" Or Mid(strPassword, 1, 1) = "@" Then
            If TXT原密码.Enabled Then TXT原密码.SetFocus
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
    
    If Not gclsMsgOracle.OraDataOpen(strServer, mstrUserName, strPassword, True) Then
        Exit Sub
    Else
        gstrDbUser = UCase(mstrUserName)
        Call gclsBusiness.InitBusiness(gclsMsgOracle, "", gstrDbUser)
        
        strSQL = "Select 参数号,Nvl(参数值,缺省值) 参数值 From zlOptions Where 参数号 in (20,21,22,23)"
        Set rsData = gclsMsgOracle.OpenSQLRecord(strSQL, Me.Caption)
        blnPwdLen = False: intPwdMin = 0: intPwdMax = 0
        blnComplex = False: strOterChrs = ""
        Do While Not rsData.EOF
            Select Case rsData!参数号
                Case 20 '是否控制密码长度
                    blnPwdLen = Val(rsData!参数值 & "") = 1
                Case 21 '密码长度下限
                    intPwdMin = Val(rsData!参数值 & "")
                Case 22 '密码长度上限
                    intPwdMax = Val(rsData!参数值 & "")
                Case 23 '是否控制密码复杂度
                    blnComplex = Val(rsData!参数值 & "") = 1
            End Select
            rsData.MoveNext
        Loop
        '生成悬浮提示
        If blnPwdLen Then
            If intPwdMin = intPwdMax Then
                TXT新密码.ToolTipText = "密码必须为" & intPwdMax & " 位字符。"
            Else
                TXT新密码.ToolTipText = "密码必须为" & intPwdMin & "至" & intPwdMax & " 位字符。"
            End If
         End If
         If blnComplex Then
            If TXT新密码.ToolTipText <> "" Then
                TXT新密码.ToolTipText = TXT新密码.ToolTipText & vbNewLine & "至少包含一个数字、一个字母与一个特殊字符组成。"
            Else
                TXT新密码.ToolTipText = "至少由一个数字、一个字母与一个特殊字符组成。"
            End If
         End If
         TXT确认密码.ToolTipText = TXT新密码.ToolTipText
         strPassword = Trim(TXT新密码.Text)
        '长度检查
        lngLen = ActualLen(strPassword)
        If lngLen <> Len(strPassword) Then
            MsgBox "新密码包含双字节字符，请检查！", vbInformation, gstrSysName
            Exit Sub
        End If
        If blnPwdLen Then
            If Not (lngLen >= intPwdMin And lngLen <= intPwdMax) Then
                If intPwdMin = intPwdMax Then
                    MsgBox "密码必须为" & intPwdMax & " 位字符！", vbInformation, gstrSysName
                    Exit Sub
                Else
                    MsgBox "密码必须为" & intPwdMin & "至" & intPwdMax & " 位字符！", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        End If
        For i = 1 To Len(strPassword)
            intChr = Asc(UCase(Mid(strPassword, i, 1)))
            If intChr >= 32 And intChr < 127 Then
                'Dim blnHaveNum As Boolean, blnAlpha As Boolean, blnChar As Boolean
                Select Case intChr
                    Case 48 To 57 '数字
                        blnHaveNum = True
                    Case 65 To 90 '字母
                        blnAlpha = True
                    Case 32, 34, 47, 64  '空格,双引号,/,@
                        strOterChrs = strOterChrs & Chr(intChr)
                    Case Is < 48, 58 To 64, 91 To 96, Is > 122
                        blnChar = True
                End Select
            Else
                strOterChrs = strOterChrs & Chr(intChr)
            End If
        Next
        If strOterChrs <> "" Then
            MsgBox "密码不容许有以下字符：" & strOterChrs, vbInformation, gstrSysName
            Exit Sub
        ElseIf Not (blnHaveNum And blnAlpha And blnChar) And blnComplex Then
            MsgBox "密码至少由一个数字、一个字母与一个特殊字符组成。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If gobjRegister.UpdateUserPassword(mstrUserName, strPassword, blnTransPassword, strError) Then
            MsgBox "密码修改成功", vbInformation + vbOKOnly, "提示"
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

Private Sub CMD放弃_Click()
    mstrUserName = ""
    mstrPwd = ""
    mstrServer = ""
    Unload Me
End Sub

Private Sub Form_Activate()
    Call SetWindowPos(Me.hwnd, HWND_TOPMOST, Me.Left / 15, Me.Top / 15, Me.Height / 15, Me.Width / 15, SWP_NOSIZE + SWP_SHOWWINDOW)
    If mstrPwd <> "" And mstrUserName = mstrPwd Then
        TXT原密码.Enabled = False
    ElseIf TXT原密码.Text = "" Then
        TXT原密码.SetFocus
    Else
        TXT新密码.SetFocus
    End If
End Sub

Private Sub Form_Load()
    TXT原密码.Text = mstrPwd
End Sub

Private Sub TXT确认密码_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call CDM确认_Click
End Sub

Private Sub TXT新密码_GotFocus()
    GetFocus TXT新密码
End Sub

Private Sub TXT新密码_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}", 1
End Sub

Private Sub TXT原密码_GotFocus()
    GetFocus TXT原密码
End Sub

Private Sub TXT确认密码_GotFocus()
    GetFocus TXT确认密码
End Sub

Private Sub GetFocus(ByVal TxtBox As TextBox)
    With TxtBox
        .SelStart = 0
        .SelLength = LenB(StrConv(.Text, vbFromUnicode))
    End With
End Sub

Private Sub TXT原密码_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}", 1
End Sub
