VERSION 5.00
Begin VB.Form frmBloodVerification 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "身份验证"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7290
   Icon            =   "frmBloodVerification.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraLine 
      Height          =   90
      Left            =   75
      TabIndex        =   18
      Top             =   1935
      Width           =   7125
   End
   Begin VB.CommandButton CMDcancel 
      Caption         =   "取消"
      Height          =   350
      Left            =   6060
      TabIndex        =   10
      Top             =   2100
      Width           =   1100
   End
   Begin VB.CommandButton CMDok 
      Caption         =   "确定"
      Height          =   350
      Left            =   4860
      TabIndex        =   9
      Top             =   2100
      Width           =   1100
   End
   Begin VB.Frame Fra1 
      Caption         =   "核对人"
      Height          =   1740
      Index           =   1
      Left            =   4035
      TabIndex        =   11
      Top             =   120
      Width           =   3135
      Begin VB.PictureBox picDown 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   2715
         Picture         =   "frmBloodVerification.frx":030A
         ScaleHeight     =   240
         ScaleWidth      =   225
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   765
         Width           =   255
      End
      Begin VB.TextBox TXT密码 
         Appearance      =   0  'Flat
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1050
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   1170
         Width           =   1920
      End
      Begin VB.TextBox txt用户 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   1
         Left            =   1050
         TabIndex        =   5
         Top             =   330
         Width           =   1920
      End
      Begin VB.TextBox txt姓名 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   1
         Left            =   1050
         TabIndex        =   6
         Top             =   750
         Width           =   1920
      End
      Begin VB.Label Lbl口令 
         AutoSize        =   -1  'True
         Caption         =   "密      码"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   1215
         Width           =   900
      End
      Begin VB.Label Lbl用户名 
         AutoSize        =   -1  'True
         Caption         =   "核对人帐号"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   375
         Width           =   900
      End
      Begin VB.Label Lbl姓名 
         AutoSize        =   -1  'True
         Caption         =   "核对人姓名"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   810
         Width           =   900
      End
   End
   Begin VB.Frame Fra1 
      Caption         =   "接收人"
      Height          =   1740
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.PictureBox picDown 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   2715
         Picture         =   "frmBloodVerification.frx":06C3
         ScaleHeight     =   240
         ScaleWidth      =   225
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   765
         Width           =   255
      End
      Begin VB.TextBox TXT密码 
         Appearance      =   0  'Flat
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   1050
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1170
         Width           =   1920
      End
      Begin VB.TextBox txt用户 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   0
         Left            =   1050
         TabIndex        =   1
         Top             =   330
         Width           =   1920
      End
      Begin VB.TextBox txt姓名 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   0
         Left            =   1050
         TabIndex        =   2
         Top             =   750
         Width           =   1920
      End
      Begin VB.Label Lbl口令 
         AutoSize        =   -1  'True
         Caption         =   "密      码"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   1215
         Width           =   900
      End
      Begin VB.Label Lbl用户名 
         AutoSize        =   -1  'True
         Caption         =   "接收人帐号"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   390
         Width           =   900
      End
      Begin VB.Label Lbl姓名 
         AutoSize        =   -1  'True
         Caption         =   "接收人姓名"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   810
         Width           =   900
      End
   End
   Begin VB.Image Image1 
      Height          =   540
      Left            =   3405
      Picture         =   "frmBloodVerification.frx":0A7C
      Stretch         =   -1  'True
      Top             =   660
      Width           =   540
   End
End
Attribute VB_Name = "frmBloodVerification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnReceive As Boolean
Private mblnAutomatic As Boolean
Private mobjfrmMain As Object
Private mblnOK As Boolean
Private mblnUserIsOk As Boolean
Private mstr核收人 As String

Public Property Get str核收人() As String
    str核收人 = mstr核收人
End Property

Public Function ShowCheck(frmMain As Object, Optional blnAutomatic As Boolean = True) As Boolean
    '功能:显示接收核对验证页面，对接收操作进行审核
    '参数:frmmain-父窗体，blnAutomatic-用户是否自己填写接收人信息，true表示自动根据登陆用户的信息进行提取，false表示用户自己填写并验证接收人信息。
    Set mobjfrmMain = frmMain
    mblnAutomatic = blnAutomatic
    
    If mblnAutomatic = True Then '当允许自动读取用户数据时
        mblnUserIsOk = userIsOk
        If mblnUserIsOk = False Then MsgBox "用户不符合条件，请手动添加接收人", vbInformation, gstrSysName: GoTo Skip
        TXT密码(0).Enabled = False
        TXT密码(0).Text = "123"
        txt用户(0).Text = UserInfo.编号
        txt姓名(0).Text = UserInfo.姓名
        picDown(0).Visible = False
        txt用户(0).Enabled = False
        txt姓名(0).Enabled = False
    End If
Skip:
    Me.Show 1, mobjfrmMain
    ShowCheck = mblnOK
    mblnOK = False
End Function
 
Private Function userIsOk() As Boolean
    '功能：判断用户是否符合接收人的条件，比如接收人是否是护士或者医生,接收人所在部门是否是临床部门等
    '参数：参数直接使用userinfo里面的内容，所以这里不传入参数
    Dim strSql As String
    Dim rspeople As ADODB.Recordset
    On Error GoTo Errorhand
    strSql = " select rownum || '-' || b.id as id,b.编号,b.姓名,b.简码,a.名称 as 部门名称 " & _
             " from 部门表 a,人员表 b,人员性质说明 c,部门人员 d,部门性质说明 e,上机人员表 f " & _
             " where a.id=d.部门id and a.id=e.部门id and Instr(',临床,护理,', ',' || e.工作性质 || ',', 1) <> 0 and f.人员id=b.id " & _
             " and d.人员id=b.id and b.id=c.人员id and c.人员性质 in('医生','护士') and b.id=[1]"
    
    Set rspeople = gobjDatabase.OpenSQLRecord(strSql, "人员信息", UserInfo.id)
    If rspeople.RecordCount = 0 Then
        userIsOk = False
    Else
        userIsOk = True
    End If
Errorhand:
End Function

Private Function GetUserName(ByVal objControl As TextBox, ByVal intIndex As Integer, Optional ByVal StrInput As String = "") As Boolean
    Dim rsUser As ADODB.Recordset
    Dim strSql As String, strWhere As String
    Dim vPoint As POINTAPI, blnCancel As Boolean

    On Error GoTo errHand

    If StrInput <> "" Then
         If IsNumeric(StrInput) Then
            strWhere = " And a.编号 Like '" & txt用户(intIndex).Text & "%'"
         ElseIf gobjCommFun.IsNumOrChar(StrInput) Then
            strWhere = " And f.用户名 Like '" & UCase(txt用户(intIndex).Text) & "%'"
         Else
            strWhere = " And a.姓名 Like '" & txt姓名(intIndex).Text & "%'"
         End If
    End If
    vPoint = GetCoordPos(Me.hWnd, objControl.Left + Fra1(intIndex).Left, objControl.Top + Fra1(intIndex).Top) ',b.名称 as 科室,b.id ||
    strSql = " Select distinct f.用户名 || '-' || a.id as ID,f.用户名,a.编号,a.姓名,a.简码 " & _
            " From 人员表 a, 部门表 b, 部门人员 c, 部门性质说明 d, 人员性质说明 e,上机人员表 f " & _
            " Where a.Id = c.人员id And b.Id = c.部门id And a.Id = e.人员id And b.Id = d.部门id and f.人员id=a.id  And Instr(',临床,护理,', ',' || d.工作性质 || ',', 1) <> 0 And " & _
            "  e.人员性质 In ('医生', '护士') And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null)" & strWhere
    Set rsUser = gobjDatabase.ShowSQLSelect(Me, strSql, 0, "", False, txt用户(intIndex).Text, "请选择一个取血人员", False, False, True, vPoint.X, vPoint.Y, objControl.Height, blnCancel, False, False, False)

    If Not rsUser Is Nothing Then
        If blnCancel = False Then
            If rsUser.EOF Then Exit Function
            Lbl用户名(intIndex).Tag = Split(rsUser!id, "-")(1) '用户id
            txt用户(intIndex).Text = Nvl(rsUser!用户名) '用户的编号
            txt用户(intIndex).Tag = Nvl(rsUser!用户名) '用户的登陆名
            objControl.Text = Nvl(rsUser!姓名) '用户的姓名
            objControl.Tag = objControl.Text '用户的姓名
            objControl.SetFocus
            GetUserName = True
        End If
    Else
        If StrInput = "" And blnCancel = False Then
            MsgBox "没有对应的临床护士和医生信息，请在人员管理中设置！", vbInformation, gstrSysName
        End If
    End If
    
    Exit Function
errHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub CMDcancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub CMDok_Click()
    '功能：点击确定后，判断用户输入数据的正确性和规范性，同时判断用户输入的密码和用户信息是否正确。
    Dim strNote As String
    Dim strUserNo接收人 As String
    Dim strUserNo核收人 As String
    Dim strUserName接收人 As String
    Dim strUserName核收人 As String
    Dim strPassword接收人 As String
    Dim strPassword核收人 As String

    
    Dim strServerName As String
    On Error GoTo InputError
    
    '取血人验证检查
    '------检验用户是否oracle合法用户----------------
    strUserNo接收人 = Trim(txt用户(0).Text)
    strUserName接收人 = Trim(txt用户(0).Tag)
    strPassword接收人 = Trim(TXT密码(0).Text)
    strUserNo核收人 = Trim(txt用户(1).Text)
    strUserName核收人 = Trim(txt用户(1).Tag)
    strPassword核收人 = Trim(TXT密码(1).Text)
    
    '有效字符串效验
    If mblnUserIsOk = False Then '在需要用户手动填写接收人的情况下，要对接收人的数据进行判断
        If Len(Trim(txt用户(0))) = 0 Then
            strNote = "请输入接收人帐号"
            Call gobjControl.ControlSetFocus(txt用户(0))
            GoTo InputError
        End If
        If Len(strUserNo接收人) <> 1 Then
            If Mid(strUserNo接收人, 1, 1) = "/" Or Mid(strUserNo接收人, 1, 1) = "@" Or Mid(strUserNo接收人, Len(strUserNo接收人) - 1, 1) = "/" Or Mid(strUserNo接收人, Len(strUserNo接收人) - 1, 1) = "@" Then
                strNote = "接收人帐号错误"
                Call gobjControl.ControlSetFocus(txt用户(0))
                Exit Sub
            End If
        End If
        If Trim(strPassword接收人) <> "" And Len(strPassword接收人) <> 1 Then
            If Mid(strPassword接收人, Len(strPassword接收人) - 1, 1) = "/" Or Mid(strPassword接收人, Len(strPassword接收人) - 1, 1) = "@" Or Mid(strPassword接收人, 1, 1) = "/" Or Mid(strPassword接收人, 1, 1) = "@" Then
                strNote = "接收人帐号密码错误"
                Call gobjControl.ControlSetFocus(TXT密码(0))
                GoTo InputError
            End If
        End If
        If Len(Trim(strPassword接收人)) = 0 Then
            strNote = "请输入接收人帐号密码"
            Call gobjControl.ControlSetFocus(TXT密码(0))
            GoTo InputError
        End If
        If GetObjectRegister = False Then Exit Sub
        strServerName = gobjRegister.GetServerName
        If gobjRegister.LoginValidate(strServerName, strUserName接收人, strPassword接收人, strNote) = False Then
            TXT密码(0).Text = ""
            Call gobjControl.ControlSetFocus(TXT密码(0))
            GoTo InputError
        End If
    End If
    
    If Len(Trim(txt用户(1))) = 0 Then
        strNote = "请输入核收人帐号"
        Call gobjControl.ControlSetFocus(txt用户(1))
        GoTo InputError
    End If
    If Len(strUserNo核收人) <> 1 Then
        If Mid(strUserNo核收人, 1, 1) = "/" Or Mid(strUserNo核收人, 1, 1) = "@" Or Mid(strUserNo核收人, Len(strUserNo核收人) - 1, 1) = "/" Or Mid(strUserNo核收人, Len(strUserNo核收人) - 1, 1) = "@" Then
            strNote = "核收人帐号错误"
            Call gobjControl.ControlSetFocus(txt用户(1))
            Exit Sub
        End If
    End If
    If Trim(strPassword核收人) <> "" And Len(strPassword核收人) <> 1 Then
        If Mid(strPassword核收人, Len(strPassword核收人) - 1, 1) = "/" Or Mid(strPassword核收人, Len(strPassword核收人) - 1, 1) = "@" Or Mid(strPassword核收人, 1, 1) = "/" Or Mid(strPassword核收人, 1, 1) = "@" Then
            strNote = "核收人帐号密码错误"
            Call gobjControl.ControlSetFocus(TXT密码(1))
            GoTo InputError
        End If
    End If
    If Len(Trim(strPassword核收人)) = 0 Then
        strNote = "请输入核收人帐号密码"
        Call gobjControl.ControlSetFocus(TXT密码(1))
        GoTo InputError
    End If


    '接收人和核收人不能是同一个
    If txt姓名(0).Text = txt姓名(1).Text Or txt用户(0).Text = txt用户(1).Text Then
        strNote = "接收人和核收人不能是同一个人，请重新核对"
        Call gobjControl.ControlSetFocus(txt姓名(1))
        GoTo InputError
    End If
    '用户登录验证
    If GetObjectRegister = False Then Exit Sub
    strServerName = gobjRegister.GetServerName
    If gobjRegister.LoginValidate(strServerName, strUserName核收人, strPassword核收人, strNote) = False Then
        TXT密码(1).Text = ""
        Call gobjControl.ControlSetFocus(TXT密码(1))
        GoTo InputError
    End If
    
    mblnOK = True
    mstr核收人 = txt姓名(1).Text
    Unload Me
    Exit Sub
InputError:
    If strNote <> "" Then
        MsgBox strNote, vbInformation, gstrSysName
    End If
    Exit Sub
End Sub

Private Sub picDown_Click(Index As Integer)
    If GetUserName(txt姓名(Index), Index) = True Then gobjCommFun.PressKey vbKeyTab
End Sub

Private Sub TXT密码_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then gobjCommFun.PressKey vbKeyTab
End Sub

Private Sub txt姓名_KeyPress(Index As Integer, KeyAscii As Integer)
    TXT密码(Index).Text = ""
    If KeyAscii = vbKeyReturn Then
        If GetUserName(txt姓名(Index), Index, txt姓名(Index).Text) = True Then gobjCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txt用户_KeyPress(Index As Integer, KeyAscii As Integer)
    TXT密码(Index).Text = ""
    If KeyAscii = vbKeyReturn Then
        If GetUserName(txt姓名(Index), Index, txt用户(Index).Text) = True Then gobjCommFun.PressKey vbKeyTab
    End If
End Sub
