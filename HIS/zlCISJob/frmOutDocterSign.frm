VERSION 5.00
Begin VB.Form frmOutDocterSign 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "书写签名"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   5700
   Icon            =   "frmOutDocterSign.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.ComboBox cboTime 
      Height          =   300
      Left            =   1290
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2580
      Width           =   2310
   End
   Begin VB.Frame fraLine 
      Height          =   15
      Index           =   0
      Left            =   -270
      TabIndex        =   18
      Top             =   510
      Width           =   5985
   End
   Begin VB.CheckBox chkPreText 
      Caption         =   "将签名级别作为前缀加入(&P)"
      Height          =   225
      Left            =   240
      TabIndex        =   8
      Top             =   1950
      Width           =   2565
   End
   Begin VB.CheckBox chkHandSign 
      Caption         =   "显示手签位置(&H)"
      Height          =   240
      Left            =   240
      TabIndex        =   9
      Top             =   2257
      Width           =   1695
   End
   Begin VB.CheckBox chkEsign 
      Caption         =   "数字签名(&E)"
      Height          =   195
      Left            =   4170
      TabIndex        =   7
      Top             =   1380
      Width           =   1365
   End
   Begin VB.TextBox txtPass 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1605
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   1305
      Width           =   1995
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   -270
      TabIndex        =   15
      Top             =   1785
      Width           =   5985
   End
   Begin VB.TextBox txtName 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1605
      MaxLength       =   50
      TabIndex        =   4
      Top             =   960
      Width           =   1995
   End
   Begin VB.OptionButton optName 
      Caption         =   "指定用户(&U)"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   1020
      Width           =   1320
   End
   Begin VB.OptionButton optName 
      Caption         =   "当前用户(&C)"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   660
      Value           =   -1  'True
      Width           =   1320
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4365
      TabIndex        =   13
      Top             =   2820
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4365
      TabIndex        =   12
      Top             =   2430
      Width           =   1095
   End
   Begin VB.ComboBox cmbLevel 
      Height          =   300
      Left            =   1605
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   90
      Width           =   1995
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      Caption         =   "签名时间(&T)"
      Height          =   180
      Left            =   240
      TabIndex        =   10
      Top             =   2640
      Width           =   990
   End
   Begin VB.Label lblPreview 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   120
      TabIndex        =   17
      Top             =   3255
      Width           =   5475
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "签名效果预览:"
      Height          =   180
      Left            =   240
      TabIndex        =   16
      Top             =   3030
      Width           =   1170
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "用户密码(&P)"
      Height          =   180
      Left            =   510
      TabIndex        =   5
      Top             =   1365
      Width           =   990
   End
   Begin VB.Label lblUserName 
      AutoSize        =   -1  'True
      Caption         =   "张三"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1605
      TabIndex        =   14
      Top             =   660
      Width           =   360
   End
   Begin VB.Label lblLevel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "签名级别(&L)"
      Height          =   180
      Left            =   570
      TabIndex        =   0
      Top             =   150
      Width           =   990
   End
End
Attribute VB_Name = "frmOutDocterSign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private frmParent As Object                 '父窗体
Private mlngPatiID As Long
Private mlngPatiPageID As Long
Private mstrSource As String                 '数字签名的源字符串
Private mstrPatiSign As String              '签名显示的名字
Private mintSign As Integer

Private Sign As cEPRSign                    '签名对象
Private objESign As Object                  '电子签名接口部件
Private lngCertID As Long                   '证书ID
Private lngPassType As Long                 '密码验证规则（系统参数） 0-密码；1－数字；2－两者皆可
Private mblnOK As Boolean
Private mobjRegister As Object  '密码验证部件，该部件不用时不能设置为Nothing,否则会清空功能缓存

Private UserSignLevel As EPRSignLevelEnum   '当前用户的签名级别

Private Enum EPRDocTypeEnum
    cpr门诊病历 = 1
    cpr住院病历 = 2
    cpr护理记录 = 3
    cpr护理病历 = 4
    cpr诊断文书 = 5
    cpr知情文件 = 6
    cpr诊疗报告 = 7             '诊疗单据：报告
    cpr诊疗申请 = 8             '诊疗单据：申请
End Enum

Public Enum PatiFromEnum
    cprPF_门诊 = 1              '1-门诊；
    cprPF_住院 = 2              '2-住院；
    cprPF_外来 = 3              '3-外来；
    cprPF_体检 = 4              '4-体检
End Enum

'签名状态
Public Enum EPRSignLevelEnum
    cprSL_空白 = 0              '未签名
    cprSL_经治 = 1              '经治医师签名
    cprSL_主治 = 2              '主治医师签名
    cprSL_主任 = 3              '主任医师签名
    cprSL_正高 = 4              '正高：签名级别不包含，只表示人员居右正高职称，以便区别副主任医师
End Enum

'################################################################################################################
'## 功能：  显示本窗体
'##
'## 参数：  edtThis     :IN     编辑器控件
'##         fParent     :IN     父窗体
'##         strSource   :IN     数字签名的源字符串（从文本中提取，去掉签名提纲）
'################################################################################################################
Public Function ShowMe(ByRef fParent As Object, ByVal strSource As String, _
    lngPatiID As Long, lngPatiPageID As Long) As cEPRSign
    
    Dim bytFileKind As Byte, bytPatiSource As Byte
    Dim lngStart As Long, strPreText As String
    
    Set frmParent = fParent
    mstrSource = strSource
    mlngPatiID = lngPatiID
    mlngPatiPageID = lngPatiPageID
    
    
    bytFileKind = cpr门诊病历
    bytPatiSource = cprPF_门诊
    
    Me.cboTime.Clear
    Me.cboTime.AddItem "不显示"
    Me.cboTime.AddItem Format(Now(), "yyyy-MM-dd hh:mm")
    Me.cboTime.AddItem Format(Now(), "yyyy年MM月dd日 hh:mm")
    
    mintSign = zlDatabase.GetPara("SignShow", glngSys, 1070, 0)
    
    UserSignLevel = GetUserSignLevel(UserInfo.ID, , mlngPatiID, mlngPatiPageID)  '获取用户签名级别
    '根据签名级别来初始化“签名级别＂
    Select Case bytFileKind
    Case cpr护理病历
        cmbLevel.AddItem "1 - 护士"
        cmbLevel.AddItem "3 - 护士长"
        cmbLevel.ListIndex = 0
        If UserSignLevel >= cprSL_主任 Then cmbLevel.ListIndex = 1
    Case cpr诊疗报告
        cmbLevel.AddItem "1 - 医生"
        cmbLevel.AddItem "2 - 主治"
        cmbLevel.AddItem "3 - 主任"
        cmbLevel.ListIndex = 0
        If UserSignLevel >= cprSL_主治 Then cmbLevel.ListIndex = 1
        If UserSignLevel >= cprSL_主任 Then cmbLevel.ListIndex = 2
    Case Else
        cmbLevel.AddItem "1 - 经治医师"
        cmbLevel.AddItem "2 - 主治医师"
        cmbLevel.AddItem "3 - 副主任医师"
        cmbLevel.AddItem "4 - 主任医师"
        cmbLevel.ListIndex = 0
        If UserSignLevel >= cprSL_主治 Then cmbLevel.ListIndex = 1
        If UserSignLevel >= cprSL_主任 Then cmbLevel.ListIndex = 2
        If UserSignLevel >= cprSL_正高 Then cmbLevel.ListIndex = 3
    End Select
    
    '读取当前签名方式（系统参数26）
    Dim lS As Long
    Select Case bytFileKind
    Case cpr门诊病历
        lS = 1
    Case cpr住院病历
        lS = 2
    Case cpr诊疗报告
        lS = 3
    Case cpr护理病历
        lS = 4
    Case Else
        Select Case bytPatiSource
        Case cprPF_门诊
            lS = 1
        Case cprPF_住院
            lS = 2
        Case Else
            lS = 2  '否则，以住院为准
        End Select
    End Select
    
    lngPassType = Val(Mid(zlDatabase.GetPara(26, glngSys), lS, 1)) '门诊,住院,医技,护理 (1111),为空默认采用密码模式
    lblUserName.Caption = UserInfo.姓名
    
    chkEsign.Value = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "chkEsign", vbUnchecked)
    chkHandSign.Value = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "chkHandSign", vbUnchecked)
    
    Dim intFormat As Integer
    intFormat = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "cboTime", 0))
    If intFormat >= 0 And intFormat < Me.cboTime.ListCount Then Me.cboTime.ListIndex = intFormat
    
    Call RefControls
    
    Me.Show vbModal, frmParent
    If mblnOK Then
        Set ShowMe = Sign
    Else
        Set ShowMe = Nothing
    End If
End Function

Public Function GetUserSignLevel(lngUserID As Long, Optional strUserName As String, _
    Optional lngPatiID As Long, Optional lngPatiPageID As Long) As EPRSignLevelEnum
    
    Dim rs As ADODB.Recordset, strSQL As String
    Dim lngR As Long, lngLevel1 As Long, lngLevel2 As Long
    
    err = 0: On Error GoTo ErrHand
    strSQL = "Select g.功能" & vbNewLine & _
            "From zlRoleGrant g, Sys.Dba_Role_Privs r, 上机人员表 p" & vbNewLine & _
            "Where r.Grantee = p.用户名 And g.角色 = r.Granted_Role And g.系统 = [2] And g.序号 = [3] And g.功能 = [4] And" & vbNewLine & _
            "      p.人员id = [1]" & vbNewLine & _
            "Union" & vbNewLine & _
            "Select [4] As 功能 From 上机人员表 p Where 用户名 = [5] And p.人员id = [1]"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngUserID, glngSys, 1070, "签名权", UserInfo.用户名)
    If rs.RecordCount <= 0 Then GetUserSignLevel = cprSL_空白: Exit Function
    
    strSQL = "select 聘任技术职务,签名 from 人员表 p where ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngUserID)
    If Not rs.EOF Then
        lngR = Nvl(rs("聘任技术职务"), 0)
        If mintSign = 1 Then mstrPatiSign = "" & rs!签名
    End If
    Select Case lngR    '1 正高  2 副高  3 中级  4 助理/师级  5 员/士  9 待聘
    Case 1: lngLevel1 = cprSL_正高
    Case 2: lngLevel1 = cprSL_主任
    Case 3: lngLevel1 = cprSL_主治
    Case Else: lngLevel1 = cprSL_经治
    End Select
    rs.Close
    
    If lngPatiID > 0 Then
        strSQL = "Select 经治医师, 主治医师, 主任医师 " & _
            " From 病人变动记录 " & _
            " Where 病人ID = [1] And 主页ID = [2] And (终止时间 Is Null Or 终止原因 = 1) " & _
            "       And 开始时间 Is Not Null And Nvl(附加床位, 0) = 0"
        Set rs = zlDatabase.OpenSQLRecord(strSQL, "cEPRDocument", lngPatiID, lngPatiPageID)
        If rs.EOF Then
            lngLevel2 = cprSL_经治
        Else
            If rs.Fields("主任医师") = IIf(strUserName = "", UserInfo.姓名, strUserName) Then
                lngLevel2 = cprSL_主任
            ElseIf rs.Fields("主治医师") = IIf(strUserName = "", UserInfo.姓名, strUserName) Then
                lngLevel2 = cprSL_主治
            Else
                lngLevel2 = cprSL_经治
            End If
        End If
    End If
    GetUserSignLevel = IIf(lngLevel1 >= lngLevel2, lngLevel1, lngLevel2)
    Exit Function

ErrHand:
    GetUserSignLevel = cprSL_空白
End Function

'################################################################################################################
'## 功能：  保存签名到内部签名组并刷新显示（验证密码或者数字签名）
'################################################################################################################
Private Function Validation() As Boolean
    On Error GoTo LL
    Dim strUserName As String, lngUserID As Long, strSign As String, str时间戳 As String
    Dim SignLevel As EPRSignLevelEnum
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    txtName = Trim(txtName)
    txtPass = Trim(txtPass)
    strUserName = ""
    
    If optName(0).Value Then
        If chkEsign.Value = vbChecked Then
            '数字签名
            err.Clear: On Error Resume Next
            If objESign Is Nothing Then
                Set objESign = CreateObject("zl9ESign.clsESign")
                If err <> 0 Then err = 0: strSign = ""
            End If
            If Not objESign Is Nothing Then
                Call objESign.Initialize(gcnOracle, glngSys)
            End If
            lngCertID = 0
            strSign = objESign.signature(mstrSource, UCase(gcnOracle.Properties(23)), lngCertID, str时间戳) '返回签名信息,lngCertID返回签名使用的证书记录ID
            If strSign = "" Then
                MsgBox "验证失败！请重新输入验证信息！", vbInformation + vbOKOnly, "书写签名"
                GoTo LL
            End If
        End If
        strUserName = IIf(mstrPatiSign = "", UserInfo.姓名, mstrPatiSign)
        lngUserID = UserInfo.ID
    Else
        If chkEsign.Value = vbUnchecked Then
            '密码签名
            If Not CreateRegister() Then
                GoTo LL
            End If
            If Not mobjRegister.LoginValidate(mobjRegister.GetServerName, txtName.Text, txtPass.Text, "") Then
                Validation = False
                MsgBox "验证失败！请重新输入验证信息！", vbInformation + vbOKOnly, "书写签名"
                GoTo LL
            End If
        ElseIf chkEsign.Value = vbChecked Then
            '数字签名
            err.Clear: On Error Resume Next
            If objESign Is Nothing Then
                Set objESign = CreateObject("zl9ESign.clsESign")
                If err <> 0 Then err = 0: strSign = ""
            End If
            If Not objESign Is Nothing Then
                Call objESign.Initialize(gcnOracle, glngSys)
            End If
            lngCertID = 0
            strSign = objESign.signature(mstrSource, UCase(txtName), lngCertID, str时间戳) '返回签名信息,lngCertID返回签名使用的证书记录ID
            If strSign = "" Then
                MsgBox "验证失败！请重新输入验证信息！", vbInformation + vbOKOnly, "书写签名"
                GoTo LL
            End If
        End If
        
        strSQL = "Select ID,姓名,签名 From 人员表 p Where ID=(Select 人员ID From 上机人员表 Where 用户名=[1])"
        On Error GoTo errH
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "Sign-GetUserInfo", UCase(txtName.Text))
        If Not rsTemp.EOF Then
            If mintSign = 1 Then strUserName = "" & rsTemp!签名
            If strUserName = "" Then strUserName = rsTemp.Fields("姓名")
            lngUserID = rsTemp.Fields("ID")     '用户ID
        End If
    End If
    SignLevel = GetUserSignLevel(lngUserID, strUserName, mlngPatiID, mlngPatiPageID) '获取指定用户的签名级别
    
    If SignLevel < cprSL_主任 And SignLevel < Val(cmbLevel.Text) Then
        MsgBox "指定用户签名级别不够！请重新输入验证信息！", vbInformation, gstrSysName
        GoTo LL
    End If
    
    Set Sign = New cEPRSign
    
    Sign.姓名 = strUserName
    Sign.签名级别 = Val(cmbLevel.Text)
    If Sign.签名级别 > cprSL_主任 Then Sign.签名级别 = cprSL_主任
    If Me.chkPreText.Value = vbChecked Then
        Sign.前置文字 = Trim(Mid(Me.cmbLevel.Text, 4)) & "："
    Else
        Sign.前置文字 = ""
    End If
    Sign.签名信息 = strSign   '数字签名的签名信息存储到“要素值域”字段中！
    Sign.显示手签 = (chkHandSign.Value = vbChecked)
    Sign.签名方式 = IIf(chkEsign.Value = vbUnchecked, 1, 2)
    Sign.签名规则 = 1
    Sign.证书ID = IIf(Sign.签名方式 = 2, lngCertID, 0)
    Sign.签名时间 = zlDatabase.Currentdate()
    Select Case Me.cboTime.ListIndex
    Case 1: Sign.显示时间 = "yyyy-MM-dd hh:mm"
    Case 2: Sign.显示时间 = "yyyy年MM月dd日 hh:mm"
    Case Else: Sign.显示时间 = ""
    End Select
    Sign.时间戳 = str时间戳
    
    Validation = True
    Exit Function

LL:
    err = 0: On Error Resume Next
    If txtName.Enabled And txtName.Visible Then
        txtName.SetFocus
    ElseIf txtPass.Enabled And txtPass.Visible Then
        txtPass.SetFocus
    Else
        optName(0).SetFocus
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


'################################################################################################################
'## 功能：  刷新控件
'################################################################################################################
Private Sub RefControls()
    If optName(0).Value Then
        txtName.Enabled = False
        txtPass.Enabled = False
        Select Case lngPassType
        Case 0
            '密码签名
            chkEsign.Value = vbUnchecked
            chkEsign.Visible = False
        Case 1
            '1－数字
            chkEsign.Value = vbChecked
            chkEsign.Left = txtPass.Left
            Me.Label2.Visible = False
            chkEsign.Visible = True
            chkEsign.Enabled = False
            txtPass.Visible = False
        Case 2
            '2－两者皆可
        End Select
    Else
        chkEsign.Enabled = True
        txtPass.Enabled = True
        txtName.Enabled = True
        Select Case lngPassType
        Case 0
            '密码签名
            chkEsign.Value = vbUnchecked
            txtPass.Enabled = True
        Case 1
            '1－数字
            chkEsign.Value = vbChecked
            chkEsign.Left = txtPass.Left
            Me.Label2.Visible = False
            chkEsign.Visible = True
            chkEsign.Enabled = False
            txtPass.Visible = False
        Case 2
            '2－两者皆可
            txtPass.Enabled = (chkEsign.Value = vbUnchecked)
        End Select
    End If
End Sub

Private Sub cboTime_Click()
     Call Preview
End Sub

Private Sub cboTime_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call ZLCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub chkEsign_Click()
    txtPass.Enabled = (chkEsign.Value = vbUnchecked)
    txtPass.Enabled = IIf(optName(0).Value, False, txtPass.Enabled)
    If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
End Sub

Private Sub chkEsign_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call ZLCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub chkHandSign_Click()
     Call Preview
End Sub

Private Sub chkHandSign_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call ZLCommFun.PressKey(vbKeyTab): Call Preview: Exit Sub
End Sub

Private Sub chkPreText_Click()
    Call Preview
End Sub

Private Sub chkPreText_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call ZLCommFun.PressKey(vbKeyTab): Call Preview: Exit Sub
End Sub

Private Sub cmbLevel_Click()
    Call Preview
End Sub

Private Sub cmbLevel_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call ZLCommFun.PressKey(vbKeyTab): Call Preview: Exit Sub
    If InStr("'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Validation Then
        mblnOK = True
        Unload Me
    End If
End Sub

Private Sub Preview()
    Dim strText As String, bln手签 As Boolean, str前置文字 As String
    If Me.chkPreText.Value = vbChecked Then
        str前置文字 = Trim(Mid(Me.cmbLevel.Text, 4)) & "："
    Else
        str前置文字 = ""
    End If
    bln手签 = (chkHandSign.Value = vbChecked)
    strText = str前置文字 & IIf(mstrPatiSign = "", UserInfo.姓名, mstrPatiSign) & IIf(bln手签, "，手签：_____________", "")
    If Me.cboTime.ListIndex > 0 Then
        strText = strText & "，" & Me.cboTime.Text
    End If
    lblPreview.Caption = strText
    
End Sub

Private Sub Form_Activate()
    If Me.Tag = "" Then
        Me.Tag = "1st."
        Me.cmbLevel.SetFocus
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If lngPassType = 2 Then
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "chkEsign", chkEsign.Value
    End If
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "chkHandSign", chkHandSign.Value
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "cboTime", chkHandSign.Value
    Set objESign = Nothing
    Set mobjRegister = Nothing
End Sub

Private Sub optName_Click(Index As Integer)
    Call RefControls
    If Index = 1 Then
        If txtName.Enabled And txtName.Visible Then txtName.SetFocus
    End If
End Sub

Private Sub optName_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call ZLCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub optPassType_Click(Index As Integer)
    If Index = 1 Then
        txtPass.Enabled = True
        If txtPass.Enabled And txtPass.Visible Then zlControl.TxtSelAll txtPass: txtPass.SetFocus
    Else
        txtPass.Enabled = False
    End If
End Sub

Private Sub optPassType_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call ZLCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub txtName_GotFocus()
    zlControl.TxtSelAll txtName
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If chkEsign.Value = vbUnchecked Then
            If txtPass.Enabled And txtPass.Visible Then zlControl.TxtSelAll txtPass: txtPass.SetFocus: Call Preview: Exit Sub
        Else
            Call ZLCommFun.PressKey(vbKeyTab): Call Preview: Exit Sub
        End If
    End If
    If InStr(" ~!@#$%^&*()[]{}_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtNames_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call ZLCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()[]{}_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtPass_GotFocus()
    zlControl.TxtSelAll txtPass
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call ZLCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()[]{}_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Function CreateRegister() As Boolean
    '创建注册部件
    If Not mobjRegister Is Nothing Then CreateRegister = True: Exit Function
    On Error Resume Next
    Set mobjRegister = CreateObject("zlRegister.clsRegister")
    If mobjRegister Is Nothing Then
        err.Clear
        MsgBox "创建zlRegister部件对象失败,请检查文件是否存在并且正确注册。", vbExclamation, gstrSysName
        Exit Function
    End If
    CreateRegister = True
End Function
