VERSION 5.00
Begin VB.Form frmCaseTendSign 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "签名"
   ClientHeight    =   2415
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5295
   Icon            =   "frmCaseTendSign.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cmbLevel 
      Height          =   300
      Left            =   1275
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   75
      Width           =   3765
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2850
      TabIndex        =   8
      Top             =   1875
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3990
      TabIndex        =   7
      Top             =   1875
      Width           =   1095
   End
   Begin VB.OptionButton optName 
      Caption         =   "当前用户(&U)"
      Height          =   195
      Index           =   0
      Left            =   165
      TabIndex        =   6
      Top             =   645
      Value           =   -1  'True
      Width           =   1320
   End
   Begin VB.OptionButton optName 
      Caption         =   "指定用户(&S)"
      Height          =   195
      Index           =   1
      Left            =   165
      TabIndex        =   5
      Top             =   1005
      Width           =   1320
   End
   Begin VB.TextBox txtName 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1530
      MaxLength       =   50
      TabIndex        =   4
      Top             =   945
      Width           =   3480
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   -345
      TabIndex        =   3
      Top             =   1770
      Width           =   5670
   End
   Begin VB.TextBox txtPass 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1530
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1290
      Width           =   1995
   End
   Begin VB.CheckBox chkEsign 
      Caption         =   "数字签名(&E)"
      Height          =   195
      Left            =   3750
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1365
      Width           =   1365
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   0
      Left            =   -345
      TabIndex        =   0
      Top             =   495
      Width           =   5805
   End
   Begin VB.Label lblLevel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "签名级别(&L)"
      Height          =   180
      Left            =   165
      TabIndex        =   12
      Top             =   135
      Width           =   990
   End
   Begin VB.Label lblUserName 
      AutoSize        =   -1  'True
      Caption         =   "张三"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1530
      TabIndex        =   11
      Top             =   645
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "密    码(&P)"
      Height          =   180
      Left            =   435
      TabIndex        =   10
      Top             =   1350
      Width           =   990
   End
End
Attribute VB_Name = "frmCaseTendSign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'######################################################################################################################

Private frmParent As Object                 '父窗体
Private Sign As cEPRSign                    '签名对象

Private lngCertID As Long                   '证书ID
Private mlngPassType As Long                 '密码验证规则（系统参数） 0-密码；1－数字；2－两者皆可
Private mblnOk As Boolean

Private strSource As String                 '数字签名的源字符串
Private UserSignLevel As EPRSignLevelEnum   '当前用户的签名级别
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mlngUnitID As Long                  '当前操作的病区ID
Private mstrSource As String                '签名原内容
Private mstr状态 As String
Private mstrPrivs As String

'######################################################################################################################

'电子签名使用场合：
'26  电子签名使用场合(4位字符) 对不同场合是否使用电子签名进行控制,数字位数分别为:门诊,住院,医技,护理 0-不控制,1-控制

Public Function ShowMe(ByRef objParent As Object, ByVal strPrivs As String, ByVal sSource As String, _
    ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lngUnitID As Long, Optional str状态 As String) As cEPRSign
    '******************************************************************************************************************
    '功能： 显示签名窗体
    '参数： edtThis     :IN     编辑器控件
    '       fParent     :IN     父窗体
    '       strSource   :IN     数字签名的源字符串（从文本中提取，去掉签名提纲）
    '******************************************************************************************************************
    
    Set Sign = New cEPRSign
    Set frmParent = objParent
    strSource = sSource
    mstrPrivs = strPrivs
    
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    mlngUnitID = lngUnitID
    mstr状态 = str状态
    UserSignLevel = GetUserSignLevel(glngUserId, , mlng病人ID, mlng主页ID)  '获取用户签名级别
    
    '根据签名级别来初始化“签名级别＂
    cmbLevel.AddItem "1 - 护士"
    cmbLevel.AddItem "3 - 护士长"
    cmbLevel.ListIndex = 0
    If UserSignLevel >= cprSL_主任 Then cmbLevel.ListIndex = 1
    
    
    '读取当前签名方式（系统参数26）
    'mlngPassType = Val(Mid(zlDatabase.GetPara(26, glngSys), 4, 1))     '门诊,住院,医技,护理 (1111),为空默认采用密码模式
    lblUserName.Caption = gstrUserName
    
    Call RefControls
    
    If mstr状态 <> "" Then
        Call cmdOk_Click
    Else
        Me.Show vbModal, frmParent
    End If
    
    If mblnOk Then
        str状态 = mstr状态
        Set ShowMe = Sign
    Else
        Set ShowMe = Nothing
    End If
End Function

Private Function Validation() As Boolean
    '******************************************************************************************************************
    '
    '功能：  保存签名到内部签名组并刷新显示（验证密码或者数字签名）
    '
    '******************************************************************************************************************
    On Error GoTo LL
    Dim strUserName As String, lngUserID As Long, strSign As String, str时间戳 As String, str时间戳信息 As String
    Dim SignLevel As EPRSignLevelEnum, strErr As String
    
    txtName = Trim(txtName)
    txtPass = Trim(txtPass)
    strUserName = ""
    
    If optName(0).Value Then
        '--------------------------------------------------------------------------------------------------------------
        If chkEsign.Value = vbUnchecked Then
            '密码签名
            strUserName = gstrUserName
            lngUserID = glngUserId
        ElseIf chkEsign.Value = vbChecked Then
            '数字签名
            Err.Clear
            If gobjTendESign Is Nothing Then
                On Error Resume Next
                Set gobjTendESign = CreateObject("zl9ESign.clsESign")
                If Err <> 0 Then Err.Clear: strSign = ""
                On Error GoTo 0
                If Not gobjTendESign Is Nothing Then
                    Call gobjTendESign.Initialize(gcnOracle, glngSys)
                End If
            End If
            If gobjTendESign Is Nothing Then
                MsgBox "电子签名部件未能正确安装，签名操作不能继续！", vbInformation, gstrSysName
                GoTo LL
            End If
            lngCertID = 0
            strSign = gobjTendESign.signature(strSource, UCase(gcnOracle.Properties(23)), lngCertID, str时间戳, , str时间戳信息) '返回签名信息,lngCertID返回签名使用的证书记录ID
            If strSign = "" Then
                MsgBox "验证失败！请重新输入验证信息！", vbInformation + vbOKOnly, "书写签名"
                GoTo LL
            End If
            strUserName = gstrUserName
            lngUserID = glngUserId
        End If
        SignLevel = GetUserSignLevel(lngUserID, , mlng病人ID, mlng主页ID) '获取指定用户的签名级别
    Else
        '--------------------------------------------------------------------------------------------------------------
        If chkEsign.Value = vbUnchecked Then
            '密码签名
            If gobjRegister Is Nothing Then Set gobjRegister = DynamicCreate("zlRegister.clsRegister", "密码验证组件")
            If Not gobjRegister.LoginValidate("", txtName, txtPass, strErr) Then
                Validation = False
                MsgBox "验证失败！请重新输入验证信息！" & strErr, vbInformation + vbOKOnly, "书写签名"
                GoTo LL
            End If
        ElseIf chkEsign.Value = vbChecked Then
            '数字签名
            Err.Clear
            If gobjTendESign Is Nothing Then
                On Error Resume Next
                Set gobjTendESign = CreateObject("zl9ESign.clsESign")
                If Err <> 0 Then Err.Clear: strSign = ""
                On Error GoTo 0
                If Not gobjTendESign Is Nothing Then
                    Call gobjTendESign.Initialize(gcnOracle, glngSys)
                End If
            End If
            If gobjTendESign Is Nothing Then
                MsgBox "电子签名部件未能正确安装，签名操作不能继续！", vbInformation, gstrSysName
                GoTo LL
            End If
            lngCertID = 0
            strSign = gobjTendESign.signature(strSource, UCase(txtName), lngCertID, str时间戳, , str时间戳信息) '返回签名信息,lngCertID返回签名使用的证书记录ID
            If strSign = "" Then
                MsgBox "验证失败！请重新输入验证信息！", vbInformation + vbOKOnly, "书写签名"
                GoTo LL
            End If
        End If
        
        Dim rsTemp As New ADODB.Recordset
        gstrSQL = "Select ID,姓名 From 人员表 p Where ID=(Select 人员ID From 上机人员表 Where 用户名='" & UCase(txtName) & "') And (p.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or p.撤档时间 Is Null) "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "Sign-GetUserInfo")
        If Not rsTemp.EOF Then
            strUserName = rsTemp.Fields("姓名") '用户姓名
            lngUserID = rsTemp.Fields("ID")     '用户ID
        End If
        rsTemp.Close
        SignLevel = GetUserSignLevel(lngUserID, strUserName, mlng病人ID, mlng主页ID) '获取指定用户的签名级别
    End If
    
    If SignLevel < Val(cmbLevel.Text) Then
        MsgBox "指定用户没有签名权限或骋任职务未达到签名级别！请重新输入验证信息！", vbInformation + vbOKOnly, gstrSysName
        GoTo LL
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    Sign.姓名 = strUserName
    Sign.签名级别 = Val(cmbLevel.Text)
    If Sign.签名级别 > cprSL_主任 Then Sign.签名级别 = cprSL_主任
    Sign.签名信息 = strSign
    Sign.签名方式 = IIf(chkEsign.Value = vbUnchecked, 1, 2)
    Sign.签名规则 = 1
    Sign.证书ID = IIf(Sign.签名方式 = 2, lngCertID, 0)
    Sign.时间戳 = str时间戳
    Sign.时间戳信息 = str时间戳信息
    
    Validation = True
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
LL:
    Err = 0: On Error Resume Next
    If txtName.Enabled And txtName.Visible Then
        txtName.SetFocus
    ElseIf txtPass.Enabled And txtPass.Visible Then
        txtPass.SetFocus
    Else
        optName(0).SetFocus
    End If
End Function

'################################################################################################################
'## 功能：  刷新控件
'################################################################################################################
Private Sub RefControls()
    Dim arrData
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    '63955:刘鹏飞,2013-09-16,启用点击签名，并且当前签名的病区在设置的电子签名启用部门中才能使用电子签名
    '说明：如果没有设置电子签名所要启用的部门,就说明启用电子签名的病区为所有病区
    
    If mstr状态 <> "" And InStr(1, mstr状态, "|") <> 0 Then
        arrData = Split(mstr状态, "|")
        cmbLevel.ListIndex = Val(arrData(0))
        optName(0).Value = arrData(1)
        optName(1).Value = arrData(2)
        txtName.Text = arrData(3)
        txtPass.Text = arrData(4)
        mlngPassType = Val(arrData(5))
    Else
        gstrSQL = "Select Zl_Fun_Getsignpar([1],[2]) 电子签名 From dual"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "电子签名启用部门", 4, mlngUnitID)
        If rsTemp.RecordCount > 0 Then
            mlngPassType = Val(NVL(rsTemp!电子签名, 0))
        Else
            mlngPassType = 0
        End If
        If mlngPassType = 1 Then
            If CertificateStoped(gstrUserName) = True Then mlngPassType = 0
        End If
    End If

    
    If optName(0).Value Then
        txtName.Enabled = False
        txtPass.Enabled = False
        Select Case mlngPassType
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
        Select Case mlngPassType
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
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function CertificateStoped(ByVal strName As String) As Boolean
'功能：检查签名人的证书是否停用，停用的话将不使用电子签名
    On Error Resume Next
    CertificateStoped = True
    Err.Clear
    If gobjTendESign Is Nothing Then
        Set gobjTendESign = CreateObject("zl9ESign.clsESign")
        If Err <> 0 Then Err.Clear
        If Not gobjTendESign Is Nothing Then Call gobjTendESign.Initialize(gcnOracle, glngSys)
    End If
    
    If gobjTendESign Is Nothing Then Exit Function
    CertificateStoped = gobjTendESign.CertificateStoped(strName)
    If Err <> 0 Then Err.Clear
End Function

Private Sub chkEsign_Click()
    txtPass.Enabled = (chkEsign.Value = vbUnchecked)
    txtPass.Enabled = IIf(optName(0).Value, False, txtPass.Enabled)
    If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
End Sub

Private Sub chkEsign_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub cmbLevel_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr("'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub cmdCancel_Click()
    mblnOk = False
    Unload Me
End Sub

Private Sub cmdOk_Click()
    If Validation Then
        mstr状态 = cmbLevel.ListIndex & "|" & optName(0).Value & "|" & optName(1).Value & "|" & txtName.Text & "|" & txtPass.Text & "|" & chkEsign.Value
        
        mblnOk = True
        Unload Me
    End If
End Sub

Private Sub Form_Activate()
    If Me.Tag = "" Then
        Me.Tag = "1st."
        Me.cmbLevel.SetFocus
    End If
End Sub

Private Sub optName_Click(Index As Integer)
    Call RefControls
    If Index = 1 Then
        If txtName.Enabled And txtName.Visible Then txtName.SetFocus
    End If
End Sub

Private Sub optName_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
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
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub txtName_GotFocus()
    zlControl.TxtSelAll txtName
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If chkEsign.Value = vbUnchecked Then
            If txtPass.Enabled And txtPass.Visible Then zlControl.TxtSelAll txtPass: txtPass.SetFocus:  Exit Sub
        Else
            Call zlCommFun.PressKey(vbKeyTab):  Exit Sub
        End If
    End If
    If InStr(" ~!@#$%^&*()[]{}_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtNames_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()[]{}_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtPass_GotFocus()
    zlControl.TxtSelAll txtPass
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()[]{}_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub


