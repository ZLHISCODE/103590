VERSION 5.00
Begin VB.Form frmSign 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "书写签名"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   5700
   Icon            =   "frmSign.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cboTime 
      Height          =   300
      Left            =   1290
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2580
      Width           =   2310
   End
   Begin VB.Frame fraLine 
      Height          =   30
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
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   300
      Left            =   1605
      MaxLength       =   50
      TabIndex        =   4
      Top             =   960
      Width           =   3840
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
      Top             =   2340
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4365
      TabIndex        =   12
      Top             =   1935
      Width           =   1095
   End
   Begin VB.ComboBox cmbLevel 
      Height          =   300
      Left            =   1350
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   90
      Width           =   4110
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
      Left            =   240
      TabIndex        =   17
      Top             =   3255
      Width           =   5235
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
      Left            =   240
      TabIndex        =   0
      Top             =   150
      Width           =   990
   End
End
Attribute VB_Name = "frmSign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mSign As cTabSign                     '签名对象
Private fParent As Object                    '签名对象
Private mlngCertID As Long                   '证书ID
Private mblnOK As Boolean
Private UserSignLevel As EPRSignLevel   '当前用户的签名级别

'################################################################################################################
'## 功能：  显示本窗体
'##
'##         fParent     :IN     父窗体
'##         strSource   :IN     数字签名的源字符串（从文本中提取，去掉签名提纲）
'################################################################################################################
Public Function ShowMe(ByVal strSignKey As String, ByVal frmParent As Object) As cTabSign
    On Error GoTo errHand
    mblnOK = False
    Set fParent = frmParent
    Set mSign = fParent.Document.Signs("K" & strSignKey)
    cboTime.AddItem "不显示"
    cboTime.AddItem Format(zlDatabase.Currentdate, "yyyy-MM-dd hh:mm")
    cboTime.AddItem Format(zlDatabase.Currentdate, "yyyy年MM月dd日 hh:mm")
    
    '根据签名级别来初始化“签名级别＂
    UserSignLevel = fParent.Document.签名级别
    Select Case fParent.Document.EPRPatiRecInfo.病历种类
     Case Tab护理病历
        cmbLevel.AddItem "1 - 护士"
        cmbLevel.AddItem "3 - 护士长"
        cmbLevel.ListIndex = 0
        If UserSignLevel >= TabSL_主任 Then cmbLevel.ListIndex = 1
    Case Tab诊疗报告
        cmbLevel.AddItem "1 - 医生"
        cmbLevel.AddItem "2 - 主治"
        cmbLevel.AddItem "3 - 主任"
        cmbLevel.ListIndex = 0
        If UserSignLevel >= TabSL_主治 Then cmbLevel.ListIndex = 1
        If UserSignLevel >= TabSL_主任 Then cmbLevel.ListIndex = 2
    Case Else
        cmbLevel.AddItem "1 - 经治医师"
        cmbLevel.AddItem "2 - 主治医师"
        cmbLevel.AddItem "3 - 副主任医师"
        cmbLevel.AddItem "4 - 主任医师"
        cmbLevel.ListIndex = 0
        If UserSignLevel >= TabSL_主治 Then cmbLevel.ListIndex = 1
        If UserSignLevel >= TabSL_主任 Then cmbLevel.ListIndex = 2
        If UserSignLevel >= TabSL_正高 Then cmbLevel.ListIndex = 3
    End Select
    
    '读取当前签名方式（系统参数26）
    Dim lS As Long
    Select Case fParent.Document.EPRPatiRecInfo.病历种类
      Case Tab门诊病历
        lS = 1
    Case Tab住院病历
        lS = 2
    Case Tab诊疗报告
        lS = 3
    Case Tab护理病历
        lS = 4
    Case Else
        Select Case fParent.Document.EPRPatiRecInfo.病人来源
        Case TabPF_门诊
            lS = 1
        Case TabPF_住院
            lS = 2
        Case Else
            lS = 2  '否则，以住院为准
        End Select
    End Select
        
    lblUserName.Caption = UserInfo.姓名
    
    chkHandSign.Value = GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\" & App.ProductName & "\" & Me.Name, "chkHandSign", vbUnchecked)
    chkPreText.Value = GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\" & App.ProductName & "\" & Me.Name, "chkPreText", vbUnchecked)
    Dim intFormat As Integer
    intFormat = Val(GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\" & App.ProductName & "\" & Me.Name, "cboTime", 0))
    If intFormat >= 0 And intFormat < Me.cboTime.ListCount Then Me.cboTime.ListIndex = intFormat
    
    Call RefControls
    
    Me.Show vbModal, frmParent

    If mblnOK Then
        Set ShowMe = mSign
    Else
        Set ShowMe = Nothing
    End If
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

'################################################################################################################
'## 功能：  保存签名到内部签名组并刷新显示（验证密码或者数字签名）
'################################################################################################################
Private Function Validation() As Boolean
    On Error GoTo LL
    Dim strUserName As String, lngUserID As Long, strSign As String, str时间戳 As String, objSignPic As Object, str时间戳信息 As String
    Dim SignLevel As EPRSignLevel, strSource As String, strFile As String, objFileStream As TextStream, strErr As String
    
    'strSource为需要生成签名的原文文本,由Doc生成
    If chkEsign.Value = vbChecked Then
        Call fParent.Document.BuildXmlFile(strFile, True, True)
        If gobjFSO.FileExists(strFile) = False Then
            MsgBox "请检查程序目录是否有文件写入权限！", vbInformation, gstrSysName
            Exit Function
        End If
        Set objFileStream = gobjFSO.OpenTextFile(strFile, ForReading)
        strSource = objFileStream.ReadAll
        Set objFileStream = Nothing
        gobjFSO.DeleteFile strFile, True
    End If
    
    txtName = Trim(txtName)
    txtPass = Trim(txtPass)
    strUserName = ""
    
    If optName(0).Value Then
        If chkEsign.Value = vbUnchecked Then
            '密码签名
        ElseIf chkEsign.Value = vbChecked Then
            '数字签名
            Err.Clear: On Error Resume Next
            If gobjESign Is Nothing Then
                Set gobjESign = CreateObject("zl9ESign.clsESign")
                If Err <> 0 Then Err = 0: strSign = ""
            End If
            If Not gobjESign Is Nothing Then
                Call gobjESign.Initialize(gcnOracle, glngSys)
            End If
            mlngCertID = 0
            strSign = gobjESign.signature(strSource, UCase(UserInfo.用户名), mlngCertID, str时间戳, objSignPic, str时间戳信息) '返回签名信息,mlngCertID返回签名使用的证书ID
            If strSign = "" Then
                MsgBox "验证失败！请重新输入验证信息！", vbInformation + vbOKOnly, "书写签名"
                GoTo LL
            End If
        End If
        strUserName = UserInfo.姓名
        lngUserID = UserInfo.ID
        SignLevel = CInt(UserSignLevel)
    Else
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
            Err.Clear: On Error Resume Next
            If gobjESign Is Nothing Then
                Set gobjESign = CreateObject("zl9ESign.clsESign")
                If Err <> 0 Then Err = 0: strSign = ""
            End If
            If Not gobjESign Is Nothing Then
                Call gobjESign.Initialize(gcnOracle, glngSys)
            End If
            mlngCertID = 0
            strSign = gobjESign.signature(strSource, UCase(txtName), mlngCertID, str时间戳, objSignPic, str时间戳信息) '返回签名信息,mlngCertID返回签名使用的证书记录ID
            If strSign = "" Then
                MsgBox "验证失败！请重新输入验证信息！", vbInformation + vbOKOnly, "书写签名"
                GoTo LL
            End If
        End If
        
        Dim rsTemp As New ADODB.Recordset
        gstrSQL = "Select b.Id, b.姓名, b.签名" & vbNewLine & _
                    "From 上机人员表 A, 人员表 B" & vbNewLine & _
                    "Where a.用户名 =[1] And a.人员id = b.Id And" & vbNewLine & _
                    "      Nvl(b.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'YYYY-MM-DD')"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "Sign-GetUserInfo", UCase(txtName))
        If Not rsTemp.EOF Then
            strUserName = rsTemp!姓名  '用户姓名
            lngUserID = rsTemp!ID      '用户ID
        End If
        rsTemp.Close
        SignLevel = GetUserSignLevel(lngUserID, strUserName, fParent.Document.EPRPatiRecInfo.病人ID, fParent.Document.EPRPatiRecInfo.主页ID)  '获取指定用户的签名级别
    End If
    
    If SignLevel < TabSL_主任 And SignLevel < Val(cmbLevel.Text) Then
        MsgBox "指定用户签名级别不够！请重新输入验证信息！", vbInformation, gstrSysName
        GoTo LL
    End If
    
    With mSign
        .姓名 = strUserName
        .签名级别 = Val(cmbLevel.Text)
        If .签名级别 > TabSL_主任 Then .签名级别 = TabSL_主任
        If chkPreText.Value = vbChecked Then
            .前置文字 = Trim(Mid(cmbLevel.Text, 4)) & "："
        Else
            .前置文字 = ""
        End If
        .签名信息 = strSign   '数字签名的签名信息存储到“要素值域”字段中！
        .显示手签 = (chkHandSign.Value = vbChecked)
        .签名方式 = IIf(chkEsign.Value = vbUnchecked, 1, 2)
        .签名规则 = 1
        .证书ID = IIf(.签名方式 = 2, mlngCertID, 0) '数字签名
        .时间戳 = str时间戳                         '数字签名
        .签名时间 = zlDatabase.Currentdate()
        Select Case cboTime.ListIndex
        Case 1: .显示时间 = "yyyy-MM-dd hh:mm"
        Case 2: .显示时间 = "yyyy年MM月dd日 hh:mm"
        Case Else: .显示时间 = ""
        End Select
    End With
    
    Validation = True
    Exit Function

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
    If optName(0).Value Then
        txtName.Enabled = False
        txtPass.Enabled = False
        Select Case gbytEsign
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
        Select Case gbytEsign
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
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub chkEsign_Click()
    txtPass.Enabled = (chkEsign.Value = vbUnchecked)
    txtPass.Enabled = IIf(optName(0).Value, False, txtPass.Enabled)
    If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
End Sub

Private Sub chkEsign_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub chkHandSign_Click()
     Call Preview
End Sub

Private Sub chkHandSign_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Call Preview: Exit Sub
End Sub

Private Sub chkPreText_Click()
    Call Preview
End Sub

Private Sub chkPreText_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Call Preview: Exit Sub
End Sub

Private Sub cmbLevel_Click()
    Call Preview
End Sub

Private Sub cmbLevel_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Call Preview: Exit Sub
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
    strText = str前置文字 & UserInfo.姓名 & IIf(bln手签, "，手签：_____________", "")
    If Me.cboTime.ListIndex > 0 Then
        strText = strText & "，" & Me.cboTime.Text
    End If
    lblPreview.Caption = strText
    
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If Me.Tag = "" Then
        Me.Tag = "1st."
        Me.cmbLevel.SetFocus
    End If
    If Err.Number <> 0 Then
        MsgBox Me.Caption & vbCrLf & Err.Description & "   请截图并通知管理员！", vbInformation, gstrSysName
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    SaveSetting "ZLSOFT", "私有模块\" & UserInfo.用户名 & "\" & App.ProductName & "\" & Me.Name, "chkHandSign", chkHandSign.Value
    SaveSetting "ZLSOFT", "私有模块\" & UserInfo.用户名 & "\" & App.ProductName & "\" & Me.Name, "cboTime", cboTime.ListIndex
    SaveSetting "ZLSOFT", "私有模块\" & UserInfo.用户名 & "\" & App.ProductName & "\" & Me.Name, "chkPreText", chkPreText.Value
    SaveSetting "ZLSOFT", "私有模块\" & UserInfo.用户名 & "\" & App.ProductName & "\" & Me.Name, "cmbLevel", cmbLevel.ListIndex
    Set fParent = Nothing
    Err.Clear
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
            If txtPass.Enabled And txtPass.Visible Then zlControl.TxtSelAll txtPass: txtPass.SetFocus: Call Preview: Exit Sub
        Else
            Call zlCommFun.PressKey(vbKeyTab): Call Preview: Exit Sub
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
