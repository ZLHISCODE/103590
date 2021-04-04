VERSION 5.00
Begin VB.Form frmEPRSign 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "书写签名"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   5700
   Icon            =   "frmEPRSign.frx":0000
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
      Top             =   1950
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
Attribute VB_Name = "frmEPRSign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private frmParent As Object                 '父窗体
Private Sign As cEPRSign                    '签名对象

Private objESign As Object                  '电子签名接口部件
Private lngCertID As Long                   '证书ID
Private lngPassType As Long                 '密码验证规则（系统参数） 0-密码；1－数字；2－两者皆可
Private mblnOK As Boolean

Private strSource As String                 '数字签名的源字符串
Private UserSignLevel As EPRSignLevelEnum   '当前用户的签名级别

'################################################################################################################
'## 功能：  显示本窗体
'##
'## 参数：  edtThis     :IN     编辑器控件
'##         fParent     :IN     父窗体
'##         strSource   :IN     数字签名的源字符串（从文本中提取，去掉签名提纲）
'################################################################################################################
Public Function ShowMe(ByRef edtThis As Editor, _
    ByRef fParent As Object, _
    ByVal sSource As String) As cEPRSign
    mblnOK=False
    Dim bytFileKind As Byte    '是否护理病历
    bytFileKind = fParent.Document.EPRPatiRecInfo.病历种类
    
    Dim lngStart As Long, strPreText As String
    
    Me.cboTime.Clear
    Me.cboTime.AddItem "不显示"
    Me.cboTime.AddItem Format(Now(), "yyyy-MM-dd hh:mm")
    Me.cboTime.AddItem Format(Now(), "yyyy年MM月dd日 hh:mm")
    
    lngStart = edtThis.Selection.StartPos
    strPreText = edtThis.Range(lngStart - 1, lngStart)
    If strPreText = ":" Or strPreText = "：" Then
        Me.chkPreText.Value = vbUnchecked
    Else
        Me.chkPreText.Value = vbChecked
    End If
    
    Set Sign = New cEPRSign
    Set frmParent = fParent
    strSource = sSource
    
    UserSignLevel = GetUserSignLevel(glngUserId, , frmParent.Document.EPRPatiRecInfo.病人ID, frmParent.Document.EPRPatiRecInfo.主页ID)  '获取用户签名级别
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
    Select Case fParent.Document.EPRPatiRecInfo.病历种类
    Case cpr门诊病历
        lS = 1
    Case cpr住院病历
        lS = 2
    Case cpr诊疗报告
        lS = 3
    Case cpr护理病历
        lS = 4
    Case Else
        Select Case fParent.Document.EPRPatiRecInfo.病人来源
        Case cprPF_门诊
            lS = 1
        Case cprPF_住院
            lS = 2
        Case Else
            lS = 2  '否则，以住院为准
        End Select
    End Select
    
    lngPassType = Val(Mid(zlDatabase.GetPara(26, glngSys), lS, 1)) '门诊,住院,医技,护理 (1111),为空默认采用密码模式
    lblUserName.Caption = gstrUserName
    
    chkEsign.Value = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "chkEsign", vbUnchecked)
    chkHandSign.Value = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "chkHandSign", vbUnchecked)
    chkPreText.Value = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "chkPreText", vbUnchecked)
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

'################################################################################################################
'## 功能：  保存签名到内部签名组并刷新显示（验证密码或者数字签名）
'################################################################################################################
Private Function Validation() As Boolean
    On Error GoTo LL
    Dim strUserName As String, lngUserID As Long, strSign As String, str时间戳 As String
    Dim SignLevel As EPRSignLevelEnum
    
    txtName = Trim(txtName)
    txtPass = Trim(txtPass)
    strUserName = ""
    
    If optName(0).Value Then
        If chkEsign.Value = vbUnchecked Then
            '密码签名
            strUserName = gstrUserName
            lngUserID = glngUserId
        ElseIf chkEsign.Value = vbChecked Then
            '数字签名
            Err.Clear: On Error Resume Next
            If objESign Is Nothing Then
                Set objESign = CreateObject("zl9ESign.clsESign")
                If Err <> 0 Then Err = 0: strSign = ""
            End If
            If Not objESign Is Nothing Then
                If objESign.Initialize(gcnOracle, glngSys) = False Then Exit Function
            End If
            If objESign.CheckCertificate(UCase(UserInfo.用户名)) = False Then Exit Function

            lngCertID = 0
            If Not objESign.CertificateStoped(UserInfo.姓名) Then
				strSign = objESign.signature(strSource, UCase(gcnOracle.Properties(23)), lngCertID, str时间戳) '返回签名信息,lngCertID返回签名使用的证书记录ID
				If strSign = "" Then
					MsgBox "验证失败！请重新输入验证信息！", vbInformation + vbOKOnly, "书写签名"
					GoTo LL
				End If
            Else
                chkEsign.Value = Unchecked
            End If

            strUserName = gstrUserName
            lngUserID = glngUserId
        End If
        SignLevel = GetUserSignLevel(lngUserID, , frmParent.Document.EPRPatiRecInfo.病人ID, frmParent.Document.EPRPatiRecInfo.主页ID) '获取指定用户的签名级别
    Else
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

        If chkEsign.Value = vbUnchecked Then
            '密码签名
            If Not OraDataOpen(txtName, IIf(UCase(txtName) = "SYS" Or UCase(txtName) = "SYSTEM", txtPass, TranPasswd(txtPass))) Then
                Validation = False
                MsgBox "验证失败！请重新输入验证信息！", vbInformation + vbOKOnly, "书写签名"
                GoTo LL
            End If
        ElseIf chkEsign.Value = vbChecked Then
            '数字签名
            Err.Clear: On Error Resume Next
            If objESign Is Nothing Then
                Set objESign = CreateObject("zl9ESign.clsESign")
                If Err <> 0 Then Err = 0: strSign = ""
            End If
            If Not objESign Is Nothing Then
                If objESign.Initialize(gcnOracle, glngSys) = False Then Exit Function
            End If
            If objESign.CheckCertificate(UCase(txtName)) = False Then Exit Function

            lngCertID = 0
            If Not objESign.CertificateStoped(strUserName) Then
				strSign = objESign.signature(strSource, UCase(txtName), lngCertID, str时间戳) '返回签名信息,lngCertID返回签名使用的证书记录ID
				If strSign = "" Then
					MsgBox "验证失败！请重新输入验证信息！", vbInformation + vbOKOnly, "书写签名"
					GoTo LL
				End If
			 Else
                chkEsign.Value = Unchecked: Exit Function '需要验证密码
            End If
        End If
    End If
    
    If SignLevel < cprSL_主任 And SignLevel < Val(cmbLevel.Text) Then
        MsgBox "指定用户签名级别不够！请重新输入验证信息！", vbInformation, gstrSysName
        GoTo LL
    End If
    
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
'## 功能：  验证用户名密码是否正确
'################################################################################################################
Private Function OraDataOpen(ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
    Dim strSQL As String
    Dim strError As String
    Dim Cn As New ADODB.Connection
    
    On Error Resume Next
    Err = 0
    With Cn
        If .State = adStateOpen Then .Close
'        .Provider = "MSDataShape"
        .Open gcnOracle.ConnectionString, strUserName, strUserPwd
        If Err <> 0 Then
            OraDataOpen = False
            Exit Function
        End If
        .Close
    End With
    Set Cn = Nothing
    OraDataOpen = True
    Exit Function
errHand:
    Set Cn = Nothing
    OraDataOpen = False
    Err = 0
End Function

'################################################################################################################
'## 功能：  密码转换函数
'##
'## 参数：  strOld  :原密码
'##
'## 返回：  加密生成的密码
'################################################################################################################
Public Function TranPasswd(strOld As String) As String
    Dim iBit As Integer, strBit As String
    Dim strNew As String
    If Len(Trim(strOld)) = 0 Then TranPasswd = "": Exit Function
    strNew = ""
    For iBit = 1 To Len(Trim(strOld))
        strBit = UCase(Mid(Trim(strOld), iBit, 1))
        Select Case (iBit Mod 3)
        Case 1
            strNew = strNew & _
                Switch(strBit = "0", "W", strBit = "1", "I", strBit = "2", "N", strBit = "3", "T", strBit = "4", "E", strBit = "5", "R", strBit = "6", "P", strBit = "7", "L", strBit = "8", "U", strBit = "9", "M", _
                   strBit = "A", "H", strBit = "B", "T", strBit = "C", "I", strBit = "D", "O", strBit = "E", "K", strBit = "F", "V", strBit = "G", "A", strBit = "H", "N", strBit = "I", "F", strBit = "J", "J", _
                   strBit = "K", "B", strBit = "L", "U", strBit = "M", "Y", strBit = "N", "G", strBit = "O", "P", strBit = "P", "W", strBit = "Q", "R", strBit = "R", "M", strBit = "S", "E", strBit = "T", "S", _
                   strBit = "U", "T", strBit = "V", "Q", strBit = "W", "L", strBit = "X", "Z", strBit = "Y", "C", strBit = "Z", "X", True, strBit)
        Case 2
            strNew = strNew & _
                Switch(strBit = "0", "7", strBit = "1", "M", strBit = "2", "3", strBit = "3", "A", strBit = "4", "N", strBit = "5", "F", strBit = "6", "O", strBit = "7", "4", strBit = "8", "K", strBit = "9", "Y", _
                   strBit = "A", "6", strBit = "B", "J", strBit = "C", "H", strBit = "D", "9", strBit = "E", "G", strBit = "F", "E", strBit = "G", "Q", strBit = "H", "1", strBit = "I", "T", strBit = "J", "C", _
                   strBit = "K", "U", strBit = "L", "P", strBit = "M", "B", strBit = "N", "Z", strBit = "O", "0", strBit = "P", "V", strBit = "Q", "I", strBit = "R", "W", strBit = "S", "X", strBit = "T", "L", _
                   strBit = "U", "5", strBit = "V", "R", strBit = "W", "D", strBit = "X", "2", strBit = "Y", "S", strBit = "Z", "8", True, strBit)
        Case 0
            strNew = strNew & _
                Switch(strBit = "0", "6", strBit = "1", "J", strBit = "2", "H", strBit = "3", "9", strBit = "4", "G", strBit = "5", "E", strBit = "6", "Q", strBit = "7", "1", strBit = "8", "X", strBit = "9", "L", _
                   strBit = "A", "S", strBit = "B", "8", strBit = "C", "5", strBit = "D", "R", strBit = "E", "7", strBit = "F", "M", strBit = "G", "3", strBit = "H", "A", strBit = "I", "N", strBit = "J", "F", _
                   strBit = "K", "O", strBit = "L", "4", strBit = "M", "K", strBit = "N", "Y", strBit = "O", "D", strBit = "P", "2", strBit = "Q", "T", strBit = "R", "C", strBit = "S", "U", strBit = "T", "P", _
                   strBit = "U", "B", strBit = "V", "Z", strBit = "W", "0", strBit = "X", "V", strBit = "Y", "I", strBit = "Z", "W", True, strBit)
        End Select
    Next
    TranPasswd = strNew
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
    strText = str前置文字 & gstrUserName & IIf(bln手签, "，手签：_____________", "")
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
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "cboTime", cboTime.ListIndex
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "chkPreText", chkPreText.Value
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "cmbLevel", cmbLevel.ListIndex
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
