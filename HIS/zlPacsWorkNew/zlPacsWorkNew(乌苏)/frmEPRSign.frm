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
   Begin VB.CommandButton cmdOK 
      Caption         =   "审核签名"
      Height          =   350
      Index           =   2
      Left            =   4320
      TabIndex        =   19
      Top             =   2430
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "诊断签名"
      Height          =   350
      Index           =   1
      Left            =   4320
      TabIndex        =   18
      Top             =   2040
      Width           =   1095
   End
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
      TabIndex        =   17
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
      TabStop         =   0   'False
      Top             =   2257
      Width           =   1695
   End
   Begin VB.CheckBox chkEsign 
      Caption         =   "数字签名(&E)"
      Height          =   195
      Left            =   4170
      TabIndex        =   7
      TabStop         =   0   'False
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
      TabIndex        =   14
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
      Left            =   4320
      TabIndex        =   12
      Top             =   2840
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
      TabIndex        =   16
      Top             =   3255
      Width           =   5235
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "签名效果预览:"
      Height          =   180
      Left            =   240
      TabIndex        =   15
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
      TabIndex        =   13
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

Private Sign As cEPRSign                    '签名对象


Private mlngPassType As Long                 '密码验证规则（系统参数） 0-密码；1－数字；2－两者皆可
Private mblnOK As Boolean
Private mlngPatientID As Long
Private mlngPageID As Long
Private mlngReportID As Long                '报告ID
Private mint开始版 As Integer               '本次报告签名的开始版

Private mstrPrivs As String                 '权限字符串

Private UserSignLevel As EPRSignLevelEnum   '当前用户的签名级别

'################################################################################################################
'## 功能：  显示本窗体
'##
'##         fParent     :IN     父窗体
'##         lngMaxSignLevel :IN  本次报告已经有的最大的签名级别
'################################################################################################################
Public Function ShowMe(ByRef fParent As Object, ByVal lngPassType As Long, ByVal lngReportID As Long, ByVal lngPatientID As Long, ByVal lngPageID As Long, _
     ByVal strPrivs As String, ByVal lngMaxSignLevel As Long, int开始版 As Integer) As cEPRSign
    Dim curDate As Date
    
    curDate = zlDatabase.Currentdate
    mlngPatientID = lngPatientID
    mlngPageID = lngPageID
    mstrPrivs = strPrivs
    mlngReportID = lngReportID
    mint开始版 = int开始版
    mlngPassType = lngPassType
    mblnOK = False
    
    Me.cboTime.Clear
    Me.cboTime.AddItem "不显示"
    Me.cboTime.AddItem Format(curDate, "yyyy-MM-dd hh:mm")
    Me.cboTime.AddItem Format(curDate, "yyyy年MM月dd日 hh:mm")
    
    '根据当前报告的签名状态，确定是否显示“诊断签名”按钮，当已经有过一次审核签名后，不再显示诊断签名按钮
    If lngMaxSignLevel > 1 Then
        cmdOK(1).Visible = False
    Else
        cmdOK(1).Visible = True
    End If
    
    '根据权限，确定是否显示“审核签名”按钮
    If InStr(mstrPrivs, "PACS报告修订") > 0 Then
        cmdOK(2).Visible = True
    Else
        cmdOK(2).Visible = False
    End If

    Set Sign = New cEPRSign
    
    UserSignLevel = GetUserSignLevel(UserInfo.ID, , lngPatientID, lngPageID)    '获取用户签名级别
    '根据签名级别来初始化“签名级别＂
    cmbLevel.AddItem "1 - 医生"
    cmbLevel.AddItem "2 - 主治"
    cmbLevel.AddItem "3 - 主任"
    cmbLevel.ListIndex = 0
    If UserSignLevel >= cprSL_主治 Then cmbLevel.ListIndex = 1
    If UserSignLevel >= cprSL_主任 Then cmbLevel.ListIndex = 2
    
    '读取当前签名方式（系统参数26）,诊疗报告是从 3开始
    'lngPassType = Val(Mid(zlDatabase.GetPara(26, glngSys), 7, 1))  '门诊,住院,医技,护理,药品,LIS,PACS (1111111),为空默认采用密码模式
    lblUserName.Caption = UserInfo.姓名
    
    chkEsign.value = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "chkEsign", vbUnchecked)
    chkHandSign.value = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "chkHandSign", vbUnchecked)
    
    Dim intFormat As Integer
    intFormat = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "cboTime", 0))
    If intFormat >= 0 And intFormat < Me.cboTime.ListCount Then Me.cboTime.ListIndex = intFormat
    
    '刷新界面中各个控件的显示状态
    Call RefControls
    
    Me.Show vbModal, fParent
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
    Dim strUserName As String, lngUserID As Long, strSign As String, str时间戳 As String, str时间戳Base64 As String
    Dim objESign As Object                  '电子签名接口部件
    Dim lngCertID As Long                   '证书ID
    Dim SignLevel As EPRSignLevelEnum
    Dim strSource As String     '报告源文
    Dim intRule As Integer      '报告源文组织规则
    Dim strDBUser As String     '传输给签名部件的数据库用户名
    Dim objPic As StdPicture
    
    txtName = Trim(txtName)
    txtPass = Trim(txtPass)
    strUserName = ""
    intRule = 1
    
    '使用当前用户签名
    If optName(0).value Then
        If chkEsign.value = vbChecked Then
            '数字签名
            strDBUser = UCase(gcnOracle.Properties(23))
        End If
        
        strUserName = UserInfo.姓名
        lngUserID = UserInfo.ID
        '获取当前用户的签名级别
        SignLevel = GetUserSignLevel(lngUserID, , mlngPatientID, mlngPageID)
    Else
    '使用指定用户签名
        If chkEsign.value = vbUnchecked Then
            '密码签名
            If Not OraDataOpen(txtName, IIf(UCase(txtName) = "SYS" Or UCase(txtName) = "SYSTEM", txtPass, TranPasswd(txtPass))) Then
                MsgBoxD Me, "验证失败！请重新输入验证信息！", vbInformation + vbOKOnly, "书写签名"
                GoTo LL
            End If
        End If

        strDBUser = UCase(txtName)
        
        '从数据库中获取指定用户签名方式的签名人姓名和ID
        Dim rsTemp As New ADODB.Recordset
        gstrSQL = "Select 姓名,ID From 人员表 p Where ID=(Select 人员ID From 上机人员表 Where 用户名='" & strDBUser & "')"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "Sign-GetUserInfo")
        If Not rsTemp.EOF Then
            strUserName = rsTemp.Fields("姓名") '用户姓名
            lngUserID = rsTemp.Fields("ID")     '用户ID
        End If
        rsTemp.Close
        
        '获取指定用户的签名级别
        SignLevel = GetUserSignLevel(lngUserID, strUserName, mlngPatientID, mlngPageID)
    End If
    
    If SignLevel < cprSL_主任 And SignLevel < Val(cmbLevel.Text) Then
        MsgBoxD Me, "指定用户签名级别不够！请重新输入验证信息！", vbInformation, gstrSysName
        GoTo LL
    End If
    
    
    Sign.姓名 = strUserName
    Sign.签名级别 = Val(cmbLevel.Text)
    If Sign.签名级别 > cprSL_主任 Then Sign.签名级别 = cprSL_主任
    If Me.chkPreText.value = vbChecked Then
        Sign.前置文字 = Trim(Mid(Me.cmbLevel.Text, 4)) & "："
    Else
        Sign.前置文字 = ""
    End If
    
    Sign.显示手签 = (chkHandSign.value = vbChecked)
    Sign.签名方式 = IIf(chkEsign.value = vbUnchecked, 1, 2)
    Sign.签名规则 = intRule
    Sign.签名时间 = zlDatabase.Currentdate()
    Sign.开始版 = mint开始版
    
    Select Case Me.cboTime.ListIndex
        Case 1: Sign.显示时间 = "yyyy-MM-dd hh:mm"
        Case 2: Sign.显示时间 = "yyyy年MM月dd日 hh:mm"
        Case Else: Sign.显示时间 = ""
    End Select
    
    '如果是数字签名，则提取源文，对源文加密
    If chkEsign.value = vbChecked Then
        '创建数字签名部件
        err.Clear: On Error Resume Next
        If objESign Is Nothing Then
            Set objESign = CreateObject("zl9ESign.clsESign")
            If err <> 0 Then err = 0: strSign = ""
        End If
        
        '初始化数字签名部件
        If Not objESign Is Nothing Then
            If objESign.Initialize(gcnOracle, glngSys) = False Then
                MsgBoxD Me, "数字证书初始化失败，请使用正确的数字证书签名。", vbInformation + vbOKOnly, "书写签名"
                GoTo LL
            End If
        End If

        '先检查数字证书跟登陆用户是否一致
        If objESign.CheckCertificate(strDBUser) = False Then
            '当是证书停用时，不使用数字签名对源文进行签名加密，并可以继续签名操作
            If Not objESign.CertificateStoped(UserInfo.姓名) Then
                'Validation = True
                Exit Function
            End If
        Else
            '获取签名的源文
            intRule = GetSignSourceString(1, mlngReportID, mint开始版, False, Sign, strSource)
            If intRule = 0 Then
                '源文提取失败，退出签名
                MsgBoxD Me, "本次报告版本为" & mint开始版 & "的签名源文提取失败，无法签名。", vbInformation + vbOKOnly, "书写签名"
                GoTo LL
            End If
            
            lngCertID = 0
            
            '使用数字签名对源文进行签名加密
            '返回：签名信息strSign-加密后的源文；lngCertID-签名使用的证书记录ID；str时间戳 --签名之后的时间戳
            strSign = objESign.signature(strSource, strDBUser, lngCertID, str时间戳, objPic, str时间戳Base64)
            If strSign = "" Then
                MsgBoxD Me, "验证失败！请重新输入验证信息！", vbInformation + vbOKOnly, "书写签名"
                GoTo LL
            End If
        End If
    End If
     
    '证书ID，是通过签名来返回的，最后记录在签名的证书ID字段，这个字段保存在“对象属性”中，所以对象属性的内容在签名前后会更改，不能作为签名的源文
    Sign.证书ID = IIf(Sign.签名方式 = 2, lngCertID, 0)
    
    '签名信息保存的是通过数字签名加密之后的密文信息
    Sign.签名信息 = strSign   '数字签名的签名信息存储到“要素值域”字段中！
    
    '签名时间戳，
    Sign.时间戳 = str时间戳
    
    '时间戳base64编码
    Sign.时间戳信息 = str时间戳Base64
    
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
End Function

'################################################################################################################
'## 功能：  验证用户名密码是否正确
'################################################################################################################
Private Function OraDataOpen(ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
    Dim strSQL As String
    Dim strError As String
    Dim Cn As New ADODB.Connection
    
    On Error Resume Next
    err = 0
    With Cn
        If .State = adStateOpen Then .Close
'        .Provider = "MSDataShape"
        .Open gcnOracle.ConnectionString, strUserName, strUserPwd
        If err <> 0 Then
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
    err = 0
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
    If optName(0).value Then
        '使用当前用户签名
        txtName.Enabled = False
        txtPass.Enabled = False
        Select Case mlngPassType
        Case 0
            '密码签名
            chkEsign.value = vbUnchecked
            chkEsign.Visible = False
        Case 1
            '1－数字
            chkEsign.value = vbChecked
            chkEsign.Left = txtPass.Left
            Me.Label2.Visible = False
            chkEsign.Visible = True
            chkEsign.Enabled = False
            txtPass.Visible = False
        Case 2
            '2－两者皆可
        End Select
    Else
        '使用指定用户签名
        chkEsign.Enabled = True
        txtPass.Enabled = True
        txtName.Enabled = True
        Select Case mlngPassType
        Case 0
            '密码签名
            chkEsign.value = vbUnchecked
            txtPass.Enabled = True
        Case 1
            '1－数字
            chkEsign.value = vbChecked
            chkEsign.Left = txtPass.Left
            Me.Label2.Visible = False
            chkEsign.Visible = True
            chkEsign.Enabled = False
            txtPass.Visible = False
        Case 2
            '2－两者皆可
            txtPass.Enabled = (chkEsign.value = vbUnchecked)
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
    txtPass.Enabled = (chkEsign.value = vbUnchecked)
    txtPass.Enabled = IIf(optName(0).value, False, txtPass.Enabled)
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

Private Sub Preview()
    Dim strText As String, bln手签 As Boolean, str前置文字 As String
    If Me.chkPreText.value = vbChecked Then
        str前置文字 = Trim(Mid(Me.cmbLevel.Text, 4)) & "："
    Else
        str前置文字 = ""
    End If
    bln手签 = (chkHandSign.value = vbChecked)
    strText = str前置文字 & UserInfo.姓名 & IIf(bln手签, "，手签：_____________", "")
    If Me.cboTime.ListIndex > 0 Then
        strText = strText & "，" & Me.cboTime.Text
    End If
    lblPreview.Caption = strText
    
End Sub

Private Sub cmdOK_Click(Index As Integer)
    '根据签名类型，自动设置签名级别
    If optName(0).value = True Then
        If Index = 1 Then   '诊断签名，自动将签名级别设定为“医生”
            If cmbLevel.ListIndex <> 0 Then cmbLevel.ListIndex = 0
        ElseIf Index = 2 Then   '审核签名，如果当前txtLevel中选择的是“医生”则自动调整为最高的签名级别，否则不修改
            If UserSignLevel < cprSL_主治 Then
                MsgBoxD Me, "您不具备审核签名的聘用职务，请检查。"
                Exit Sub
            End If
            If cmbLevel.ListIndex = 0 Then
                If UserSignLevel >= cprSL_主治 Then cmbLevel.ListIndex = 1
                If UserSignLevel >= cprSL_主任 Then cmbLevel.ListIndex = 2
            End If
        End If
    End If
    
    '检查数据有效性，同时对源文进行加密
    If Validation Then
        mblnOK = True
        Unload Me
    End If
End Sub

Private Sub Form_Activate()
    If Me.Tag = "" Then
        Me.Tag = "1st."
        Me.cmbLevel.SetFocus
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mlngPassType = 2 Then
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "chkEsign", chkEsign.value
    End If
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "chkHandSign", chkHandSign.value
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "cboTime", chkHandSign.value
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
        If chkEsign.value = vbUnchecked Then
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
