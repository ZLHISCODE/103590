VERSION 5.00
Begin VB.Form frmEPRSign 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "书写签名"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   6330
   Icon            =   "frmEPRSign.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   -270
      TabIndex        =   13
      Top             =   1785
      Width           =   6555
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   0
      Left            =   -270
      TabIndex        =   11
      Top             =   510
      Width           =   6555
   End
   Begin VB.CheckBox chkEsign 
      Caption         =   "数字签名(&E)"
      Height          =   195
      Left            =   3105
      TabIndex        =   7
      Top             =   1013
      Width           =   1365
   End
   Begin VB.TextBox txtPass 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1605
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   1387
      Width           =   1365
   End
   Begin VB.TextBox txtName 
      Height          =   300
      Left            =   1605
      MaxLength       =   50
      TabIndex        =   4
      Top             =   960
      Width           =   1365
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
      Left            =   5010
      TabIndex        =   9
      Top             =   1875
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3765
      TabIndex        =   8
      Top             =   1875
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
   Begin VB.PictureBox pic签名图片 
      AutoRedraw      =   -1  'True
      Height          =   810
      Left            =   5265
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   12
      Top             =   690
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "用户密码(&P)"
      Height          =   180
      Left            =   510
      TabIndex        =   5
      Top             =   1440
      Width           =   990
   End
   Begin VB.Label lblUserName 
      AutoSize        =   -1  'True
      Caption         =   "张三"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1605
      TabIndex        =   10
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

Private mlngPassType As Long                 '密码验证规则（系统参数） 0-密码；1－数字；2－两者皆可
Private mblnOk As Boolean
Private msSource As String                 '数字签名的源字符串
Private mpicSign  As StdPicture
Private morgSign  As StdPicture             '签名原始图(人员表.签名图片)
Public Function ShowMe(ByRef edtThis As Editor, ByRef fParent As Object, ByVal sSource As String, ByRef picSign As StdPicture) As cEPRSign
Dim bytFileKind As Byte    '是否护理病历
Dim lS As Long, rsTemp As ADODB.Recordset, strUserKind As String

    bytFileKind = fParent.Document.EPRPatiRecInfo.病历种类
    Set mpicSign = Nothing
    Set morgSign = Nothing
    Set Sign = New cEPRSign
    Set frmParent = fParent
    msSource = sSource
    
    lblUserName.Caption = gstrUserName
    '根据签名级别来初始化“签名级别＂
    Select Case bytFileKind
    Case cpr诊疗报告
        cmbLevel.AddItem "1 - 医生"
        cmbLevel.AddItem "2 - 主治"
        cmbLevel.AddItem "3 - 主任"
        cmbLevel.ListIndex = 0
        If frmParent.Document.用户签名级别 >= cprSL_主治 Then cmbLevel.ListIndex = 1
        If frmParent.Document.用户签名级别 >= cprSL_主任 Then cmbLevel.ListIndex = 2
    Case Else
        gstrSQL = "Select 人员性质 From 人员性质说明 Where 人员id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "人员性质", glngUserId)
        Do Until rsTemp.EOF
            strUserKind = strUserKind & "," & rsTemp!人员性质
            rsTemp.MoveNext
        Loop
        
        If InStr(strUserKind, "医生") > 0 And InStr(strUserKind, "护士") = 0 Then '操作员只是医生
            cmbLevel.AddItem "1 - 经治医师"
            cmbLevel.AddItem "2 - 主治医师"
            cmbLevel.AddItem "3 - 副主任医师"
            cmbLevel.AddItem "4 - 主任医师"
            cmbLevel.ListIndex = 0
            If frmParent.Document.用户签名级别 >= cprSL_主治 Then cmbLevel.ListIndex = 1
            If frmParent.Document.用户签名级别 >= cprSL_主任 Then cmbLevel.ListIndex = 2
            If frmParent.Document.用户签名级别 >= cprSL_正高 Then cmbLevel.ListIndex = 3
        ElseIf InStr(strUserKind, "医生") = 0 And InStr(strUserKind, "护士") > 0 Then '操作员只是护士
            cmbLevel.AddItem "1 - 护士"
            cmbLevel.AddItem "3 - 护士长"
            cmbLevel.ListIndex = 0
            If frmParent.Document.用户签名级别 >= cprSL_主任 Then cmbLevel.ListIndex = 1
        ElseIf bytFileKind = cpr护理病历 Then ''操作员既是医生又是护士 或都不是时按病历文件类型区别
            cmbLevel.AddItem "1 - 护士"
            cmbLevel.AddItem "3 - 护士长"
            cmbLevel.ListIndex = 0
            If frmParent.Document.用户签名级别 >= cprSL_主任 Then cmbLevel.ListIndex = 1
        Else
            cmbLevel.AddItem "1 - 经治医师"
            cmbLevel.AddItem "2 - 主治医师"
            cmbLevel.AddItem "3 - 副主任医师"
            cmbLevel.AddItem "4 - 主任医师"
            cmbLevel.ListIndex = 0
            If frmParent.Document.用户签名级别 >= cprSL_主治 Then cmbLevel.ListIndex = 1
            If frmParent.Document.用户签名级别 >= cprSL_主任 Then cmbLevel.ListIndex = 2
            If frmParent.Document.用户签名级别 >= cprSL_正高 Then cmbLevel.ListIndex = 3

        End If
    End Select
    
    '读取当前签名方式（系统参数26）
    Select Case bytFileKind
    Case cpr门诊病历
        lS = 1
    Case cpr住院病历
        lS = 2
    Case cpr诊疗报告
        Select Case fParent.Document.EPRFileInfo.lngModule
            Case 1290, 1291, 1294
                lS = 7
            Case Else
                lS = 3
        End Select
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
    
    mlngPassType = Val(Mid(zlDatabase.GetPara(26, glngSys), lS, 1)) '门诊,住院,医技,护理,药品,LIS,PACS (1111111),为空默认采用密码模式
    If mlngPassType = 1 Then
        If gstrESign = "" Or (lS = 3 And gstrESign = "0") Then '医技工作站书写报告没有调用clsDockxx类,如果先刷新"住院病历"页面，再填写报告会在clsDockInEPR中产生gstrESign = "0"
            gstrESign = getPassESign(3, fParent.Document.EPRPatiRecInfo.科室ID)
        End If
        mlngPassType = Val(gstrESign)
    End If
    
    Call optName_Click(0)
    
    Me.Show vbModal, frmParent
    If mblnOk Then
        Set ShowMe = Sign
        If Sign.签名图片 Then
            Set picSign = mpicSign
        Else
            Set picSign = Nothing
        End If
    Else
        Set picSign = Nothing
        Set ShowMe = Nothing
    End If
    Set mpicSign = Nothing
    Set morgSign = Nothing
End Function

'################################################################################################################
'## 功能：  保存签名到内部签名组并刷新显示（验证密码或者数字签名）
'################################################################################################################
Private Function Validation() As Boolean
    Dim blnSpecify As Boolean, strSpecifySign, lngSpecifyId As Long, lngSpecifyLevel As Long, intSign As Integer
    Dim lngCertID As Long, strSign As String, str时间戳 As String, objSignPic As Object, str时间Base64 As String
    Dim rsTemp As ADODB.Recordset, l As Long, strFile As String, strErr As String
    
    On Error GoTo errHand
    intSign = zlDatabase.GetPara("SignShow", glngSys, 1070, 0) '显示姓名还是显示签名 避免同名情况
    
    If optName(1).Value Then  '指定帐号签名
        blnSpecify = True
        txtName = Trim(txtName)
        txtPass = Trim(txtPass)
        
        If frmParent.Document.EPRPatiRecInfo.病历种类 = cpr住院病历 Or frmParent.Document.EPRPatiRecInfo.病历种类 = cpr门诊病历 Or frmParent.Document.EPRPatiRecInfo.病历种类 = cpr护理病历 Then
            gstrSQL = "Select 1 From 上机人员表 A, 部门人员 B Where a.用户名 = [1] And a.人员id = b.人员id And b.部门id = [2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查签名用户与当前用户是否同科室", UCase(txtName.Text), frmParent.Document.EPRPatiRecInfo.科室ID)
            If rsTemp.EOF Then
                MsgBox "指定签名用户与当前操作人员不属于同一科室，禁止操作该科室病人病历！", vbExclamation, gstrSysName: Exit Function
            End If
        End If
        
        If chkEsign.Value = vbUnchecked Then '密码签名
            If Trim(txtPass) = "" Then MsgBox "指定帐号密码不能为空，请检查！", vbExclamation: Exit Function
            If gobjRegister Is Nothing Then Set gobjRegister = DynamicCreate("zlRegister.clsRegister", "密码验证组件")
            If Not gobjRegister.LoginValidate("", txtName, txtPass, strErr) Then
                MsgBox "指定帐号/密码错误,请重新输入登录帐号和密码！" & strErr, vbInformation + vbOKOnly, gstrSysName: Exit Function
            End If
        End If
        
        gstrSQL = "Select b.Id, b.姓名, b.签名" & vbNewLine & _
                    "From 上机人员表 A, 人员表 B" & vbNewLine & _
                    "Where a.用户名 =[1] And a.人员id = b.Id And" & vbNewLine & _
                    "      Nvl(b.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'YYYY-MM-DD')"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "Sign-GetUserInfo", UCase(txtName))
        If rsTemp.EOF Then MsgBox "指定帐号不存在，请重新输入登录帐号和密码!", vbInformation, gstrSysName: Exit Function
        
        If intSign = 0 Then
            strSpecifySign = rsTemp!姓名
        Else
            strSpecifySign = NVL(rsTemp!签名, rsTemp!姓名)         '显示签名
        End If
        lngSpecifyId = rsTemp.Fields("ID")   '用户ID
        
        lngSpecifyLevel = GetUserSignLevel(lngSpecifyId, rsTemp!姓名, frmParent.Document.EPRPatiRecInfo.病人ID, frmParent.Document.EPRPatiRecInfo.主页ID) '获取指定用户的签名级别
        If lngSpecifyLevel = cprSL_空白 Then MsgBox "指定帐号尚未设置签名级别，请在人员管理中调整聘任职务！", vbInformation, gstrSysName: Exit Function
        For l = 1 To frmParent.Document.Signs.Count
            If frmParent.Document.Signs(l).签名级别 > lngSpecifyLevel Then
                MsgBox "当前病历已有更高级别的签名,当前签名级别无权审签本病历", vbInformation, gstrSysName: Exit Function
            End If
        Next
    End If
    
    If Not (IIf(blnSpecify, lngSpecifyLevel, frmParent.Document.用户签名级别) >= Val(cmbLevel.Text)) Then '
        MsgBox "用户拥有的签名级别低于选定的签名级别,请重新选定签名级别！", vbInformation, gstrSysName: Exit Function
    End If

    If chkEsign.Value = vbChecked Then '数字签名,在此窗口中对签名对象进行初始化，此窗口关闭后，数据保存，提取数据生成源内容进行签名，若签名对象初始化失败则不保存
        If gobjESign Is Nothing Then
            Set gobjESign = CreateObject("zl9ESign.clsESign")
            If gobjESign.Initialize(gcnOracle, glngSys) = False Then Exit Function
        End If
        
		If gobjESign.CheckCertificate(IIf(blnSpecify, UCase(txtName), gstrDBUser)) = False Then Exit Function

        If Not gobjESign.CertificateStoped(IIf(blnSpecify, strSpecifySign, gstrUserName)) Then
            strSign = gobjESign.signature(msSource, IIf(chkEsign.Value = vbChecked, IIf(blnSpecify, UCase(txtName), gstrDBUser), ""), lngCertID, str时间戳, objSignPic, str时间Base64) '返回签名信息,lngCertID返回签名使用的证书记录ID
            If strSign = "" Then MsgBox "数字签名失败！请再次签名！", vbInformation + vbOKOnly, gstrSysName: Exit Function
        Else
            chkEsign.Value = vbUnchecked
        End If
    End If
    
    Sign.姓名 = IIf(blnSpecify, strSpecifySign, IIf(intSign = 0, gstrUserName, gstrSignName))
    Sign.签名人ID = IIf(blnSpecify, lngSpecifyId, glngUserId)
    Sign.签名级别 = Val(cmbLevel.Text)
    If Sign.签名级别 > cprSL_主任 Then Sign.签名级别 = cprSL_主任
    
    If zlDatabase.GetPara("将签名级别作为前缀加入", glngSys, 1070, "0") = 1 Then
        Sign.前置文字 = Trim(Mid(Me.cmbLevel.Text, 4)) & "："
    Else
        Sign.前置文字 = ""
    End If
    Sign.显示手签 = (zlDatabase.GetPara("显示手签位置", glngSys, 1070, "0") = 1)
    Sign.签名方式 = IIf(chkEsign.Value = vbUnchecked, 1, 2)
    Sign.签名时间 = zlDatabase.Currentdate()
    Select Case Val(zlDatabase.GetPara("签名时间", glngSys, 1070, "0"))
        Case 1: Sign.显示时间 = "yyyy-MM-dd hh:mm"
        Case 2: Sign.显示时间 = "yyyy年MM月dd日 hh:mm"
        Case Else: Sign.显示时间 = ""
    End Select
    
    '签名规则=2 使用RTF.Text做为数字签名原文 见cEPRSign注释
    Sign.签名规则 = 2
    Sign.签名信息 = strSign
    Sign.证书ID = lngCertID
    Sign.时间戳 = str时间戳
    Sign.时间戳信息 = str时间Base64
'    '签名规则=3 使用保存数据库后的内容文本（不含签名要素，签名对象,图片、表格及子对象）为数字签名原文
'    '数字签名信息在保存后进行数字签名后返回并单独保存
'    Sign.签名规则 = 3
'    Sign.签名信息 = IIf(chkEsign.Value = vbChecked, IIf(blnSpecify, UCase(txtName), gstrDBUser), "") '如果数字签名，先存签名帐号，用于数字签名传入参数,签名完成后更改
'    Sign.证书ID = 0
'    Sign.时间戳 = ""

    pic签名图片.Cls: Set pic签名图片.Picture = Nothing
    pic签名图片.Move pic签名图片.Left, pic签名图片.Top, 810, 810
    
    If zlDatabase.GetPara("签名使用图片", glngSys, 1070, "0", , , , frmParent.Document.EPRPatiRecInfo.科室ID) = 1 Then
        pic签名图片.Visible = True
        strFile = zlBlobRead(15, IIf(optName(0).Value, glngUserId, lngSpecifyId), "", False)
        
        If strFile = "" Then
            MsgBox IIf(optName(0).Value, "当前", "指定") & "帐号没有可用的签名图，不能使用图片签名功能，请联系管理员！", vbExclamation, gstrSysName
            Exit Function
        Else
            Set morgSign = LoadPicture(strFile)
            DrawSignPicture
            Kill strFile
            
            Sign.签名图片 = True
            Set mpicSign = pic签名图片.Picture
        End If
        '使用图片签名后 不显示签名前缀、不显示签名时间、不显示手签
        Sign.前置文字 = ""
        Sign.显示时间 = ""
        Sign.显示手签 = False
    Else
        pic签名图片.Visible = False
        Set morgSign = Nothing
        Sign.签名图片 = False
        Set mpicSign = Nothing
    End If
    
    Validation = True
    Exit Function

errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cmdCancel_Click()
    mblnOk = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Validation Then
        mblnOk = True
        Unload Me
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmParent = Nothing
End Sub
Private Sub optName_Click(Index As Integer)
    Select Case mlngPassType
    Case 0 '密码签名
        chkEsign.Value = vbUnchecked
        chkEsign.Visible = False: chkEsign.Enabled = True
        txtName.Enabled = (Index = 1): txtName.Visible = True
        txtPass.Enabled = (Index = 1): txtPass.Visible = True
        Label2.Visible = True
    Case 1 '1－数字
        chkEsign.Value = vbChecked
        chkEsign.Move txtPass.Left, txtPass.Top
        chkEsign.Visible = True: chkEsign.Enabled = False
        txtName.Enabled = (Index = 1): txtName.Visible = True
        txtPass.Enabled = False: txtPass.Visible = False
        Label2.Visible = False
    End Select

    If txtName.Enabled And txtName.Visible Then
        txtName.SetFocus
    End If
End Sub
Private Sub txtName_GotFocus()
    zlControl.TxtSelAll txtName
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If chkEsign.Value = vbUnchecked Then
            If txtPass.Enabled And txtPass.Visible Then zlControl.TxtSelAll txtPass: txtPass.SetFocus:  Exit Sub
        Else
            Call zlCommFun.PressKey(vbKeyTab): Exit Sub
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
    If KeyAscii = vbKeyBack Then Exit Sub
    If KeyAscii < 32 Or KeyAscii > 126 Then KeyAscii = 0
    If InStr("""@\ ", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub
Private Sub DrawSignPicture()
Dim lngHeight As Long, lngWidth As Long
    On Error Resume Next
    If Not morgSign Is Nothing Then
        pic签名图片.Appearance = 0: pic签名图片.BorderStyle = 0
        If zlDatabase.GetPara("签名使用原图", glngSys, 1070, "1") = 1 Then
            Set pic签名图片.Picture = morgSign
            If pic签名图片.Width <> 810 Then pic签名图片.Move pic签名图片.Left, pic签名图片.Top, 810, 810
            pic签名图片.PaintPicture pic签名图片.Picture, 0, 0, pic签名图片.ScaleX(pic签名图片.Width, vbTwips, vbPixels), pic签名图片.ScaleY(pic签名图片.Height, vbTwips, vbPixels)
        Else
            lngHeight = zlDatabase.GetPara("签名图片高度", glngSys, 1070, "50")
            lngWidth = CLng(lngHeight * (morgSign.Width / morgSign.Height))
            pic签名图片.Move pic签名图片.Left, pic签名图片.Top, pic签名图片.ScaleX(lngWidth, vbPixels, vbTwips), pic签名图片.ScaleY(lngHeight, vbPixels, vbTwips)
            pic签名图片.PaintPicture morgSign, 0, 0, lngWidth, lngHeight
            Set pic签名图片.Picture = pic签名图片.Image
        End If
    End If
    Err.Clear
End Sub
