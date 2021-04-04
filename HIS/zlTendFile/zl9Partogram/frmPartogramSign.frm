VERSION 5.00
Begin VB.Form frmPartogramSign 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "签名"
   ClientHeight    =   3345
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5295
   Icon            =   "frmPartogramSign.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox PicInfo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   60
      Picture         =   "frmPartogramSign.frx":000C
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   13
      Top             =   2500
      Width           =   240
   End
   Begin VB.ComboBox cboMan 
      Height          =   300
      Left            =   1365
      TabIndex        =   2
      Text            =   "cboMan"
      Top             =   1800
      Width           =   2505
   End
   Begin VB.ComboBox cmbLevel 
      Height          =   300
      Left            =   1365
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2160
      Width           =   2505
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2640
      TabIndex        =   6
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3960
      TabIndex        =   7
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   -375
      TabIndex        =   8
      Top             =   2760
      Width           =   5670
   End
   Begin VB.CheckBox chkEsign 
      BackColor       =   &H00FFFFFF&
      Caption         =   "数字签名(&E)"
      Height          =   195
      Left            =   3930
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2220
      Width           =   1305
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "提示："
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   360
      TabIndex        =   12
      Top             =   2530
      Width           =   540
   End
   Begin VB.Label lbl签名人 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "签名人(&P)"
      Height          =   180
      Left            =   435
      TabIndex        =   1
      Top             =   1860
      Width           =   810
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "审签："
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   1
      Left            =   750
      TabIndex        =   11
      Top             =   990
      Width           =   540
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "平签："
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   0
      Left            =   750
      TabIndex        =   10
      Top             =   180
      Width           =   540
   End
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "根据所选数据的最后审签人的最高级别，程序自动选择您相应的更高级别。"
      ForeColor       =   &H00FF0000&
      Height          =   360
      Index           =   1
      Left            =   1080
      TabIndex        =   9
      Top             =   1200
      Width           =   3960
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   210
      Picture         =   "frmPartogramSign.frx":034E
      Stretch         =   -1  'True
      Top             =   240
      Width           =   360
   End
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "对自己修改过的数据进行签名，程序缺省选择最高级别；对他人已签名的数据修改后签名，程序自动选择相同级别。"
      ForeColor       =   &H00FF0000&
      Height          =   540
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   420
      Width           =   3960
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblLevel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "签名级别(&L)"
      Height          =   180
      Left            =   255
      TabIndex        =   3
      Top             =   2220
      Width           =   990
   End
End
Attribute VB_Name = "frmPartogramSign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private frmParent As Object                 '父窗体
Private mblnOK As Boolean
Private Sign As cPartogramSign                    '签名对象

Private lngCertID As Long                   '证书ID
Private mlngPassType As Long                 '密码验证规则（系统参数） 0-密码；1－数字；2－两者皆可
Private mbln审签 As Boolean                 '是否审签
Private mlngCur As Long, mlngLast As Long   '当前人员级别，审签人级别

Private mlng文件ID As Long
Private mlngDeptID As Long
Private mstrSource As String                 '数字签名的源字符串
Private mstr状态 As String
Private mstrUserInfo As String
Private mstrPrivs As String
Private mblnDrop As Boolean
'--人员信息
Private mlngUserID As Long
Private mstrUserName As String               '当前用户姓名
Private mstrUserAbbr As String               '当前用户简码
Private mrs签名人 As New ADODB.Recordset
Private gstrLike As String

Private Enum SignLevel
    正高 = 1
    副高 = 2
    中级 = 3
    师级 = 4
    员士 = 5
    未定义 = 9
End Enum

'######################################################################################################################
'病人护理数据.审核人：保存最后一次签名人与第一次签名人，格式为：审签/签名
'记录类型 = 1 And 终止版本 Is NULL为原始记录
'病人护理数据.审核人为NULL，未签名；不含/表示已签名；含/表示已审签
'未审签之前，同级可以相互修改，签名，一旦审签后，就只能继续审签。
'取消审签时，自动删除修改痕迹
'审签后，不再增加记录类型=5的签名记录
'存在审签记录时，不允许修改，要么不断的审签，要么一直回退到普通签名记录状态
'产生新的审签记录或审签回退时，审核人字段要更新
'######################################################################################################################

'电子签名使用场合：
'26  电子签名使用场合(4位字符) 对不同场合是否使用电子签名进行控制,数字位数分别为:门诊,住院,医技,护理 0-不控制,1-控制
Public Function ShowMe(ByVal objParent As Object, ByVal strPrivs As String, ByVal lng文件ID As Long, ByVal lngDeptID As Long, ByVal intLevel As Integer, _
    ByVal sSource As String, ByVal bln审签 As Boolean, Optional str状态 As String, Optional strUserInfo As String) As cPartogramSign
    '******************************************************************************************************************
    '功能： 显示签名窗体
    '参数： edtThis     :IN     编辑器控件
    '       fParent     :IN     父窗体
    '       mstrSource   :IN     数字签名的源字符串（从文本中提取，去掉签名提纲）
    '       str状态     :IN     用于连续签名时传入，避免频繁弹出签名窗体
    '       str审签人   :IN     审签时传入上次审签人姓名，以便核实审签权限
    '       strUseringo :IN   人员ID'人员简码'人员姓名
    '******************************************************************************************************************
    
    Dim arrUser
    
    Set Sign = New cPartogramSign
    Set frmParent = objParent
    mstrSource = sSource
    mstr状态 = str状态
    mbln审签 = bln审签
    mlngLast = intLevel
    mlng文件ID = lng文件ID
    mlngDeptID = lngDeptID
    mstrPrivs = strPrivs
    mblnOK = False
    '第一次调用
    If mstr状态 = "" Then
        mlngUserID = glngUserId
        mstrUserName = Replace(gstrUserName, "-", "")
        mstrUserAbbr = Replace(gstrUserAbbr, "-", "")
    Else
        arrUser = Split(strUserInfo, "'")
        mlngUserID = Val(arrUser(0))
        mstrUserName = CStr(arrUser(1))
        mstrUserAbbr = CStr(arrUser(2))
    End If
    
    Call GetUser(lngDeptID)
    
    gstrLike = IIf(zlDatabase.GetPara("输入匹配") = "0", "%", "")

    Call RefControls
    
    If mstr状态 <> "" Then
        '连续签名时
        Call cmdOK_Click
    Else
        Me.Show vbModal, frmParent
    End If
    
    If mblnOK Then
        str状态 = mstr状态
        strUserInfo = mstrUserInfo
        Set ShowMe = Sign
    Else
        Set ShowMe = Nothing
    End If
End Function

Public Sub GetUser(ByVal lngWorkID As Long)
    Dim rs As New ADODB.Recordset
    
    Err = 0: On Error GoTo errHand
    
    '提取本科室所有人员信息
    gstrSQL = " Select Distinct a.Id, b.部门id, a.编号, a.姓名, Upper(a.简码) As 简码, c.人员性质, Nvl(a.聘任技术职务, 0) As 职务, b.缺省" & vbNewLine & _
            " From 人员表 A, 部门人员 B, 人员性质说明 C, 部门性质说明 D" & vbNewLine & _
            " Where a.Id = b.人员id And a.Id = c.人员id And b.部门id = d.部门id And c.人员性质 In ('医生', '护士') And" & vbNewLine & _
            "      (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) And b.部门id = [1]" & vbNewLine & _
            " Order By 简码, 缺省 Desc"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mRichEPR", lngWorkID)
    Set mrs签名人 = rs
    With cboMan
        .Clear
        If Not rs.EOF Then
            Do While Not rs.EOF
            .AddItem Replace(NVL(rs!简码), "-", "") & "-" & Replace(NVL(rs!姓名), "-", "")
            .ItemData(.NewIndex) = Val(rs!ID)
            rs.MoveNext
            Loop
        Else
            .AddItem mstrUserAbbr & "-" & mstrUserName
            .ItemData(.NewIndex) = mlngUserID
        End If
    End With
    
    '定位到当前操作员
    Call isCheck签名人Exists(mstrUserName, True)
    If cboMan.ListIndex = -1 Then cboMan.ListIndex = 0
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub GetUserLevel(ByVal lngUserID As Long)
    Dim str签名人 As String, str审签人 As String
    Dim rs As New ADODB.Recordset
    
    Err = 0: On Error GoTo errHand
    mlngCur = 未定义
    '级别是反着的，1正高最大，所以，判断值必须小于审签人的级别，否则不允许签名
    
    '取当前操作员的级别
    gstrSQL = "select  聘任技术职务 from 人员表 p where ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mRichEPR", lngUserID)
    If Not rs.EOF Then
        mlngCur = NVL(rs("聘任技术职务"), 未定义)
    End If
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Function Validation() As Boolean
    '******************************************************************************************************************
    '
    '功能：  保存签名到内部签名组并刷新显示（验证密码或者数字签名）
    '
    '******************************************************************************************************************
    On Error GoTo errHand
    Dim intLevel As Integer '0-正高,原级别减1,为了兼容签名级别的定义
    Dim strUserName As String, lngUserID As Long, strSign As String, str时间戳 As String, str时间戳信息 As String
    
    '检查签名人是否存在
    If Not isCheck签名人Exists(Mid(cboMan.Text, InStr(cboMan.Text, "-") + 1)) Then
        lblInfo.Caption = "提示：签名人信息不不存在,请检查！"
        If cboMan.Enabled = True And cboMan.Visible = True Then cboMan.SetFocus
        Exit Function
    End If
    
    mstrUserName = Mid(cboMan.Text, InStr(cboMan.Text, "-") + 1)
    mstrUserAbbr = Mid(cboMan.Text, 1, InStr(cboMan.Text, "-") - 1)
    mlngUserID = Val(cboMan.ItemData(cboMan.ListIndex))
    If chkEsign.Value = vbChecked Then
        '数字签名
        Err.Clear
        If gobjESign Is Nothing Then
            On Error Resume Next
            Set gobjESign = CreateObject("zl9ESign.clsESign")
            If Err <> 0 Then Err.Clear: strSign = ""
            On Error GoTo 0
            If Not gobjESign Is Nothing Then
                Call gobjESign.Initialize(gcnOracle, glngSys)
            End If
        End If
        If gobjESign Is Nothing Then
            MsgBox "电子签名部件未能正确安装，签名操作不能继续！", vbInformation, gstrSysName
            Exit Function
        End If
        lngCertID = 0
        strSign = gobjESign.signature(mstrSource, UCase(gcnOracle.Properties(23)), lngCertID, str时间戳, , str时间戳信息) '返回签名信息,lngCertID返回签名使用的证书记录ID
        If strSign = "" Then
            MsgBox "签名失败！", vbInformation + vbOKOnly, "签名"
            Exit Function
        End If
    End If
    strUserName = mstrUserName
    lngUserID = mlngUserID
    
    '下次读取会+1
    Select Case Mid(cmbLevel.Text, 1, 1)
    Case 5
        intLevel = 0    '1
    Case 4
        intLevel = 1    '2
    Case 3
        intLevel = 2    '3
    Case 2
        intLevel = 3    '4
    Case 1
        intLevel = 4    '5
    End Select
    
    '------------------------------------------------------------------------------------------------------------------
    Sign.姓名 = mstrUserName
    Sign.签名级别 = intLevel                    '-1是为了兼容签名级别定义
    Sign.签名信息 = strSign
    Sign.签名方式 = IIf(chkEsign.Value = vbUnchecked, 1, 2)
    Sign.签名规则 = 1
    Sign.证书ID = IIf(Sign.签名方式 = 2, lngCertID, 0)
    Sign.时间戳 = str时间戳
    Sign.时间戳信息 = str时间戳信息
    
    Validation = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
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
'## 功能：  刷新控件
'################################################################################################################
Private Sub RefControls()
    Dim rsTemp As New ADODB.Recordset
    Dim arrData
    On Error GoTo errHand
    
    '63955:刘鹏飞,2013-09-16,启用点击签名，并且当前签名的病区在设置的电子签名启用部门中才能使用电子签名
    '说明：如果没有设置电子签名所要启用的部门,就说明启用电子签名的病区为所有病区
    If mstr状态 <> "" And InStr(1, mstr状态, "|") <> 0 Then
        arrData = Split(mstr状态, "|")
        mlngPassType = Val(arrData(1))
        cmbLevel.ListIndex = Val(arrData(0))
    Else
        gstrSQL = "Select Zl_Fun_Getsignpar([1],[2]) 电子签名 From dual"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "电子签名启用部门", 4, mlngDeptID)
        If rsTemp.RecordCount > 0 Then
            mlngPassType = Val(NVL(rsTemp!电子签名, 0))
        Else
            mlngPassType = 0
        End If
        If mlngPassType = 1 Then
            If CertificateStoped(gstrUserName) = True Then mlngPassType = 0
        End If
    End If
    
    Select Case mlngPassType
    Case 1
        '1－启用电子签名
        chkEsign.Value = vbChecked
        chkEsign.Visible = True
        chkEsign.Enabled = False
    Case Else
        '不启用电子签名
        chkEsign.Value = vbUnchecked
        chkEsign.Visible = False
    End Select
    
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
    If gobjESign Is Nothing Then
        Set gobjESign = CreateObject("zl9ESign.clsESign")
        If Err <> 0 Then Err.Clear
        If Not gobjESign Is Nothing Then Call gobjESign.Initialize(gcnOracle, glngSys)
    End If
    If gobjESign Is Nothing Then Exit Function
    CertificateStoped = gobjESign.CertificateStoped(strName)
    If Err <> 0 Then Err.Clear
End Function

Private Sub cboMan_Click()
    cmbLevel.Clear
    lblInfo.Caption = "提示："
     '根据用户的签名级别来初始化“签名级别＂
    Call GetUserLevel(Val(cboMan.ItemData(cboMan.ListIndex)))            '获取用户签名级别
    
    '审签则加入比上次级别高的;平签则只能以上次相同级别进行
    If mbln审签 Or mlngLast = 未定义 Then
        If Not (mlngCur < mlngLast) Then
            If mbln审签 = True Then
                lblInfo.Caption = "提示：当前签名人要超过本记录的签名者或上次审签者的级别才能审签！"
            Else
                '说明记录还没有进行过签名
                lblInfo.Caption = "提示：该签名人还未设置聘任职务级别，请在人员管理中进行设置！"
            End If
            cmdOK.Enabled = False
            Exit Sub
        End If
        If mlngCur <= 正高 And 正高 < mlngLast Then cmbLevel.AddItem "5-主任护师"
        If mlngCur <= 副高 And 副高 < mlngLast Then cmbLevel.AddItem "4-副主任护师"
        If mlngCur <= 中级 And 中级 < mlngLast Then cmbLevel.AddItem "3-主管护师"
        If mlngCur <= 师级 And 师级 < mlngLast Then cmbLevel.AddItem "2-护师"
        If mlngCur <= 员士 And 员士 < mlngLast Then cmbLevel.AddItem "1-护士"
        If mlngCur > 员士 Then cmbLevel.AddItem "0-未定义"
    Else
        If Not (mlngCur <= mlngLast) Then
            lblInfo.Caption = "提示：当前签名人至少要达到上次签名者的级别才能签名！"
            cmdOK.Enabled = False
            Exit Sub
        End If
        Select Case mlngCur
        Case 正高
            cmbLevel.AddItem "5-主任护师"
        Case 副高
            cmbLevel.AddItem "4-副主任护师"
        Case 中级
            cmbLevel.AddItem "3-主管护师"
        Case 师级
            cmbLevel.AddItem "2-护师"
        Case 员士
            cmbLevel.AddItem "1-护士"
        End Select
    End If
    cmbLevel.ListIndex = 0
End Sub

Private Sub cboMan_KeyDown(KeyCode As Integer, Shift As Integer)
    If cboMan.Locked Then Exit Sub
    mblnDrop = False
    If KeyCode = 13 Then mblnDrop = SendMessage(cboMan.Hwnd, &H157, 0, 0) = 1
End Sub

Private Sub cboMan_KeyPress(KeyAscii As Integer)
    'Call zlControl.CboMatchIndex(cboMan.Hwnd, KeyAscii)
    Dim i As Long, intIdx As Integer, iCount As Integer
    Dim strText As String, strResult As String, strFilter As String
    Dim rsTemp As ADODB.Recordset
    If KeyAscii = 13 Then
        If cboMan.Locked Then
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        strText = UCase(cboMan.Text)
        If cboMan.ListIndex <> -1 Then
            '弹出列表时,又在文本框输入了内容
            If strText <> cboMan.List(cboMan.ListIndex) Then Call zlControl.CboSetIndex(cboMan.Hwnd, -1)
        End If
        If strText = "" Then
            cboMan.ListIndex = -1
        ElseIf cboMan.ListIndex = -1 Then
            intIdx = -1
            strFilter = ""
            '先复制记录集
            Set rsTemp = zlDatabase.zlCopyDataStructure(mrs签名人)
            Dim intInputType As Integer '0-输入的是全数字,1-输入的是全字母,2-其他
            Dim strCompents As String '匹配串
            
            strCompents = Replace(gstrLike, "%", "*") & strText & "*"
            
            If IsNumeric(strText) Then
                intInputType = 0
            ElseIf zlCommFun.IsCharAlpha(strText) Then
                intInputType = 1
            Else
                intInputType = 2
            End If
            
            mrs签名人.Filter = strFilter: iCount = 0
            With mrs签名人
                If .RecordCount <> 0 Then .MoveFirst
                Do While Not mrs签名人.EOF
                    Select Case intInputType
                    Case 0  '输入的是全数字
                        '如果输入的数字,需要检查:
                        '1.编号输入值相等,主要输入如:12 匹配000012这种情况
                        '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                        
                        
                        '主要是检查输入的内容与编号完全相同,则直接就定位到该姓名
                        If NVL(!编号) = strText Then strResult = NVL(!姓名): iCount = 0: Exit Do
                        
                        '1.编号输入值相等,主要输入如:12 匹配000012这种情况,因为这种情况有很多:如0012,012,000012等.因此如果存在此种情况,需要弹出选择器供选择
                        If Val(NVL(!编号)) = Val(strText) Then
                            If iCount = 0 Then strResult = NVL(!姓名)
                            iCount = iCount + 1
                        End If
                        
                        '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                         If Val(NVL(!编号)) Like strText & "*" Then
                            If isCheck签名人Exists(NVL(!姓名)) Then Call zlDatabase.zlInsertCurrRowData(mrs签名人, rsTemp)
                         End If
                    Case 1  '输入的是全字母
                        '规则:
                        ' 1.输入的简码相等,则直接定位
                        ' 2.根据参数来匹配相同数据
                        
                        '1.输入的简码相等,则直接定位
                        If Trim(NVL(!简码)) = strText Then
                            If iCount = 0 Then strResult = NVL(!姓名)   '可能存在多个相同的多个
                            iCount = iCount + 1
                        End If
                        
                        '2.根据参数来匹配相同数据
                        If Trim(NVL(!简码)) Like strCompents Then
                            If isCheck签名人Exists(NVL(!姓名)) Then Call zlDatabase.zlInsertCurrRowData(mrs签名人, rsTemp)
                        End If
                    Case Else  ' 2-其他
                        '规则:可能存在汉字等情况,或编号类似于N001简码可能有ZYK01这种情况
                        '1.编码\简码相等,直接定位
                        '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                        
                        '1.编码\简码相等,直接定位
                        If Trim(!编号) = strText Or Trim(!简码) = strText Or Trim(!姓名) = strText Then
                            If iCount = 0 Then strResult = NVL(!姓名)   '可能存在多个相同的多个
                            iCount = iCount + 1
                        End If
                        
                        '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                        If Trim(!编号) Like strText & "*" Or Trim(NVL(!简码)) Like strCompents Or Trim(NVL(!姓名)) Like strCompents Then
                            If isCheck签名人Exists(NVL(!姓名)) Then Call zlDatabase.zlInsertCurrRowData(mrs签名人, rsTemp)
                        End If
                    End Select
                    mrs签名人.MoveNext
                Loop
            End With
            If iCount > 1 Then strResult = ""
            If strResult = "" And rsTemp.RecordCount = 1 Then strResult = NVL(rsTemp!姓名)
            '直接定位
            If strResult <> "" Then
                rsTemp.Close: Set rsTemp = Nothing
                If isCheck签名人Exists(strResult, True) Then zlCommFun.PressKey vbKeyTab
                Exit Sub
            End If
            
            '需要检查是否有多条满足条件的记录
            If rsTemp.RecordCount <> 0 Then
                '先按某种方式进行排序
                Select Case intInputType
                Case 0 '输入全数字
                    rsTemp.Sort = "编号"
                Case 1 '输入全拼音
                    rsTemp.Sort = "简码"
                Case Else
                    '根据选择来定
                    rsTemp.Sort = "简码"
                End Select
                '弹出选择器
                Dim rsReturn As ADODB.Recordset
                If zlDatabase.zlShowListSelect(Me, glngSys, 1133, cboMan, rsTemp, True, "", "缺省,职务,优先级别", rsReturn) Then
                    If Not rsReturn Is Nothing Then
                        If rsReturn.RecordCount <> 0 Then
                            '进行定位
                            If isCheck签名人Exists(NVL(rsReturn!姓名), True) Then
                                'zlCommFun.PressKey vbKeyTab
                            End If
                        End If
                    End If
                End If
            Else
                '未找到
                rsTemp.Close: Set rsTemp = Nothing
                KeyAscii = 0: zlControl.TxtSelAll cboMan: Exit Sub
            End If
            rsTemp.Close: Set rsTemp = Nothing
             
        ElseIf Not mblnDrop Then
            '回车光标经过
            Call cboMan_Click
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        If cboMan.ListIndex = -1 Then
            cboMan.Text = ""
            Exit Sub
        Else
            If intIdx <> -1 And mblnDrop Then
                '弹出回车-强行激活Click
                Call cboMan_Click
            ElseIf intIdx <> cboMan.ListIndex And intIdx <> -1 Then
                '弹出让选择-自动激活Click
                cboMan.SetFocus
                Call zlCommFun.PressKey(vbKeyF4)
                Exit Sub
            ElseIf intIdx <> -1 Then
                '一次性输中-强行激活Click
                Call cboMan_Click
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Function isCheck签名人Exists(ByVal str姓名 As String, Optional blnLocateItem As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查姓名是否在开单人下拉列表中.
    '入参:str姓名-姓名
    '     blnLocateItem:是否直接定位
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-07-20 17:53:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    For i = 0 To cboMan.ListCount - 1
        If NeedName(cboMan.List(i)) = str姓名 Then
            If blnLocateItem Then cboMan.ListIndex = i
            isCheck签名人Exists = True
            Exit Function
        End If
    Next
End Function

Private Function NeedName(strList As String) As String
    NeedName = Mid(strList, InStr(strList, "-") + 1)
End Function

Private Sub cboMan_Validate(Cancel As Boolean)
    If cboMan.Text <> "" Then
        If GetCboIndex(cboMan, NeedName(cboMan.Text)) = -1 Then cboMan.ListIndex = -1: cboMan.Text = ""
    End If
    If cboMan.Text = "" Then '说明录入的信息，不存在列表中
        lblInfo.Caption = "提示：签名人信息不不存在,请检查！"
        cmdOK.Enabled = False
        Cancel = True
    End If
End Sub

Private Sub cmbLevel_Click()
    cmdOK.Enabled = (Mid(Me.cmbLevel.Text, 1, 1) > 0)
End Sub

Private Sub cmbLevel_KeyPress(KeyAscii As Integer)
    Call zlControl.CboMatchIndex(cmbLevel.Hwnd, KeyAscii)
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Validation Then
        mstr状态 = cmbLevel.ListIndex & "|" & chkEsign.Value
        mstrUserInfo = mlngUserID & "'" & mstrUserName & "'" & mstrUserAbbr
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

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Function GetCboIndex(cbo As ComboBox, strFind As String, Optional blnKeep As Boolean, Optional blnLike As Boolean) As Long
'功能：由字符串在ComboBox中查找索引
    Dim i As Long
    If strFind = "" Then GetCboIndex = -1: Exit Function
    '先精确查找
    For i = 0 To cbo.ListCount - 1
        If InStr(cbo.List(i), "-") > 0 Then
            If NeedName(cbo.List(i)) = strFind Then GetCboIndex = i: Exit Function
        Else
            If cbo.List(i) = strFind Then GetCboIndex = i: Exit Function
        End If
    Next
    '最后模糊查找
    If blnLike Then
        For i = 0 To cbo.ListCount - 1
            If InStr(cbo.List(i), strFind) > 0 Then GetCboIndex = i: Exit Function
        Next
    End If
    If Not blnKeep Then GetCboIndex = -1
End Function

Private Sub PicInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call zlCommFun.ShowTipInfo(PicInfo.Hwnd, lblInfo.Caption)
End Sub
