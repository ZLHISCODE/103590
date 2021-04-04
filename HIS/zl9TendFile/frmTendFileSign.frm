VERSION 5.00
Begin VB.Form frmTendFileSign 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "签名"
   ClientHeight    =   2835
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5295
   Icon            =   "frmTendFileSign.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cmbLevel 
      Height          =   300
      ItemData        =   "frmTendFileSign.frx":000C
      Left            =   1365
      List            =   "frmTendFileSign.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1800
      Width           =   2505
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2670
      TabIndex        =   4
      Top             =   2370
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3930
      TabIndex        =   5
      Top             =   2370
      Width           =   1095
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   -375
      TabIndex        =   6
      Top             =   2250
      Width           =   5670
   End
   Begin VB.CheckBox chkEsign 
      BackColor       =   &H00FFFFFF&
      Caption         =   "数字签名(&E)"
      Height          =   195
      Left            =   3930
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1860
      Width           =   1305
   End
   Begin VB.Label lblsinName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "签名人：刘鹏飞"
      Height          =   180
      Left            =   255
      TabIndex        =   10
      Top             =   2455
      Width           =   1260
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
      TabIndex        =   9
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
      TabIndex        =   8
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
      TabIndex        =   7
      Top             =   1200
      Width           =   3960
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   210
      Picture         =   "frmTendFileSign.frx":0010
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
      TabIndex        =   1
      Top             =   1860
      Width           =   990
   End
End
Attribute VB_Name = "frmTendFileSign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private frmParent As Object                 '父窗体
Private mblnOK As Boolean
Private Sign As cTendSign                    '签名对象

Private lngCertID As Long                   '证书ID
Private mlngPassType As Long                '控制是否启用电子签名（系统参数） 0-不控制；1－控制
Private mbln审签 As Boolean                 '是否审签
Private mlngCur As Long, mlngLast As Long   '当前人员级别，审签人级别

Private mlng文件ID As Long
Private mstrSource As String                 '数字签名的源字符串
Private mstr状态 As String
Private mstrPrivs As String
Private mlngUnitID As Long                   '当前病区ID

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
Public Function ShowMe(ByVal objParent As Object, ByVal strPrivs As String, ByVal lng文件ID As Long, ByVal lngUnitId As Long, ByVal intLevel As Integer, _
    ByVal sSource As String, ByVal bln审签 As Boolean, Optional str状态 As String, Optional str错误 As String, _
    Optional ByVal intSignMode As Integer = 0, Optional ByVal blnExchange As Boolean = False) As cTendSign
    '******************************************************************************************************************
    '功能： 显示签名窗体
    '参数： edtThis     :IN     编辑器控件
    '       objParent     :IN     父窗体
    '       lng文件ID   :IN      文件ID
    '       lngUnitId   :IN      病区ID
    '       mstrSource   :IN     数字签名的源字符串（从文本中提取，去掉签名提纲）
    '       str状态     :IN     用于连续签名时传入，避免频繁弹出签名窗体
    '       str审签人   :IN     审签时传入上次审签人姓名，以便核实审签权限
    '******************************************************************************************************************
    Dim strLastInfo As String
    
    Set Sign = New cTendSign
    Set frmParent = objParent
    mstrSource = sSource
    mstr状态 = str状态
    mbln审签 = bln审签
    mlngLast = intLevel
    mlng文件ID = lng文件ID
    mstrPrivs = strPrivs
    mlngUnitID = lngUnitId
    '76700:LPF:签名成功在审签时，点击签名窗体关闭按钮，就会导致签名人为空的记录。
    mblnOK = False
    
    '根据用户的签名级别来初始化“签名级别＂
    Call GetUserLevel(glngUserId)           '获取用户签名级别
    strLastInfo = ""
    If Not mlngLast = 未定义 Then
        Select Case mlngLast
        Case 正高
            strLastInfo = "5-主任护师"
        Case 副高
            strLastInfo = "4-副主任护师"
        Case 中级
            strLastInfo = "3-主管护师"
        Case 师级
            strLastInfo = "2-护师"
        Case 员士
            strLastInfo = "1-护士"
        End Select
    End If
    '审签则加入比上次级别高的;平签则只能以上次相同级别进行
    If bln审签 Or mlngLast = 未定义 Then
        '43588:刘鹏飞,2012-09-13,添加记录单审签模式
        If mlngCur = 未定义 Then
            If bln审签 = True Then
                str错误 = "您要超过本记录的签名者或上次审签者的级别才能审签，" & vbCrLf & _
                    "而您当前还未设置聘任技术职务，请在人员管理中设置！"
            Else
                str错误 = "您当前还未设置聘任技术职务，请在人员管理中设置！"
            End If
            Unload Me
            Exit Function
        End If

        If IIf(bln审签 = True And intSignMode = 1, False, (Not (mlngCur < mlngLast))) Then
            str错误 = "您要超过本记录的签名者或上次审签者的级别才能审签！" & IIf(strLastInfo = "", "", "上次签名人级别【" & strLastInfo & "】")
            Unload Me
            Exit Function
        End If

        If mlngCur <= 正高 And 正高 < mlngLast Then cmbLevel.AddItem "5-主任护师"
        If mlngCur <= 副高 And 副高 < mlngLast Then cmbLevel.AddItem "4-副主任护师"
        If mlngCur <= 中级 And 中级 < mlngLast Then cmbLevel.AddItem "3-主管护师"
        If mlngCur <= 师级 And 师级 < mlngLast Then cmbLevel.AddItem "2-护师"
        If mlngCur <= 员士 And 员士 < mlngLast Then cmbLevel.AddItem "1-护士"
        If mlngCur > 员士 Then cmbLevel.AddItem "0-未定义"
    Else
        If Not (mlngCur <= mlngLast) Then
            str错误 = "您至少要达到上次签名者的级别才能签名！" & IIf(strLastInfo = "", "", "上次签名人级别【" & strLastInfo & "】")
            Unload Me
            Exit Function
        End If
        '51589:刘鹏飞,2013-03-01,添加交班签名
        If bln审签 = False And blnExchange = True Then
            Select Case mlngLast
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
        Else
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
    End If
    If bln审签 = True And intSignMode = 1 Then
        cmbLevel.ListIndex = cmbLevel.ListCount - 1
    Else
        cmbLevel.ListIndex = 0
    End If
    
    lblsinName.Caption = "签名人：" & gstrUserName
    
    If RefControls = False Then
        Unload Me
        Exit Function
    End If
    
    '43588:刘鹏飞,2012-09-13,添加记录单审签模式
    '51589:刘鹏飞,2013-03-01,添加交班签名
    If mstr状态 <> "" Or (bln审签 = True And intSignMode = 1) Or (bln审签 = False And blnExchange = True) Then
        '连续签名时
        Call cmdOK_Click
    Else
        Me.Show vbModal, frmParent
    End If
    
    If mblnOK Then
        str状态 = mstr状态
        Set ShowMe = Sign
    Else
        Set ShowMe = Nothing
    End If
End Function

Public Sub GetUserLevel(ByVal lngUserID As Long)
    Dim str签名人 As String, str审签人 As String
    Dim rs As New ADODB.Recordset
    
    Err = 0: On Error GoTo ErrHand
    mlngCur = 未定义
    '级别是反着的，1正高最大，所以，判断值必须小于审签人的级别，否则不允许签名
    
    '取当前操作员的级别
    gstrSQL = "select  聘任技术职务 from 人员表 p where ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mRichEPR", lngUserID)
    If Not rs.EOF Then
        mlngCur = NVL(rs("聘任技术职务"), 未定义)
    End If
    Exit Sub
ErrHand:
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
    On Error GoTo ErrHand
    Dim intLevel As Integer '0-正高,原级别减1,为了兼容签名级别的定义
    Dim strUserName As String, lngUserID As Long, strSign As String, str时间戳 As String, str时间戳信息 As String
    
    If chkEsign.Value = vbChecked Then
        '数字签名
        If InitESign = False Then
            MsgBox "电子签名部件未能正确安装，签名操作不能继续！", vbInformation, gstrSysName
            Exit Function
        End If
        lngCertID = 0
        strSign = gobjESign.signature(mstrSource, gstrDBUser, lngCertID, str时间戳, , str时间戳信息)  '返回签名信息,lngCertID返回签名使用的证书记录ID
        If strSign = "" Then
            MsgBox "签名失败！", vbInformation + vbOKOnly, "签名"
            Exit Function
        End If
    End If
    strUserName = gstrUserName
    lngUserID = glngUserId
    
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
    Sign.姓名 = strUserName
    Sign.签名级别 = intLevel                    '-1是为了兼容签名级别定义
    Sign.签名信息 = strSign
    Sign.签名方式 = IIf(chkEsign.Value = vbUnchecked, 1, 2)
    Sign.签名规则 = 1
    Sign.证书ID = IIf(Sign.签名方式 = 2, lngCertID, 0)
    Sign.时间戳 = str时间戳
    Sign.时间戳信息 = str时间戳信息
    
    Validation = True
    Exit Function
ErrHand:
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
ErrHand:
    Set Cn = Nothing
    OraDataOpen = False
    Err = 0
End Function

'################################################################################################################
'## 功能：  刷新控件
'################################################################################################################
Private Function RefControls() As Boolean
    Dim arrData
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    
    '63955:刘鹏飞,2013-09-16,启用点击签名，并且当前签名的病区在设置的电子签名启用部门中才能使用电子签名
    '说明：如果没有设置电子签名所要启用的部门,就说明启用电子签名的病区为所有病区
    If mstr状态 <> "" And InStr(1, mstr状态, "|") <> 0 Then
        arrData = Split(mstr状态, "|")
        mlngPassType = Val(arrData(1))
        cmbLevel.ListIndex = Val(arrData(0))
    Else
        gstrSQL = "Select Zl_Fun_Getsignpar([1],[2]) 电子签名 From dual"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "电子签名启用部门", 4, mlngUnitID)
        If rsTemp.RecordCount > 0 Then
            mlngPassType = Val(NVL(rsTemp!电子签名, 0))
        Else
            mlngPassType = 0
        End If
        '123565,数据签名调整
        If mlngPassType = 1 Then
            If InitESign = True Then
                If gobjESign.CheckCertificate(gstrDBUser) = True Then ''证书已经注册，且证书没有停用，且插入了key，则允许进行数据签名，否则签名终止
                    If gobjESign.CertificateStoped(gstrUserName) = True Then mlngPassType = 0 '检查签名人的证书是否停用，停用的话将不使用电子签名，使用密码签名
                Else
                    '终止签名操作
                    Exit Function
                End If
            Else
                mlngPassType = 0  '签名部件创建失败，数据密码签名
            End If
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
    
    RefControls = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function InitESign() As Boolean
'功能：电子签名初始化
    If gobjESign Is Nothing Then
        On Error Resume Next
        Err.Clear

        Set gobjESign = CreateObject("zl9ESign.clsESign")
        If Err <> 0 Then Err.Clear
        On Error GoTo 0
        If Not gobjESign Is Nothing Then
            Call gobjESign.Initialize(gcnOracle, glngSys)
        End If
    End If
    InitESign = Not gobjESign Is Nothing
End Function

Private Sub cmbLevel_Click()
    cmdOK.Enabled = (Mid(Me.cmbLevel.Text, 1, 1) > 0)
End Sub

Private Sub cmdCanCel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Validation Then
        mstr状态 = cmbLevel.ListIndex & "|" & chkEsign.Value
        
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
