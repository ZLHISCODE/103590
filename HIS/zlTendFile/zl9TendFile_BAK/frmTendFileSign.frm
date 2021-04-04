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
      Left            =   1365
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
      Picture         =   "frmTendFileSign.frx":000C
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
Private Sign As cEPRSign                    '签名对象

Private objESign As Object                  '电子签名接口部件
Private lngCertID As Long                   '证书ID
Private lngPassType As Long                 '密码验证规则（系统参数） 0-密码；1－数字；2－两者皆可
Private mbln审签 As Boolean                 '是否审签
Private mlngCur As Long, mlngLast As Long   '当前人员级别，审签人级别

Private mlng文件ID As Long
Private mstrSource As String                 '数字签名的源字符串
Private mstr状态 As String
Private mstrPrivs As String

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
Public Function ShowMe(ByVal objParent As Object, ByVal strPrivs As String, ByVal lng文件ID As Long, ByVal intLevel As Integer, _
    ByVal sSource As String, ByVal bln审签 As Boolean, Optional str状态 As String, Optional str错误 As String) As cEPRSign
    '******************************************************************************************************************
    '功能： 显示签名窗体
    '参数： edtThis     :IN     编辑器控件
    '       fParent     :IN     父窗体
    '       mstrSource   :IN     数字签名的源字符串（从文本中提取，去掉签名提纲）
    '       str状态     :IN     用于连续签名时传入，避免频繁弹出签名窗体
    '       str审签人   :IN     审签时传入上次审签人姓名，以便核实审签权限
    '******************************************************************************************************************
    
    Set Sign = New cEPRSign
    Set frmParent = objParent
    mstrSource = sSource
    mstr状态 = str状态
    mbln审签 = bln审签
    mlngLast = intLevel
    mlng文件ID = lng文件ID
    mstrPrivs = strPrivs
    
    '根据用户的签名级别来初始化“签名级别＂
    Call GetUserLevel(glngUserId)           '获取用户签名级别
    
    '审签则加入比上次级别高的;平签则只能以上次相同级别进行
    If bln审签 Or mlngLast = 未定义 Then
        If Not (mlngCur < mlngLast) Then
            str错误 = "您要超过本记录的签名者或上次审签者的级别才能审签！"
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
            str错误 = "您至少要达到上次签名者的级别才能签名！"
            Unload Me
            Exit Function
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
    
    '读取当前签名方式（系统参数26）
    lngPassType = Val(Mid(zlDatabase.GetPara(26, glngSys), 4, 1))     '门诊,住院,医技,护理 (1111),为空默认采用密码模式
    chkEsign.Value = Val(zlDatabase.GetPara("护理数字签名", glngSys, 1255, "0"))
    
    Call RefControls
    Call RestoreState
    
    If mstr状态 <> "" Then
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
    
    Err = 0: On Error GoTo errHand
    mlngCur = 未定义
    '级别是反着的，1正高最大，所以，判断值必须小于审签人的级别，否则不允许签名
    
    '取当前操作员的级别
    gstrSQL = "select /*+ RULE */ 聘任技术职务 from 人员表 p where ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mRichEPR", lngUserID)
    If Not rs.EOF Then
        mlngCur = NVL(rs("聘任技术职务"), 未定义)
    End If
errHand:
    Exit Sub
End Sub

Private Sub RestoreState()
    Dim arrData
    
    If mstr状态 <> "" Then
        arrData = Split(mstr状态, "|")
        cmbLevel.ListIndex = arrData(0)
        chkEsign.Value = arrData(1)
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
    Dim strUserName As String, lngUserID As Long, strSign As String, str时间戳 As String
    
    If chkEsign.Value = vbChecked Then
        '数字签名
        Err.Clear
        On Error Resume Next
        If objESign Is Nothing Then
            Set objESign = CreateObject("zl9ESign.clsESign")
            If Err <> 0 Then Err = 0: strSign = ""
        End If
        If Not objESign Is Nothing Then
            Call objESign.Initialize(gcnOracle, glngSys)
        End If
        lngCertID = 0
        strSign = objESign.signature(mstrSource, UCase(gcnOracle.Properties(23)), lngCertID, str时间戳) '返回签名信息,lngCertID返回签名使用的证书记录ID
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
    Select Case lngPassType
    Case 0
        '密码签名
        chkEsign.Value = vbUnchecked
        chkEsign.Visible = False
    Case 1
        '1－数字
        chkEsign.Value = vbChecked
        chkEsign.Visible = True
        chkEsign.Enabled = False
    Case 2
        '2－两者皆可
    End Select
End Sub

Private Sub cmbLevel_Click()
    cmdOK.Enabled = (Mid(Me.cmbLevel.Text, 1, 1) > 0)
End Sub

Private Sub cmdCancel_Click()
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

Private Sub Form_Unload(Cancel As Integer)
    If lngPassType = 2 Then
        Call zlDatabase.SetPara("护理数字签名", chkEsign.Value, glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    End If
End Sub
