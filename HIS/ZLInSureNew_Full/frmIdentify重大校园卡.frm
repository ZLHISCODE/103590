VERSION 5.00
Begin VB.Form frmIdentify重大校园卡 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "用户卡信息"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7305
   Icon            =   "frmIdentify重大校园卡.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmd验卡 
      Caption         =   "重读(&R)"
      Height          =   350
      Left            =   240
      TabIndex        =   37
      Top             =   4185
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   1
      Left            =   -30
      TabIndex        =   41
      Top             =   630
      Width           =   8340
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6090
      TabIndex        =   39
      Top             =   4185
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4710
      TabIndex        =   38
      Top             =   4185
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   0
      Left            =   -15
      TabIndex        =   40
      Top             =   4020
      Width           =   8340
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "年龄"
      Height          =   180
      Index           =   10
      Left            =   5370
      TabIndex        =   47
      Top             =   2130
      Width           =   360
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   10
      Left            =   5790
      TabIndex        =   46
      Top             =   2070
      Width           =   1335
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "出生日期"
      Height          =   180
      Index           =   8
      Left            =   5010
      TabIndex        =   45
      Top             =   1725
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   8
      Left            =   5790
      TabIndex        =   44
      Top             =   1665
      Width           =   1335
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "性别"
      Height          =   180
      Index           =   7
      Left            =   3045
      TabIndex        =   43
      Top             =   1725
      Width           =   360
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   7
      Left            =   3450
      TabIndex        =   42
      Top             =   1665
      Width           =   1335
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   20
      Left            =   990
      TabIndex        =   36
      ToolTipText     =   "日交易累计金额"
      Top             =   3645
      Width           =   1335
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "等待时间"
      Height          =   180
      Index           =   20
      Left            =   225
      TabIndex        =   35
      Top             =   3705
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   13
      Left            =   5790
      TabIndex        =   34
      ToolTipText     =   "日交易累计金额"
      Top             =   2475
      Width           =   1335
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "日累计额"
      Height          =   180
      Index           =   13
      Left            =   5010
      TabIndex        =   33
      Top             =   2535
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   19
      Left            =   5790
      TabIndex        =   32
      ToolTipText     =   "上次交易终端号"
      Top             =   3255
      Width           =   1335
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "上次终端号"
      Height          =   180
      Index           =   19
      Left            =   4830
      TabIndex        =   31
      Top             =   3315
      Width           =   900
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   18
      Left            =   3450
      TabIndex        =   30
      ToolTipText     =   "上次交易时间"
      Top             =   3255
      Width           =   1335
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "上次时间"
      Height          =   180
      Index           =   18
      Left            =   2685
      TabIndex        =   29
      Top             =   3315
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   17
      Left            =   990
      TabIndex        =   28
      ToolTipText     =   "上次交易金额"
      Top             =   3255
      Width           =   1365
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "上次流水额"
      Height          =   180
      Index           =   17
      Left            =   45
      TabIndex        =   27
      Top             =   3315
      Width           =   900
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "钱包1余额"
      Height          =   180
      Index           =   14
      Left            =   135
      TabIndex        =   21
      Top             =   2925
      Width           =   810
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   14
      Left            =   990
      TabIndex        =   22
      Top             =   2865
      Width           =   1365
   End
   Begin VB.Label lblInfor 
      AutoSize        =   -1  'True
      Caption         =   "通过IC卡验证人员身份，并将验证结果信息显示出来。"
      Height          =   180
      Left            =   600
      TabIndex        =   0
      Top             =   390
      Width           =   4320
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   30
      Picture         =   "frmIdentify重大校园卡.frx":000C
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "上次流水号"
      Height          =   180
      Index           =   16
      Left            =   4830
      TabIndex        =   25
      Top             =   2925
      Width           =   900
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   16
      Left            =   5790
      TabIndex        =   26
      ToolTipText     =   "上次交易流水"
      Top             =   2865
      Width           =   1335
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   15
      Left            =   3450
      TabIndex        =   24
      Top             =   2865
      Width           =   1335
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "钱包2余额"
      Height          =   180
      Index           =   15
      Left            =   2595
      TabIndex        =   23
      Top             =   2925
      Width           =   810
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   9
      Left            =   990
      TabIndex        =   20
      Top             =   2070
      Width           =   3795
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "身分证号"
      Height          =   180
      Index           =   9
      Left            =   225
      TabIndex        =   19
      Top             =   2130
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   12
      Left            =   3450
      TabIndex        =   18
      Top             =   2475
      Width           =   1335
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "卡序列号"
      Height          =   180
      Index           =   12
      Left            =   2685
      TabIndex        =   17
      Top             =   2535
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   11
      Left            =   990
      TabIndex        =   16
      Top             =   2475
      Width           =   1335
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "注册日期"
      Height          =   180
      Index           =   11
      Left            =   225
      TabIndex        =   15
      Top             =   2535
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   6
      Left            =   990
      TabIndex        =   14
      Top             =   1665
      Width           =   1365
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   1
      Left            =   3450
      TabIndex        =   4
      Top             =   900
      Width           =   1335
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "姓名"
      Height          =   180
      Index           =   6
      Left            =   585
      TabIndex        =   13
      Top             =   1725
      Width           =   360
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   5
      Left            =   5790
      TabIndex        =   12
      Top             =   1290
      Width           =   1335
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "卡有效期"
      Height          =   180
      Index           =   5
      Left            =   5010
      TabIndex        =   11
      Top             =   1350
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   4
      Left            =   3450
      TabIndex        =   10
      Top             =   1290
      Width           =   1335
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "卡消费分组"
      Height          =   180
      Index           =   4
      Left            =   2505
      TabIndex        =   9
      Top             =   1350
      Width           =   900
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   3
      Left            =   990
      TabIndex        =   8
      Top             =   1290
      Width           =   1365
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "校园卡号"
      Height          =   180
      Index           =   3
      Left            =   225
      TabIndex        =   7
      Top             =   1350
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   2
      Left            =   5790
      TabIndex        =   6
      Top             =   900
      Width           =   1335
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "卡流水号"
      Height          =   180
      Index           =   2
      Left            =   5010
      TabIndex        =   5
      Top             =   960
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "卡类型"
      Height          =   180
      Index           =   1
      Left            =   2865
      TabIndex        =   3
      Top             =   960
      Width           =   540
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   0
      Left            =   990
      TabIndex        =   2
      Top             =   900
      Width           =   1365
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "串口设备号"
      Height          =   180
      Index           =   0
      Left            =   45
      TabIndex        =   1
      Top             =   960
      Width           =   900
   End
End
Attribute VB_Name = "frmIdentify重大校园卡"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnFirst  As Boolean
Dim mstrReturn As String    '返回信息串
Dim mlng病人ID As Long
'mbytType-识别类型，0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐
Dim mbytType As Byte
Dim mblnOK As Boolean
Dim mbln自动挂号 As Boolean
Private Function IsValid() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:数据验证
    '--入参数:
    '--出参数:
    '--返  回:验证成功返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    IsValid = False
    
    '检查病人状态
    Dim lng病人ID As Long
    gstrSQL = "select 病人id,nvl(当前状态,0) as 状态 from 保险帐户 where 险类=[1] and 医保号=[2]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_重大校园卡, g病人身份_重大校园卡.卡号)
    If mbytType <> 4 Then   '不是住院结算时，需验证当前状态
        If rsTemp.RecordCount > 0 Then
            If rsTemp("状态") > 0 Then
                MsgBox "该病人已经在院，不能通过身份验证。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    Else
        '住院结算时,需处理是否为同一个人
        If rsTemp.EOF Then
            ShowMsgbox "在保险帐户中不存在当前病人!"
            Exit Function
        Else
            lng病人ID = Nvl(rsTemp!病人ID, 0)
            If mlng病人ID <> lng病人ID Then
                ShowMsgbox "虚假结帐的当前病人与身份验证的病人不一致!"
                Exit Function
            End If
        End If
    End If
    IsValid = True
End Function

Private Sub cmdCancel_Click()
    mstrReturn = ""
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strTmp As String
    Dim strTmp1 As String
    
    '验证数据
    If IsValid = False Then Exit Sub
    
    '确定相关返回串
    
    '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);
    '8.中心代码;9.顺序号;10人员身份;11帐户余额;12当前状态;13病种ID;14在职(0,1);15退休证号;16年龄段;17灰度级
    '18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计

    '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);8《病人ID》
    '9中心;10.顺序号;11人员身份;12帐户余额;13当前状态;14病种ID;15在职(0,1);16退休证号;17年龄段;18灰度级
    '19帐户增加累计,20帐户支出累计,21进入统筹累计,22统筹报销累计,23住院次数累计;24就诊类型 (1、急诊门诊);25开单科室
    
    
    mstrReturn = ""
    With g病人身份_重大校园卡
        strTmp = .卡号                         '0卡号
        strTmp = strTmp & ";" & .卡流水号    '1医保号
        strTmp = strTmp & ";"               '2密码
        strTmp = strTmp & ";" & .姓名       '3姓名
        strTmp = strTmp & ";" & .性别       '4性别
        strTmp = strTmp & ";" & .出生日期   '5出生日期
        strTmp = strTmp & ";" & .身份证号   '6身份证
        strTmp = strTmp & ";" & .卡类型   '7单位名称(编码)
        
        strTmp1 = ""
        strTmp1 = strTmp1 & ";"    '8中心代码
        strTmp1 = strTmp1 & ";" & .卡消费分组    '9顺序号
        strTmp1 = strTmp1 & ";"       '10人员身份,存的是转诊单号
        strTmp1 = strTmp1 & ";" & (.电子钱包1余额 + .电子钱包2余额) / 100     '11帐户余额
        strTmp1 = strTmp1 & ";0"               '12当前状态
        strTmp1 = strTmp1 & ";"               '13病种ID
        strTmp1 = strTmp1 & ";"   '.就诊分类  '14在职(0,1)
        strTmp1 = strTmp1 & ";"   '15退休证号,目前我存的是补助个人帐户余额
        strTmp1 = strTmp1 & ";" & IIf(.年龄 = 0, "", .年龄) '16年龄段
        strTmp1 = strTmp1 & ";"     '17灰度级,存的就诊分类编码
        strTmp1 = strTmp1 & ";"         '18帐户增加累计
        strTmp1 = strTmp1 & ";"        '19帐户支出累计
        strTmp1 = strTmp1 & ";"  '20进入统筹累计
        strTmp1 = strTmp1 & ";"  '21统筹报销累计
        strTmp1 = strTmp1 & ";"        '22住院次数累计
        
    End With
    
    mlng病人ID = BuildPatiInfo(0, strTmp & strTmp1, mlng病人ID, TYPE_重大校园卡)
    
    '返回格式:中间插入病人ID
    If mlng病人ID > 0 Then
        mstrReturn = strTmp & ";" & mlng病人ID & strTmp1
    Else
        Unload Me
    End If
    '存储校园卡信息
    '过程:zl_校园卡信息_Insert(病人ID_IN,险类_IN,中心_IN,卡流水号_IN,卡号_IN,卡消费分组_IN,卡有效期_IN,姓名_IN,
    '   注册日期_IN,身份证号_IN,电子钱包1余额_IN,电子钱包2余额_IN,卡固有序列号_IN,上次交易流水号_IN,上次交易金额_IN
    '   上次交易时间_IN,上次交易终端号_IN,卡等待时间_IN,日交易累计金额_IN
    
    
    strTmp = "zl_校园卡信息_Insert("
    With g病人身份_重大校园卡
        strTmp = strTmp & _
        mlng病人ID & "," & _
        TYPE_重大校园卡 & "," & _
        0 & "," & _
        .卡流水号 & "," & _
        .卡号 & "," & _
        .卡消费分组 & "," & _
        IIf(.卡有效期 = "", "NULL", "'" & .卡有效期 & "'") & "," & _
        IIf(.姓名 = "", "NULL", "'" & .姓名 & "'") & "," & _
        IIf(.注册日期 = "", "NULL", "'" & .注册日期 & "'") & "," & _
        IIf(.身份证号 = "", "NULL", "'" & .身份证号 & "'") & "," & _
        .电子钱包1余额 & "," & _
        .电子钱包2余额 & "," & _
        .卡固有序号 & "," & _
        .上次交易流水号 & "," & _
        .上次交易金额 & "," & _
        .上次交易时间 & "," & _
        .上次交易终端号 & "," & _
        .卡等待时间 & "," & _
        .日交易累计金额 & ")"
    End With
    Err = 0
    On Error GoTo errHand:
    zlDatabase.ExecuteProcedure strTmp, Me.Caption
     Unload Me
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmd验卡_Click()
    cmd验卡.Enabled = False
    SetCtlEn False
    mblnOK = ReadCard
    If mblnOK And mbln自动挂号 Then
        cmdOK_Click
        Unload Me
        Exit Sub
    End If
    SetCtlEn True
    cmd验卡.Enabled = True
End Sub
Private Sub SetCtlEn(ByVal blnTrue As Boolean)
    cmd验卡.Enabled = blnTrue
    cmdOK.Enabled = blnTrue And mblnOK
    cmdCancel.Enabled = blnTrue
End Sub
Private Sub Form_Activate()
    
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    mblnOK = False
    '进入时读卡
    cmd验卡_Click
End Sub
Private Function ReadCard() As Boolean
    Dim int性别 As Integer
    Dim i As Integer
    ReadCard = False
   
   '验证用户身份
    If GetUserCardInfor() = False Then
        For i = 0 To 20
            lblEdit(i) = ""
        Next
        Exit Function
    End If
    
    Err = 0
    On Error Resume Next
    '给窗体的相关信息赋值
    With g病人身份_重大校园卡
        lblEdit(0) = .串口设备
        lblEdit(1) = .卡类型
        lblEdit(2) = .卡流水号
        lblEdit(3) = .卡号
        lblEdit(4) = .卡消费分组
        lblEdit(5) = .卡有效期
        lblEdit(6) = .姓名
        lblEdit(7) = .性别
        lblEdit(8) = .出生日期
        lblEdit(9) = .身份证号
        lblEdit(10) = .年龄
        lblEdit(11) = .注册日期
        lblEdit(12) = .卡固有序号
        lblEdit(13) = .日交易累计金额 / 100
        lblEdit(14) = .电子钱包1余额 / 100
        lblEdit(15) = .电子钱包2余额 / 100
        lblEdit(16) = .上次交易流水号
        lblEdit(17) = .上次交易金额 / 100
        lblEdit(18) = .上次交易时间
        lblEdit(19) = .上次交易终端号
        lblEdit(20) = .卡等待时间
    End With
    ReadCard = True
End Function

Public Function GetPatient(ByVal bytType As Byte, Optional ByVal lng病人ID As Long = 0, Optional bln自动挂号 As Boolean = False) As String
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:获取病人的相关信息
    '--入参数:bytType-类型(mbytType-识别类型，0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐)
    '         lng病人ID-病人ID
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    mstrReturn = ""
    mlng病人ID = lng病人ID
    mbytType = bytType
    mbln自动挂号 = bln自动挂号
    Me.Show 1
    GetPatient = mstrReturn
End Function

Private Sub Form_Load()
    mblnFirst = True
End Sub



