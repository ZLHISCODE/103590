VERSION 5.00
Begin VB.Form frmIdentify沈阳 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "医保病人身份识别"
   ClientHeight    =   6780
   ClientLeft      =   1665
   ClientTop       =   2985
   ClientWidth     =   9345
   Icon            =   "frmIdentify沈阳.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   9345
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "补充信息"
      Height          =   2115
      Left            =   240
      TabIndex        =   38
      Top             =   4110
      Width           =   8925
      Begin VB.Frame Frame3 
         Caption         =   "规定病信息(&F)"
         Height          =   1575
         Left            =   4650
         TabIndex        =   41
         Top             =   390
         Width           =   4095
         Begin VB.TextBox txt申请序号 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1410
            MaxLength       =   12
            TabIndex        =   43
            Top             =   300
            Width           =   2415
         End
         Begin VB.TextBox txt疾病编码 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1410
            MaxLength       =   20
            TabIndex        =   45
            Top             =   690
            Width           =   2415
         End
         Begin VB.TextBox txt疾病名称 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1410
            MaxLength       =   60
            TabIndex        =   47
            Top             =   1080
            Width           =   2415
         End
         Begin VB.Label lbl申请序号 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "申请序号(&D)"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   360
            TabIndex        =   42
            Top             =   360
            Width           =   990
         End
         Begin VB.Label lbl疾病编码 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "疾病编码(&B)"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   360
            TabIndex        =   44
            Top             =   750
            Width           =   990
         End
         Begin VB.Label lbl疾病名称 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "疾病名称(&X)"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   360
            TabIndex        =   46
            Top             =   1140
            Width           =   990
         End
      End
      Begin VB.TextBox txt冻结基金信息 
         Enabled         =   0   'False
         Height          =   1410
         Left            =   300
         MaxLength       =   12
         MultiLine       =   -1  'True
         TabIndex        =   40
         Tag             =   "persfundcon"
         Top             =   540
         Width           =   3885
      End
      Begin VB.Label lbl冻结基金信息 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "冻结基金信息(&U)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   300
         TabIndex        =   39
         Top             =   300
         Width           =   1350
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6510
      TabIndex        =   48
      Top             =   6360
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7860
      TabIndex        =   49
      Top             =   6360
      Width           =   1100
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "修改密码(&M)"
      Height          =   405
      Left            =   240
      TabIndex        =   50
      Top             =   6330
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "病人基本信息"
      Height          =   3885
      Left            =   180
      TabIndex        =   0
      Top             =   150
      Width           =   8985
      Begin VB.CommandButton cmd疾病信息 
         Caption         =   "…"
         Height          =   300
         Left            =   8460
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   3450
         Width           =   285
      End
      Begin VB.TextBox txt疾病信息 
         Height          =   300
         Left            =   1530
         TabIndex        =   37
         Top             =   3450
         Width           =   6915
      End
      Begin VB.ComboBox cbo业务类型 
         Height          =   300
         Left            =   1530
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   720
         Width           =   1485
      End
      Begin VB.TextBox txt密码 
         Height          =   300
         Left            =   1530
         MaxLength       =   10
         TabIndex        =   2
         Top             =   330
         Width           =   1485
      End
      Begin VB.TextBox txt住院次数 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Left            =   1530
         MaxLength       =   2
         TabIndex        =   17
         Top             =   2670
         Width           =   525
      End
      Begin VB.TextBox txt帐户余额 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Left            =   1530
         MaxLength       =   18
         TabIndex        =   14
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox txt参保人单位 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5970
         MaxLength       =   100
         TabIndex        =   35
         Tag             =   "corp_id|corp_name"
         Top             =   3060
         Width           =   2775
      End
      Begin VB.TextBox txt异地安置城市 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5970
         MaxLength       =   100
         TabIndex        =   33
         Tag             =   "city_code|city_name"
         Top             =   2670
         Width           =   2775
      End
      Begin VB.TextBox txt职务 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5970
         MaxLength       =   10
         TabIndex        =   23
         Tag             =   "position_name"
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txt特殊工种 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5970
         MaxLength       =   50
         TabIndex        =   31
         Tag             =   "work_type|work_type_name"
         Top             =   2280
         Width           =   2775
      End
      Begin VB.TextBox txt特殊照顾人群 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5970
         MaxLength       =   50
         TabIndex        =   29
         Tag             =   "special_code|special_name"
         Top             =   1890
         Width           =   2775
      End
      Begin VB.TextBox txt公务员级别 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5970
         MaxLength       =   50
         TabIndex        =   27
         Tag             =   "official_code|official_name"
         Top             =   1500
         Width           =   2775
      End
      Begin VB.TextBox txt民族 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5970
         MaxLength       =   50
         TabIndex        =   25
         Tag             =   "folk_code|folk_name"
         Top             =   1110
         Width           =   2775
      End
      Begin VB.TextBox txt人员类别 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5970
         MaxLength       =   50
         TabIndex        =   21
         Tag             =   "pers_type|pers_name"
         Top             =   330
         Width           =   2775
      End
      Begin VB.TextBox txt身份证号 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1530
         MaxLength       =   18
         TabIndex        =   19
         Tag             =   "idcard"
         Top             =   3060
         Width           =   2715
      End
      Begin VB.ComboBox cbo性别 
         Enabled         =   0   'False
         Height          =   300
         Left            =   3420
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1890
         Width           =   825
      End
      Begin VB.TextBox txt姓名 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1530
         MaxLength       =   10
         TabIndex        =   10
         Tag             =   "name"
         Top             =   1890
         Width           =   1035
      End
      Begin VB.TextBox txt医保号 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1530
         MaxLength       =   30
         TabIndex        =   8
         Tag             =   "insr_code"
         Top             =   1500
         Width           =   2715
      End
      Begin VB.TextBox txt个人编号 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1530
         MaxLength       =   8
         TabIndex        =   6
         Tag             =   "indi_id"
         Top             =   1110
         Width           =   1485
      End
      Begin VB.Label lbl疾病信息 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "疾病信息(&J)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   480
         TabIndex        =   36
         Top             =   3510
         Width           =   990
      End
      Begin VB.Label lbl业务类型 
         Caption         =   "业务类型(&E)"
         Height          =   225
         Left            =   465
         TabIndex        =   3
         Top             =   758
         Width           =   1005
      End
      Begin VB.Label lbl密码 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "密码(&R)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   840
         TabIndex        =   1
         Top             =   390
         Width           =   630
      End
      Begin VB.Label lbl单位 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "元"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3060
         TabIndex        =   15
         Top             =   2340
         Width           =   180
      End
      Begin VB.Label lbl住院次数 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "住院次数(&Y)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   480
         TabIndex        =   16
         Top             =   2730
         Width           =   990
      End
      Begin VB.Label lbl帐户余额 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "帐户余额(&Q)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   480
         TabIndex        =   13
         Top             =   2340
         Width           =   990
      End
      Begin VB.Label lbl参保人单位 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "参保人单位(&I)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4740
         TabIndex        =   34
         Top             =   3120
         Width           =   1170
      End
      Begin VB.Label lbl异地安置城市 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "异地安置城市(&A)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4560
         TabIndex        =   32
         Top             =   2730
         Width           =   1350
      End
      Begin VB.Label lbl职务 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "职务(&Z)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   5280
         TabIndex        =   22
         Top             =   780
         Width           =   630
      End
      Begin VB.Label lbl特殊工种 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "特殊工种(&G)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4920
         TabIndex        =   30
         Top             =   2340
         Width           =   990
      End
      Begin VB.Label lbl特殊照顾人群 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "特殊照顾人群(&T)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4560
         TabIndex        =   28
         Top             =   1950
         Width           =   1350
      End
      Begin VB.Label lbl公务员级别 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "公务员级别(&S)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4740
         TabIndex        =   26
         Top             =   1560
         Width           =   1170
      End
      Begin VB.Label lbl民族 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "民族(&R)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   5280
         TabIndex        =   24
         Top             =   1170
         Width           =   630
      End
      Begin VB.Label lbl人员类别 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "人员类别(&L)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4920
         TabIndex        =   20
         Top             =   390
         Width           =   990
      End
      Begin VB.Label lbl身份证号 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "身份证号(&K)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   480
         TabIndex        =   18
         Top             =   3120
         Width           =   990
      End
      Begin VB.Label lbl性别 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "性别(&S)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2730
         TabIndex        =   11
         Top             =   1950
         Width           =   630
      End
      Begin VB.Label lbl姓名 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "姓名(&N)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   840
         TabIndex        =   9
         Top             =   1950
         Width           =   630
      End
      Begin VB.Label lbl医保号 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "医保号(&Y)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   660
         TabIndex        =   7
         Top             =   1560
         Width           =   810
      End
      Begin VB.Label lbl个人编号 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "个人编号(&P)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   480
         TabIndex        =   5
         Top             =   1170
         Width           =   990
      End
   End
End
Attribute VB_Name = "frmIdentify沈阳"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'编译常量不能定义成公共的，必须在使用到的地方单独定义，在编译时统一修改
#Const gverControl = 99  ' 0-不支持动态医保(9.19以前),1-支持动态医保无附加参数(9.22以前) , _
'    2-解决了虚拟结算与正式结算结果不一致;结算作废与原始结算结果不一致;门诊收费死锁的问题;3-公共部件增加GetNextNO();
'    99-所有交易增加附加参数(最新版)

Private mstrReturn As String
Private mstr新密码 As String
Private mlng病人ID As Long
Private mbytType As Byte
Private mblnStart As Boolean
Private mstr卡号 As String
Private mbln允许门诊 As Boolean
Private mbln多病种 As Boolean
Private mrs病种 As New ADODB.Recordset
Private Const madLongVarCharDefault As Integer = 10          '字符型字段缺省长度
Private Const madDoubleDefault As Integer = 18               '数字型字段缺省长度
Private Const madDbDateDefault As Integer = 20               '日期型字段缺省长度

'--------------------------------------------------
'门诊规定病的病人只能从接口返回的已审批的病种中选择
'其他病人可以从所有病种中选择

Public Function GetPatient(ByVal bytType As Byte, Optional ByVal lng病人ID As Long = 0) As String
    mstrReturn = ""
    mstr新密码 = ""
    mlng病人ID = lng病人ID
    mbytType = bytType
    Me.Show 1
    
    GetPatient = mstrReturn
End Function

Private Sub cbo业务类型_Click()
    If Not mblnStart Then Exit Sub
    Call txt密码_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdChange_Click()
    mstr新密码 = frm修改密码.ChangePassword(txt密码.Text)
End Sub

Private Sub cmdOK_Click()
    Dim bln在院 As Boolean
    Dim str出生日期 As String
    Dim intDays As Integer
    Dim strBeginDate As String
    Dim strIdentify As String, strAddition As String
    Dim lng疾病ID As Long
    Dim str顺序号 As String
    Dim rsTemp As New ADODB.Recordset
    
    If Trim(txt个人编号.Text) = "" Then
        MsgBox "还未获取病人的基本信息，请输入密码后按回车！", vbInformation, gstrSysName
        txt密码.SetFocus
        Exit Sub
    End If
    
    '检查病人状态
    gstrSQL = "select nvl(当前状态,0) as 状态,顺序号 from 保险帐户 where 险类=[1] and 医保号=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_沈阳市, CStr(txt医保号.Text))
    
    If rsTemp.RecordCount > 0 Then
        If rsTemp("状态") > 0 Then
            str顺序号 = Nvl(rsTemp!顺序号)
            bln在院 = True
            If Not mbln允许门诊 Then
                MsgBox "该病人已经在院，不能通过身份验证。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    End If
    
    '转换出生日期
    str出生日期 = "1980-01-01"
    
    '必须选择病种信息
    If txt疾病信息.Tag = "" Then
        MsgBox "请为该参保病人选择疾病编码信息！", vbInformation, gstrSysName
        txt疾病信息.SetFocus
        Exit Sub
    End If
    gstrSQL = "Select ID From 保险病种 Where 编码=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取保险病种", CStr(txt疾病信息.Tag))
    If Not rsTemp.EOF Then
        lng疾病ID = rsTemp!ID
    End If
    
    '构成字符串
    '建立病人档案信息，传入格式：
    '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);
    '8.中心代码;9.顺序号;10人员身份;11帐户余额;12当前状态;13病种ID;14在职(0,1);15退休证号;16年龄段;17灰度级
    '18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计
    strIdentify = txt个人编号.Text                              '0卡号
    strIdentify = strIdentify & ";" & txt医保号.Text            '1医保号
    strIdentify = strIdentify & ";" & txt密码.Text              '2密码
    strIdentify = strIdentify & ";" & txt姓名.Text              '3姓名
    strIdentify = strIdentify & ";" & cbo性别.Text              '4性别
    strIdentify = strIdentify & ";" & str出生日期               '5出生日期
    strIdentify = strIdentify & ";" & txt身份证号.Text          '6身份证
    strIdentify = strIdentify & ";" & txt参保人单位.Text        '7.单位名称(编码)
    strAddition = ";0"                                          '8.中心代码
    strAddition = strAddition & ";" & str顺序号                 '9.顺序号
    strAddition = strAddition & ";" & txt人员类别.Text          '10人员身份
    strAddition = strAddition & ";" & Val(txt帐户余额.Text)     '11帐户余额
    strAddition = strAddition & ";0"                            '12当前状态
    strAddition = strAddition & ";" & lng疾病ID                 '13病种ID
    strAddition = strAddition & ";1"                            '14在职(1,2,3)
    strAddition = strAddition & ";"                             '15退休证号
    strAddition = strAddition & ";"                             '16年龄段
    strAddition = strAddition & ";" & cbo业务类型.ItemData(cbo业务类型.ListIndex)                           '17灰度级
    strAddition = strAddition & ";" & Val(txt帐户余额.Text)     '18帐户增加累计
    strAddition = strAddition & ";0"                            '19帐户支出累计
    strAddition = strAddition & ";0"                            '20上年工资总额
    strAddition = strAddition & ";" & Val(txt住院次数.Text)     '21住院次数累计
    
    mlng病人ID = BuildPatiInfo(0, strIdentify & strAddition, mlng病人ID, TYPE_沈阳市)
    '返回格式:中间插入病人ID
    If mlng病人ID > 0 Then
        mstrReturn = strIdentify & ";" & mlng病人ID & strAddition
    End If
    
    '如果是门诊业务，如果在系统规定的挂号有效天数内没有挂号记录，则不允许收费
    If mbytType = 0 And bln在院 = False Then
        #If gverControl >= 4 Then
            intDays = Val(zlDatabase.GetPara("挂号有效天数", glngSys, , 0)) - 1
        #Else
            intDays = Val(GetPara("挂号有效天数", glngSys, , , 0)) - 1
        #End If
        
        '如果挂号有效天数为零，表示门诊收费前，可以不挂号
        If intDays > -1 Then
            strBeginDate = Format(DateAdd("d", IIf(intDays = -1, 30, intDays) * -1, zlDatabase.Currentdate()), "yyyy-MM-dd")
            
            '取该段时间内，有无挂号记录
            gstrSQL = " Select 1 From 门诊费用记录" & _
                      " Where 记录性质=4 And 记录状态=1 And 病人ID=[1] And 发生时间>[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取规定的时间内，该病人有无挂号记录", mlng病人ID, CDate(strBeginDate))
            If rsTemp.EOF Then
                MsgBox "该病人还没挂号，不能进行门诊身份登记！", vbInformation, gstrSysName
                mstrReturn = ""
                Exit Sub
            End If
        End If
    End If
    
    If mbytType = 1 Then
        gstrSQL = "zl_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_沈阳市 & ",'业务类型','''" & cbo业务类型.ItemData(cbo业务类型.ListIndex) & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存住院业务类型")
    End If
    
    If mstr新密码 <> "" Then
        '调用修改密码接口，不成功给出提示并继续
'        1   card_no    医保卡卡号  20  否
'        2   card_type  卡证类型    12  否  "01"：医保卡
'        3   password   原密码      6   否
'        4   newpassword新密码      6   否
        If 调用接口_准备_沈阳市(Function_沈阳市.其他_修改密码) Then
            '写入口参数
            gstrField_沈阳市 = "card_no||card_type||password||newpassword"
            gstrValue_沈阳市 = mstr卡号 & "||" & "01" & "||" & txt密码.Text & "||" & mstr新密码
            Call 调用接口_写入口参数_沈阳市(1)
            If 调用接口_执行_沈阳市() Then
                '更新个人帐户中的信息
                gstrSQL = "zl_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_沈阳市 & ",'密码','''" & mstr新密码 & "''')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "更新密码")
            Else
                MsgBox "密码修改失败，仍可继续操作！", vbInformation, gstrSysName
            End If
        End If
    End If
    
    gCominfo_沈阳市.业务类型 = cbo业务类型.ItemData(cbo业务类型.ListIndex)
    gCominfo_沈阳市.疾病编码 = txt疾病信息.Tag
    gCominfo_沈阳市.个人编号 = txt个人编号.Text
    gCominfo_沈阳市.帐户余额 = Val(txt帐户余额.Text)
    
    Unload Me
End Sub

Private Sub cmd疾病信息_Click()
    Dim bln特殊病 As Boolean
    Dim rs病种 As ADODB.Recordset
    bln特殊病 = (Me.cbo业务类型.ItemData(Me.cbo业务类型.ListIndex) = 业务分类_沈阳市.门诊规定病)
    
    If Not bln特殊病 Then
        gstrSQL = " Select A.ID,A.编码,A.名称,A.简码 " & _
                " From 保险病种 A where A.险类=[1]"
        Set rs病种 = zlDatabase.OpenSQLRecord(gstrSQL, "身份验证", TYPE_沈阳市)
        If rs病种.RecordCount > 0 Then
            If frmListSel.ShowSelect(TYPE_沈阳市, rs病种, "ID", "医保病种选择", "请选择医保病种：") = True Then
                txt疾病信息.Tag = rs病种!编码
                txt疾病信息.Text = "(" & rs病种!编码 & ")" & rs病种!名称
                lbl疾病信息.Tag = txt疾病信息.Text '用于恢复显示
            End If
        End If
    Else
        If mrs病种.RecordCount > 0 Then
            If frmListSel.ShowSelect(TYPE_沈阳市, mrs病种, "ID", "特殊病种选择", "请选择特定的医保病种：") = True Then
                txt疾病信息.Tag = mrs病种!编码
                txt疾病信息.Text = "(" & mrs病种!编码 & ")" & mrs病种!名称
                lbl疾病信息.Tag = txt疾病信息.Text '用于恢复显示
            End If
        End If
    End If
    cmdOK.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    mblnStart = False
    
    With cbo性别
        .Clear
        .AddItem "女"
        .AddItem "男"
        .ListIndex = 1
    End With
    
    With cbo业务类型
        .Clear
        If mbytType = 0 Or mbytType = 2 Then
            .AddItem "普通门诊"
            .ItemData(.NewIndex) = 业务分类_沈阳市.普通门诊
            .AddItem "门诊规定病"
            .ItemData(.NewIndex) = 业务分类_沈阳市.门诊规定病
            .AddItem "门诊急救"
            .ItemData(.NewIndex) = 业务分类_沈阳市.门诊急救
            .AddItem "特治特检"
            .ItemData(.NewIndex) = 业务分类_沈阳市.特治特检
            .AddItem "生育门诊"
            .ItemData(.NewIndex) = 业务分类_沈阳市.生育门诊
            .AddItem "工伤门诊"
            .ItemData(.NewIndex) = 业务分类_沈阳市.工伤门诊
        ElseIf mbytType = 1 Or mbytType = 2 Then
            .AddItem "普通住院"
            .ItemData(.NewIndex) = 业务分类_沈阳市.普通住院
            .AddItem "家庭病床"
            .ItemData(.NewIndex) = 业务分类_沈阳市.家庭病床
            .AddItem "生育住院"
            .ItemData(.NewIndex) = 业务分类_沈阳市.生育住院
            .AddItem "工伤住院"
            .ItemData(.NewIndex) = 业务分类_沈阳市.工伤住院
        ElseIf mbytType = 3 Then
            .AddItem "普通挂号"
            .ItemData(.NewIndex) = 业务分类_沈阳市.普通门诊
        End If
        .ListIndex = 0
        .Enabled = (mbytType <> 2)
    End With
    lbl疾病信息.Enabled = True
    txt疾病信息.Enabled = True
    
    '取住院病人是否允许门诊业务
    gstrSQL = "Select Nvl(参数值,0) 参数值 From 保险参数 Where 序号=7 And 险类=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取住院病人是否允许门诊业务", TYPE_沈阳市)
    If Not rsTemp.EOF Then
        mbln允许门诊 = (rsTemp!参数值 = 1)
    End If
    
    '如果是挂号，则显示缺省诊断0000000-其它
    If mbytType = 3 Then
        gstrSQL = "Select ID,编码,名称 From 保险病种 Where 编码='0000000'"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取门诊挂号缺省病种")
        If rsTemp.EOF Then
            MsgBox "请初始化挂号缺省病种！（(0000000)其它）"
        Else
            txt疾病信息.Tag = rsTemp!编码
            txt疾病信息.Text = "(" & rsTemp!编码 & ")" & rsTemp!名称
            lbl疾病信息.Tag = txt疾病信息.Text '用于恢复显示
        End If
    End If
    
    mblnStart = True
End Sub

Private Sub ClearCons()
    Dim strFields As String
    Dim objTextBox As Control
    strFields = "ID" & "," & adDouble & "," & "18" & "|" & _
                "编码" & "," & adVarChar & "," & "50" & "|" & _
                "名称" & "," & adLongVarChar & "," & "100"
    Call Record_Init(mrs病种, strFields)
    
    For Each objTextBox In Controls
        If UCase(TypeName(objTextBox)) = "TEXTBOX" And Not (objTextBox.Name = "txt密码" Or objTextBox.Name = "txt疾病信息") Then
            objTextBox.Text = ""
        End If
    Next
End Sub

Private Sub txt疾病信息_GotFocus()
    OpenIme ""
    Call zlControl.TxtSelAll(txt疾病信息)
End Sub

Private Sub txt疾病信息_KeyPress(KeyAscii As Integer)
    Dim rsTemp As New ADODB.Recordset, strText As String
    Dim blnReturn As Boolean, lst As ListItem
    Dim bln特殊病 As Boolean
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If txt疾病信息.Text = "" And txt疾病信息.Tag <> "" Then Exit Sub
    bln特殊病 = (Me.cbo业务类型.ItemData(Me.cbo业务类型.ListIndex) = 业务分类_沈阳市.门诊规定病)
    
    On Error GoTo errHandle
    
    strText = txt疾病信息.Text
    If InStr(1, strText, "(") <> 0 Then
        If InStr(1, strText, ")") <> 0 Then
            strText = Mid(strText, 2, InStr(1, strText, ")") - 2)
        End If
    End If
    If Not bln特殊病 Then
        gstrSQL = "Select A.ID,A.编码,A.名称,A.简码" & _
                 "   FROM 保险病种 A WHERE A.险类=[1] And (" & _
                 "A.编码 like [2] || '%' or A.名称 like [2] || '%' or A.简码 like [2] || '%')"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_沈阳市, strText)
        If rsTemp.RecordCount = 0 Then
            MsgBox "不存在该病种，请重新输入！", vbInformation, gstrSysName
            txt疾病信息.Text = lbl疾病信息.Tag
            zlControl.TxtSelAll txt疾病信息
            Exit Sub
        Else
            '出现选择器
            If rsTemp.RecordCount > 1 Then
                '对于字段大于3的，即使只有一条记录把该对话框显示出来，以便让用户得到更多的信息
                blnReturn = frmListSel.ShowSelect(TYPE_沈阳市, rsTemp, "ID", "医保病种选择", "请选择医保病种：")
            Else
                blnReturn = True
            End If
        End If
    Else
        If IsNumeric(strText) Then
            mrs病种.Filter = "编码 Like '" & strText & "*'"
        Else
            mrs病种.Filter = "名称 Like '" & strText & "*'"
        End If
        If mrs病种.RecordCount = 0 Then
            MsgBox "不存在该特殊病种，请重新输入！", vbInformation, gstrSysName
            mrs病种.Filter = 0
            txt疾病信息.Text = lbl疾病信息.Tag
            zlControl.TxtSelAll txt疾病信息
            Exit Sub
        Else
            If mrs病种.RecordCount > 1 Then
                blnReturn = frmListSel.ShowSelect(TYPE_沈阳市, mrs病种, "ID", "特殊病种选择", "请选择特定的医保病种：")
            Else
                blnReturn = True
            End If
        End If
    End If
    
    If blnReturn = False Then
        '记录集中没有可选择的数据
        txt疾病信息.Text = lbl疾病信息.Tag
        zlControl.TxtSelAll txt疾病信息
        If bln特殊病 Then mrs病种.Filter = 0
        Exit Sub
    Else
        '肯定是有记录集的
        If Not bln特殊病 Then
            txt疾病信息.Tag = rsTemp!编码
            txt疾病信息.Text = "(" & rsTemp!编码 & ")" & rsTemp!名称
            lbl疾病信息.Tag = txt疾病信息.Text '用于恢复显示
        Else
            txt疾病信息.Tag = mrs病种!编码
            txt疾病信息.Text = "(" & mrs病种!编码 & ")" & mrs病种!名称
            lbl疾病信息.Tag = txt疾病信息.Text '用于恢复显示
        End If
    End If
    
    If bln特殊病 Then mrs病种.Filter = 0
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    If bln特殊病 Then mrs病种.Filter = 0
End Sub

Private Sub txt密码_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim blnTrans As Boolean
    Dim lng业务类型 As Long
    Dim strTemp As String, strData As String
    Dim str个人编号 As String, str参保人单位 As String
    Dim rsTemp As New ADODB.Recordset
    Const str个人帐户 As String = "003"
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Not (mbytType = 0 Or mbytType = 1 Or mbytType = 3) Then Exit Sub
    
    On Error GoTo errHand
    '清除所有内容
    Call ClearCons
    lng业务类型 = cbo业务类型.ItemData(cbo业务类型.ListIndex)
    
    '--读IC卡
    '读出IC卡中的信息
    If Not 调用接口_准备_沈阳市(Function_沈阳市.其他_读卡) Then Exit Sub
    If Not 调用接口_执行_沈阳市 Then Exit Sub
    
    '取返回的记录集
    'If Not 调用接口_指定记录集_沈阳市("ICInfo") Then Exit Sub
    'Modified By 朱玉宝 地区：长沙 原因：将传入参数由卡号改为身份证号
    'If Not 调用接口_读取数据_沈阳市("indi_id", str个人编号) Then Exit Sub
    If Not 调用接口_读取数据_沈阳市("card_no", str个人编号) Then Exit Sub
    
    '--检查该卡是否在黑名单中（调用读卡接口时，接口商已处理此内容）
'    If Not 调用接口_准备_沈阳市(Function_沈阳市.其他_黑名单校验) Then Exit Sub
'    If Not 调用接口_执行_沈阳市 Then Exit Sub
    
    '--读取病人基本信息、业务信息（str个人编号="12010619671003053"）
    Select Case mbytType
    Case 0 '门诊
'        glngReturn_沈阳市 = CZ_Start(glngInterface_沈阳市, Function_沈阳市.普通门诊_身份验证)
'        1   idcard         身份证号        20  否
'        2   hospital_id    医疗机构编码    20  否
'        3   busi_type      业务类型        2   否  "11"：门诊
'        4   password       密码            6   否
        If Not 调用接口_准备_沈阳市(IIf(lng业务类型 = 业务分类_沈阳市.门诊规定病, Function_沈阳市.门诊规定病_身份验证, Function_沈阳市.普通门诊_身份验证)) Then Exit Sub
    Case 1  '住院
        'glngReturn_沈阳市 = CZ_Start(glngInterface_沈阳市, Function_沈阳市.普通住院_身份验证)
'        1   iccardno       磁卡卡号        20  否
'        2   hospital_id    医疗机构编码    20  否
'        3   busi_type      业务类型        2   否  "12"：住院
'        4   reg_flag       登记标志        1   否  "0"：普通住院登记
'        5   password       密码            6   否
        If Not 调用接口_准备_沈阳市(Function_沈阳市.普通住院_身份验证) Then Exit Sub
    Case 3 '挂号
        If Not 调用接口_准备_沈阳市(Function_沈阳市.普通门诊_身份验证) Then Exit Sub
    End Select
    
    '填写入口参数
    'Modified By 朱玉宝 地区：长沙 原因：将传入参数由卡号改为身份证号
    Call CZ_DataPut(glngInterface_沈阳市, 1, "iccardno", str个人编号)
    Call CZ_DataPut(glngInterface_沈阳市, 1, "hospital_id", gCominfo_沈阳市.医院编码)
    Call CZ_DataPut(glngInterface_沈阳市, 1, "busi_type", lng业务类型)
    If mbytType = 1 Then Call CZ_DataPut(glngInterface_沈阳市, 1, "reg_flag", "0")
    Call CZ_DataPut(glngInterface_沈阳市, 1, "password", txt密码.Text)
    '调用接口
    If Not 调用接口_执行_沈阳市 Then Exit Sub
    '读取返回记录集（根据TextBox的Tag读取基本信息1-22）
'    1   indi_id        个人编号            8
'    2   insr_code      保险号              30
'    3   name           姓名                10
'    4   sex            性别                1   "0"：女    "1"：男
'    5   idcard         身份证号码          20
'    6   pers_type      人员类别编码        2
'    7   pers_name      人员类别名称        20
'    8   folk_code      民族编码            2
'    9   folk_name      民族名称            20
'    10  official_code  公务员级别          2
'    11  official_name  公务员级别名称      20
'    12  special_code   特殊照顾人群编码    3
'    13  special_name   特殊照顾人群名称    20
'    14  position_name  职务                10
'    15  work_type      特殊工种            3
'    16  work_type_name 特殊工种名称        20
'    17  city_code      异地安置城市编码    30
'    18  city_name      异地安置城市名称    30
'    19  corp_id        参保人单位编码      20
'    20  corp_name      参保人单位名称      50
'    22  persfundcon    已冻结基金信息      1024
'--------------------以上是任何接口都要返回的基本信息--------------------
'（住院）      21  sum_year       年内累计住院次数    2
'（门诊）      21  last_balance   个人帐户余额        18  单位：元
'（规定病）    22  serial_apply   申请序号            12
'（规定病）    23  icd            疾病编码            20  申请的门诊规定病的病种编码
'（规定病）    24  disease        疾病名称            60  申请的门诊规定病的疾病名称
    If Not 调用接口_指定记录集_沈阳市("PersonInfo") Then Exit Sub
    
    '根据TextBox的Tag读取基本信息1-22
    Call ReadFromInterface
    '如果是沈阳地区，并且是住院病人，根据特殊人群的值判断
    If mint适用地区_沈阳 = 2 And mbytType = 1 Then
        '沈阳地区
        Call 调用接口_读取数据_沈阳市("special_code", strTemp)
        Select Case strTemp
        Case "1"
            MsgBox "该病人年度内因恶性肿瘤住过院，选择病种时请注意选择！", vbInformation, gstrSysName
        Case "0"
            MsgBox "该病人年度内没有恶性肿瘤住院记录，选择病种时请注意选择！", vbInformation, gstrSysName
        End Select
    End If
    
    Call 调用接口_读取数据_沈阳市("sex", strData)
    Me.cbo性别.ListIndex = Val(strData)
    
    '根据不同的业务类型，读取字段值
    Call 调用接口_读取数据_沈阳市("corp_id", str参保人单位)
    Select Case cbo业务类型.ItemData(cbo业务类型.ListIndex)
    Case 业务分类_沈阳市.普通门诊, 业务分类_沈阳市.门诊急救, _
         业务分类_沈阳市.特治特检, 业务分类_沈阳市.生育门诊, 业务分类_沈阳市.工伤门诊
        Call 调用接口_读取数据_沈阳市("last_balance", strData)
        txt帐户余额 = Format(strData, "#####0.00;-#####0.00; ;")
    Case 业务分类_沈阳市.门诊规定病
        '取返回的疾病记录集
        On Error Resume Next
        Dim lngID As Long
        Dim str编码 As String, str名称 As String
        Dim strColumns As String, strValues As String
        '显示申请序号始终是显示的最后一次申请的信息
        Call 调用接口_指定记录集_沈阳市("PersonInfo")
        Call 调用接口_读取数据_沈阳市("serial_apply", strData)
        txt申请序号 = strData
        
        glngReturn_沈阳市 = CZ_SetRecordset(glngInterface_沈阳市, "spinfo")
        mbln多病种 = (glngReturn_沈阳市 > 0)        '设定记录集时，如果成功返回记录数，否则返回-1
        
        '将疾病记录集保存在内存中
        On Error GoTo errHand
        strColumns = "ID|编码|名称"
        blnTrans = True
        
        gcnOracle.BeginTrans
        If mbln多病种 Then
            Call DebugTool("多病种")
            Do While True
                strValues = ""
                Call 调用接口_读取数据_沈阳市("icd", str编码)
                strValues = strValues & "|" & str编码
                Call 调用接口_读取数据_沈阳市("disease", str名称)
                strValues = strValues & "|" & str名称
                '判断是否存在该病种
                gstrSQL = " Select ID From 保险病种" & _
                          " Where 险类=[1] And 编码=[2]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "是否存在此病种", TYPE_沈阳市, str编码)
                If rsTemp.RecordCount = 0 Then
                    lngID = zlDatabase.GetNextID("保险病种")
                    gstrSQL = "zl_保险病种_INSERT(" & lngID & "," & TYPE_沈阳市 & ",'" & str编码 & _
                                "','" & str名称 & "',NULL,0,NULL,NULL)"
                Else
                    lngID = rsTemp!ID
                End If
                
                strValues = lngID & strValues
                Call Record_Add(mrs病种, strColumns, strValues)
                
                Call DebugTool("已加入一行记录")
                If Not 调用接口_移动记录集_沈阳市(MoveNext) Then Exit Do
            Loop
        Else
            '不是多病种的话，肯定只有一条
            Call DebugTool("单病种")
            strValues = ""
            Call 调用接口_指定记录集_沈阳市("PersonInfo")
            Call 调用接口_读取数据_沈阳市("icd", str编码)
            strValues = strValues & "|" & str编码
            Call 调用接口_读取数据_沈阳市("disease", str名称)
            strValues = strValues & "|" & str名称
            
            '判断是否存在该病种
            gstrSQL = " Select ID From 保险病种" & _
                      " Where 险类=[1] And 编码=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "是否存在此病种", TYPE_沈阳市, str编码)
            If rsTemp.RecordCount = 0 Then
                lngID = zlDatabase.GetNextID("保险病种")
                gstrSQL = "zl_保险病种_INSERT(" & lngID & "," & TYPE_沈阳市 & ",'" & str编码 & _
                            "','" & str名称 & "',NULL,0,NULL,NULL)"
            Else
                lngID = rsTemp!ID
            End If
            
            strValues = lngID & strValues
            Call Record_Add(mrs病种, strColumns, strValues)
            Call DebugTool("已加入一行记录")
        End If
        gcnOracle.CommitTrans
        
        '如果只有一条病种信息，直接显示出来
        blnTrans = False
        If mrs病种.RecordCount = 1 Then
            txt疾病信息.Tag = mrs病种!编码
            txt疾病信息.Text = "(" & mrs病种!编码 & ")" & mrs病种!名称
            lbl疾病信息.Tag = txt疾病信息.Text '用于恢复显示
        ElseIf mrs病种.RecordCount > 1 Then
            If frmListSel.ShowSelect(TYPE_沈阳市, mrs病种, "ID", "疾病选择", "请选择医保病种：") = True Then
                txt疾病信息.Tag = mrs病种!编码
                txt疾病信息.Text = "(" & mrs病种!编码 & ")" & mrs病种!名称
                lbl疾病信息.Tag = txt疾病信息.Text '用于恢复显示
                
                mrs病种.Filter = 0
            End If
        End If
    Case Else   '所有住院
        Call 调用接口_读取数据_沈阳市("sum_year", strData)
        txt住院次数 = strData
    End Select
    
    '--读取个人帐户余额（由于住院没有返回，强制取一次）
    If Not 调用接口_准备_沈阳市(Function_沈阳市.其他_基金余额) Then Exit Sub
    '写入口参数
'    1   fund_id    基金编号    3   否
'    2   indi_id    个人编号    8   否
'    3   corp_ID    单位编号    3
    'Modified By 朱玉宝 地区：长沙 原因：需要多传一个参数（corp_id）
    gstrField_沈阳市 = "fund_id||indi_id||corp_id"
    gstrValue_沈阳市 = str个人帐户 & "||" & txt个人编号 & "||" & str参保人单位
    Call 调用接口_写入口参数_沈阳市(1)
    If Not 调用接口_执行_沈阳市 Then Exit Sub
    If Not 调用接口_指定记录集_沈阳市("PersonAccount") Then Exit Sub
    Call 调用接口_读取数据_沈阳市("last_balance", strData)
    txt帐户余额 = Format(strData, "#####0.00;-#####0.00; ;")
    gCominfo_沈阳市.帐户余额 = Val(strData)
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
    If blnTrans Then gcnOracle.RollbackTrans
End Sub

Public Function NeedName(strList As String) As String
    NeedName = Mid(strList, InStr(strList, "-") + 1)
End Function

Public Sub CheckInputLen(txt As Object, KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: Exit Sub
    If KeyAscii < 32 And KeyAscii >= 0 Then Exit Sub
    If txt.MaxLength = 0 Then Exit Sub
    If zlCommFun.ActualLen(txt.Text & Chr(KeyAscii)) > txt.MaxLength Then KeyAscii = 0
End Sub

Private Sub ReadFromInterface()
    Dim objTextBox As TextBox, objCons As Object
    Dim arrData, strData As String, strTemp As String
    
    For Each objCons In Controls
        If UCase(TypeName(objCons)) = "TEXTBOX" Then
            Set objTextBox = objCons
            If Trim(objTextBox.Tag) <> "" And objTextBox.Name <> "txt疾病信息" Then
                arrData = Split(objTextBox.Tag, "|")
                If UBound(arrData) = 0 Then
                    Call 调用接口_读取数据_沈阳市(arrData(0), strData)
                Else
                    Call 调用接口_读取数据_沈阳市(arrData(0), strData)
                    Call 调用接口_读取数据_沈阳市(arrData(1), strTemp)
                    If strData <> "" Then strData = "[" & strData & "]" & strTemp
                End If
                objTextBox.Text = strData
            End If
        End If
    Next
End Sub

Private Sub Record_Add(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String)
    Dim arrFields, arrValues, intField As Integer
    '添加记录
    'strFields:字段名|字段名
    'strValues:值|值
    
    '例子：
    'Dim strFields As String, strValues As String
    'strFields = "RecordID|科目ID|摘要"
    'strValues = "5188|6666|科目名称"
    'Call Record_Update(rsVoucher, strFields, strValues)

    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField = 0 Then Exit Sub

    With rsObj
        .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Private Sub Record_Init(ByRef rsObj As ADODB.Recordset, ByVal strFields As String)
    Dim arrFields, intField As Integer
    Dim strFieldName As String, intType As Integer, lngLength As Long
    '初始化映射记录集
    'strFields:字段名,类型,长度|字段名,类型,长度    如果长度为零,则取默认长度
    '字符型:adLongVarChar;数字型:adDouble;日期型:adDBDate
    
    '例子：
    'Dim rsVoucher As New ADODB.Recordset, strFields As String
    'strFields = "RecordID," & adDouble & ",18|科目ID," & adDouble & ",18|摘要, " & adLongVarChar & ",50|" & _
    '"删除," & adDouble & ",1"
    'Call Record_Init(rsVoucher, strFields)

    arrFields = Split(strFields, "|")
    Set rsObj = New ADODB.Recordset

    With rsObj
        If .State = 1 Then .Close
        For intField = 0 To UBound(arrFields)
            strFieldName = Split(arrFields(intField), ",")(0)
            intType = Split(arrFields(intField), ",")(1)
            lngLength = Split(arrFields(intField), ",")(2)

            '获取字段缺省长度
            If lngLength = 0 Then
                Select Case intType
                Case adDouble
                    lngLength = madDoubleDefault
                Case adVarChar
                    lngLength = madLongVarCharDefault
                Case adLongVarChar
                    lngLength = madLongVarCharDefault
                Case Else
                    lngLength = madDbDateDefault
                End Select
            End If
            .Fields.Append strFieldName, intType, lngLength, adFldIsNullable
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub
