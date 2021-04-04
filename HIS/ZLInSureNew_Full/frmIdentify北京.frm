VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BILLEDIT.OCX"
Begin VB.Form frmIdentify北京 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "身份验证"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11985
   Icon            =   "frmIdentify北京.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   11985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmd刷新 
      Caption         =   "刷新(&R)"
      Height          =   350
      Left            =   6120
      TabIndex        =   46
      Top             =   6480
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   10575
      TabIndex        =   45
      Top             =   6480
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   9330
      TabIndex        =   44
      Top             =   6480
      Width           =   1100
   End
   Begin VB.CommandButton cmd功能 
      Caption         =   "删除历史记录(&D)"
      Enabled         =   0   'False
      Height          =   350
      Index           =   2
      Left            =   3840
      TabIndex        =   49
      ToolTipText     =   "快捷键：DEL"
      Top             =   6480
      Width           =   1725
   End
   Begin VB.CommandButton cmd功能 
      Caption         =   "插入历史记录(&I)"
      Enabled         =   0   'False
      Height          =   350
      Index           =   1
      Left            =   1980
      TabIndex        =   48
      ToolTipText     =   "快捷键：Ctrl+I"
      Top             =   6480
      Width           =   1725
   End
   Begin VB.CommandButton cmd功能 
      Caption         =   "增加历史记录(&A)"
      Enabled         =   0   'False
      Height          =   350
      Index           =   0
      Left            =   120
      TabIndex        =   47
      ToolTipText     =   "快捷键：Ctrl+A"
      Top             =   6480
      Width           =   1725
   End
   Begin TabDlg.SSTab tabShow 
      Height          =   3735
      Left            =   120
      TabIndex        =   41
      Top             =   2640
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   6588
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      Enabled         =   0   'False
      TabCaption(0)   =   "住院记录(&1)"
      TabPicture(0)   =   "frmIdentify北京.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Bill(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "门诊记录(&2)"
      TabPicture(1)   =   "frmIdentify北京.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Bill(1)"
      Tab(1).ControlCount=   1
      Begin ZL9BillEdit.BillEdit Bill 
         Height          =   3255
         Index           =   0
         Left            =   90
         TabIndex        =   42
         Top             =   390
         Width           =   11625
         _ExtentX        =   20505
         _ExtentY        =   5741
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483643
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483643
      End
      Begin ZL9BillEdit.BillEdit Bill 
         Height          =   3255
         Index           =   1
         Left            =   -74910
         TabIndex        =   43
         Top             =   390
         Width           =   11625
         _ExtentX        =   20505
         _ExtentY        =   5741
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483643
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483643
      End
   End
   Begin VB.Frame fra基本信息 
      Caption         =   "基本信息(&X)"
      Enabled         =   0   'False
      Height          =   1905
      Left            =   90
      TabIndex        =   6
      Top             =   600
      Width           =   11835
      Begin VB.ComboBox cbo隶属关系 
         Height          =   300
         Left            =   7320
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   690
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.ComboBox cbo家床方式 
         Enabled         =   0   'False
         Height          =   300
         Left            =   10320
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   1470
         Width           =   1365
      End
      Begin VB.ComboBox cbo入院类型 
         Enabled         =   0   'False
         Height          =   300
         Left            =   3510
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   1470
         Width           =   1335
      End
      Begin VB.ComboBox cbo入院方式 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1170
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   1470
         Width           =   1335
      End
      Begin VB.ComboBox cbo家床类型 
         Enabled         =   0   'False
         Height          =   300
         Left            =   7860
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   1470
         Width           =   1395
      End
      Begin MSMask.MaskEdBox txt截止日期 
         Height          =   300
         Left            =   5760
         TabIndex        =   30
         Top             =   1080
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Format          =   "yyyy-MM-dd"
         Mask            =   "####-##-##"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cbo特殊病种 
         Height          =   300
         Left            =   1170
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   1080
         Width           =   2745
      End
      Begin VB.ComboBox cbo公务员待遇 
         Enabled         =   0   'False
         Height          =   300
         Left            =   10080
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   690
         Width           =   1575
      End
      Begin VB.ComboBox cbo公务员 
         Height          =   300
         Left            =   7320
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   690
         Width           =   1395
      End
      Begin VB.ComboBox cbo报销区县 
         Height          =   300
         Left            =   4500
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   690
         Width           =   1725
      End
      Begin VB.ComboBox cbo参保类别 
         Height          =   300
         Left            =   1170
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   690
         Width           =   1965
      End
      Begin VB.TextBox txt社保证号 
         Height          =   300
         Left            =   10080
         MaxLength       =   16
         TabIndex        =   16
         Top             =   300
         Width           =   1575
      End
      Begin VB.TextBox txt身份证号 
         Height          =   300
         Left            =   7320
         MaxLength       =   18
         TabIndex        =   14
         Top             =   300
         Width           =   1725
      End
      Begin VB.ComboBox cbo性别 
         Height          =   300
         Left            =   5250
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   300
         Width           =   975
      End
      Begin VB.TextBox txt姓名 
         Height          =   300
         Left            =   3240
         MaxLength       =   20
         TabIndex        =   10
         Top             =   300
         Width           =   1275
      End
      Begin VB.ComboBox cbo医疗类别 
         Height          =   300
         Left            =   1170
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   300
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txt入院日期 
         Height          =   300
         Left            =   5760
         TabIndex        =   36
         Top             =   1470
         Visible         =   0   'False
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Format          =   "yyyy-MM-dd"
         Mask            =   "####-##-##"
         PromptChar      =   "_"
      End
      Begin VB.Label lbl隶属关系 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "隶属关系*"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6390
         TabIndex        =   25
         Top             =   750
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label lbl家床方式 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "家床方式"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   9480
         TabIndex        =   39
         Top             =   1530
         Width           =   720
      End
      Begin VB.Label lbl入院日期 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "入院日期"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4980
         TabIndex        =   35
         Top             =   1530
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lbl入院类型 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "入院类型"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2670
         TabIndex        =   33
         Top             =   1530
         Width           =   720
      End
      Begin VB.Label lbl入院方式 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "入院方式"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   330
         TabIndex        =   31
         Top             =   1530
         Width           =   720
      End
      Begin VB.Label lbl家床类型 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "家床类型"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   7020
         TabIndex        =   37
         Top             =   1530
         Width           =   720
      End
      Begin VB.Label lbl特殊病截止日期 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "特殊病有效截止日期"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4080
         TabIndex        =   29
         Top             =   1140
         Width           =   1620
      End
      Begin VB.Label lbl特殊病种 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "特殊病种"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   330
         TabIndex        =   27
         Top             =   1140
         Width           =   720
      End
      Begin VB.Label lbl公务员待遇 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "公务员待遇"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   9060
         TabIndex        =   23
         Top             =   750
         Width           =   900
      End
      Begin VB.Label lbl公务员 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "公务员*"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6570
         TabIndex        =   21
         Top             =   750
         Width           =   630
      End
      Begin VB.Label lbl报销区县 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "报销区(县)*"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3390
         TabIndex        =   19
         Top             =   750
         Width           =   990
      End
      Begin VB.Label lbl参保类别 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "参保类别*"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   240
         TabIndex        =   17
         Top             =   750
         Width           =   810
      End
      Begin VB.Label lbl社保证号 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "社保证号*"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   9150
         TabIndex        =   15
         Top             =   360
         Width           =   810
      End
      Begin VB.Label lbl身份证号 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "身份证号*"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6390
         TabIndex        =   13
         Top             =   360
         Width           =   810
      End
      Begin VB.Label lbl性别 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "性别*"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4680
         TabIndex        =   11
         Top             =   360
         Width           =   450
      End
      Begin VB.Label lbl姓名 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "姓名*"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2670
         TabIndex        =   9
         Top             =   360
         Width           =   450
      End
      Begin VB.Label lbl医疗类别 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "医疗类别*"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   810
      End
   End
   Begin VB.CommandButton cmd读卡 
      Caption         =   "读卡(&R)"
      Enabled         =   0   'False
      Height          =   345
      Left            =   3150
      TabIndex        =   2
      Top             =   180
      Width           =   1155
   End
   Begin VB.TextBox txt确认卡号 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8130
      MaxLength       =   12
      TabIndex        =   5
      Top             =   210
      Width           =   1905
   End
   Begin VB.TextBox txt卡号 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6120
      MaxLength       =   12
      TabIndex        =   4
      Top             =   210
      Width           =   1905
   End
   Begin VB.ComboBox cbo病人类型 
      Height          =   300
      Left            =   1470
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   210
      Width           =   1605
   End
   Begin VB.Label lbl卡号 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "卡/手册号(&S)*"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   4860
      TabIndex        =   3
      Top             =   270
      Width           =   1170
   End
   Begin VB.Label lbl病人类型 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "病人类型(&T)*"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   300
      TabIndex        =   0
      Top             =   270
      Width           =   1080
   End
End
Attribute VB_Name = "frmIdentify北京"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Enum 历史记录
    住院 = 0
    门诊
End Enum
Enum 功能
    增加 = 0
    插入
    删除
End Enum

'表格控件常量
Private Const col_医疗机构 As Integer = 0
Private Const col_就诊日期 As Integer = 2
Private Const col_入院日期 As Integer = 2
Private Const col_出院日期 As Integer = 4
'住院历史记录
Private Const col_入院类型 As Integer = 1
Private Const col_出院类型 As Integer = 3
'门诊历史记录
Private Const col_医疗类别 As Integer = 1

Private mbytType As Byte                '模式 0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号
Private mlng病人ID  As Long
Private mstrReturn As String
Private mdbl帐户余额 As Double
Private mbln定点医院 As Boolean
Private mbln特殊病定点医院 As Boolean
'1、定点医院的判断（程序上不对定点医院进行判断，由操作员判断，非定点医院则走普通流程）
'A、 中医（医院类型02）、专科医院（医院类型03）视为所有参保人的定点医院。
'B、 急诊、急诊收住院可在参保人的非定点医院就诊，享受医保待遇。
'C、 参保人持住院单可在参保人的非定点医院就诊，享受医保待遇。
'
'2、手册填写问题:
'A、跨周期跨年度住院费用在手册上应记录分段分解信息


'离休、在乡二等乙级伤残，需要确定隶属关系
Private Sub LoadInitData()
    Dim strData As String
    '装入缺省数据（本来指标体系里已有这些数据，考虑到这些都是基础数据，因此此处写死）
    
    strData = "手册,0|卡,1"
    Call LoadCboData(cbo病人类型, strData)
    strData = "男,1|女,2|未知,9"
    Call LoadCboData(cbo性别, strData)
    strData = ",0|肾透,1|恶性肿瘤放化疗,2|肾透+恶性肿瘤放化疗,3|抗排异,4|肾透+抗排异,5|恶性肿瘤放化疗+抗排异,6|肾透+恶性肿瘤放化疗+抗排异,7"
    Call LoadCboData(cbo特殊病种, strData)
    strData = "在职,11|在职长期驻外,12|在职二等乙级伤残军人,13|退休,21|退休异地安置,22|退职二等乙级伤残军人,23|退休二等乙级伤残军人,24|" & _
            "退职,25|退职异地安置,26|离休,31|老红军,32|特殊全免人员,34|在职司局级医照人员,35|退休司局级医照人员,36|在职副部级医照人员,37|" & _
            "退休副部级医照人员,38|最低保障在职,40|最低保障退休,41|最低保障退职,42|在乡二等乙级伤残军人,49|两院院士,51|优诊待遇人员,52|" & _
            "社会退休人员,61|社会退职人员,63|企业改组退休人员,65|企业改组退休易地安置人员,66|企业改组退职人员,67|企业改组退职易地安置人员,68|" & _
            "三资退休人员,71|三资退职人员,73|支援乡镇退休人员,75|支援乡镇退职人员,77|破产退休人员,81|破产退职人员,83|" & _
            "企业注销吊销退休人员,85|企业注销吊销退休易地安置人员,86|企业注销吊销退职人员,87|企业注销吊销退职易地安置人员,88|其它人员,91"
    Call LoadCboData(cbo参保类别, strData)
    strData = "东城区,1010|西城区,1020|崇文区,1030|宣武区,1040|朝阳区,1050|朝阳区酒仙桥,1051|朝阳区安贞,1052|朝阳区九龙山,1053|" & _
            "丰台区,1060|丰台区大红门,1061|石景山区,1070|海淀区,1080|海淀区万寿路,1081|海淀区上地,1082|门头沟区,1090|房山区,1110|" & _
            "昌平区,2210|顺义区,2220|通州区,2230|大兴县,2240|平谷县,2260|怀柔县,2270|密云县,2280|延庆县,2290|北京市经济技术开发区,2310|北京市医保中心,2320"
    Call LoadCboData(cbo报销区县, strData)
    strData = "不享受,1|享受,0"
    Call LoadCboData(cbo公务员, strData)
    strData = ",-1|中央,110|市级,120|东城区,140|西城区,150|崇文区,160|宣武区,170|朝阳区,180|" & _
            "海淀区,190|丰台区,200|石景山区,210|门头沟区,220|房山区,230|通州区,240|" & _
            "大兴区,250|昌平区,260|顺义区,270|怀柔县,280|密云县,290|平谷县,310|延庆县,320|经济开发区,330"
    Call LoadCboData(cbo公务员待遇, strData)
    strData = ",-1|中央,1|省,2|计划单列市,3|市,4|区,41|县,5|街道,51|乡镇,6|部队,7|其它,9"
    Call LoadCboData(cbo隶属关系, strData)
    
        
    '如果是门诊，只显示门诊特殊病；如果是住院，仅显示住院，否则全部显示
    If mbytType = 0 Then
        strData = "门诊特殊病,12"
    ElseIf mbytType = 1 Then
        strData = "门诊特殊病,12|住院,21"
    Else
        strData = "门诊特殊病,12|住院,21"
    End If
    Call LoadCboData(cbo医疗类别, strData)
    
    strData = "普通,0|精神病,2"
    Call LoadCboData(cbo家床类型, strData)
    strData = "正常,0|新转入,1"
    Call LoadCboData(cbo家床方式, strData)
    strData = "新入院,0|转入院,1"
    Call LoadCboData(cbo入院方式, strData)
    strData = "普通,0|特殊病病人,1|精神病,3|中医医院针灸科,4"   '器官移植,2
    Call LoadCboData(cbo入院类型, strData)
End Sub

Private Sub LoadCboData(ByVal cboObj As ComboBox, ByVal strData As String)
    Dim arrData
    Dim intIndex As Integer, intCOUNT As Integer
    
    arrData = Split(strData, "|")
    intCOUNT = UBound(arrData)
    With cboObj
        .Clear
        For intIndex = 0 To intCOUNT
            .AddItem Split(arrData(intIndex), ",")(0)
            .ItemData(.NewIndex) = Split(arrData(intIndex), ",")(1)
        Next
        .ListIndex = 0
    End With
End Sub

Private Sub InitBill()
    Dim arrCol
    Dim billObj As BillEdit
    Dim intCol As Integer, intCols As Integer
    '格式说明：列名,宽度,列属性
    Const str住院 As String = "医院名称,1500,1|入院类型,1000,3|入院日期,1000,2|出院类型,1000,3|出院日期,1000,2|" & _
        "统筹支付,1000,4|大额/公务员补助,1600,4|个人自付,1000,4|个人自费,1000,4|统筹封顶后医保内,1800,4"
    Const str门诊 As String = "医院名称,1500,1|医疗类别,1200,3|就诊日期,1000,2|统筹支付,1000,4|" & _
        "大额/公务员补助,1600,4|个人自付,1000,4|个人自费,1000,4|统筹封顶后医保内,1700,4"
    
    '对住院表格进行初始化
    arrCol = Split(str住院, "|")
    intCols = UBound(arrCol)
    Set billObj = Bill(住院)
    billObj.ClearBill
    billObj.Active = True
    billObj.Cols = intCols + 1
    For intCol = 0 To intCols
        billObj.TextMatrix(0, intCol) = Split(arrCol(intCol), ",")(0)
        billObj.ColWidth(intCol) = Split(arrCol(intCol), ",")(1)
        billObj.ColData(intCol) = Split(arrCol(intCol), ",")(2)
    Next
    
    '对门诊表格进行初始化
    arrCol = Split(str门诊, "|")
    intCols = UBound(arrCol)
    Set billObj = Bill(门诊)
    billObj.ClearBill
    billObj.Active = True
    billObj.Cols = intCols + 1
    For intCol = 0 To intCols
        billObj.TextMatrix(0, intCol) = Split(arrCol(intCol), ",")(0)
        billObj.ColWidth(intCol) = Split(arrCol(intCol), ",")(1)
        billObj.ColData(intCol) = Split(arrCol(intCol), ",")(2)
    Next
End Sub

Private Sub ReadBill()
    On Error GoTo errHand
    Dim rsHostory As New ADODB.Recordset
    '读取指定病人的历史就诊记录
    If Trim(txt确认卡号.Text) = "" Then Exit Sub
    Call InitBill
    
    Call DebugTool("提取历史就诊记录")
    Call WriteBusinessLOG("提取历史就诊记录", "", "")
    '标志为2表示住院；否则当作门诊
    gstrSQL = "SELECT DECODE(医疗类别,21,2,22,2,23,2,1) AS 标志,医疗机构,卡号,B.名称 AS 医疗类别," & _
             " TO_CHAR(入院日期,'yyyy-MM-dd') AS 入院日期,C.名称 AS 入院类型," & _
             " TO_CHAR(出院日期,'yyyy-MM-dd') AS 出院日期,D.名称 AS 出院类型, " & _
             " 费用总额,统筹支付,大额支付,个人自付,个人自费,统筹封顶后医保内 " & _
             " FROM 手册消费记录 A, " & _
             "      (SELECT B.编码,B.名称" & _
             "       FROM 指标主表 A,指标体系对照表 B" & _
             "       WHERE A.类别=B.类别 And A.名称='医疗类别') B," & _
             "      (SELECT B.编码,B.名称" & _
             "       FROM 指标主表 A,指标体系对照表 B" & _
             "       WHERE A.类别=B.类别 And A.名称='入院方式') C," & _
             "      (SELECT B.编码,B.名称" & _
             "       FROM 指标主表 A,指标体系对照表 B" & _
             "       WHERE A.类别=B.类别 And A.名称='出院类别') D" & _
             " WHERE A.医疗类别=B.编码 AND A.入院类型=C.编码(+) And A.出院类型=D.编码(+)" & _
             " AND A.卡号='" & txt确认卡号.Text & "'" & _
             " ORDER BY 标志,入院日期"
    If rsHostory.State = 1 Then rsHostory.Close
    Call SQLTest(App.Title, "ZL9INSURE\READBILL", gstrSQL): rsHostory.Open gstrSQL, gcnBJYB: Call SQLTest
    Call DebugTool("历史就诊记录条数：" & rsHostory.RecordCount)
    Call WriteBusinessLOG("历史就诊记录条数：" & rsHostory.RecordCount, "", "")
    
    Call WriteBill(rsHostory)
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub ReadPatient()
    On Error GoTo errHand
    Dim RSPATIENT As New ADODB.Recordset
    If Trim(txt确认卡号.Text) = "" Then Exit Sub
    
    Call DebugTool("提取该病人的基本信息--以前在本院就诊过的就有信息")
    Call WriteBusinessLOG("提取该病人的基本信息--以前在本院就诊过的就有信息", "", "")
    gstrSQL = "SELECT Nvl(病人ID,0) AS 病人ID,卡号,社保证号,姓名,B.名称 AS 性别,血型,身份证号, " & _
             "     C.名称 AS 参保类别,D.名称 AS 缴费地区代码,公务员,公务员待遇,病种标识, " & _
             "     TO_CHAR(特殊病截止日期,'yyyy-MM-dd') AS 特殊病截止日期 " & _
             " FROM 保险帐户 A, " & _
             "     (SELECT B.编码,B.名称 " & _
             "      FROM 指标主表 A,指标体系对照表 B " & _
             "      WHERE A.名称='性别' AND A.类别=B.类别) B, " & _
             "     (SELECT B.编码,B.名称 " & _
             "      FROM 指标主表 A,指标体系对照表 B " & _
             "      WHERE A.名称='医保参保人员类别' AND A.类别=B.类别) C, " & _
             "     (SELECT B.编码,B.名称 " & _
             "      FROM 指标主表 A,指标体系对照表 B " & _
             "      WHERE A.名称='报销区县' AND A.类别=B.类别) D" & _
             " WHERE 卡号='" & Trim(txt确认卡号.Text) & "'" & _
             " AND A.性别=B.编码(+) AND A.参保类别=C.编码(+) AND A.缴费地区代码=D.编码(+)"
    If RSPATIENT.State = 1 Then RSPATIENT.Close
    Call SQLTest(App.Title, "ZL9INSURE\READBILL", gstrSQL): RSPATIENT.Open gstrSQL, gcnBJYB: Call SQLTest
    Call DebugTool("成功提取该病人以往就诊时登记的基本信息")
    Call WriteBusinessLOG("成功提取该病人以往就诊时登记的基本信息", "", "")
    
    Call WritePatient(RSPATIENT)
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub WritePatient(ByVal RSPATIENT As ADODB.Recordset)
    On Error GoTo errHand
    '将病人的基本信息写入界面
    
    With RSPATIENT
        If .RecordCount = 0 Then Exit Sub
        txt姓名.Text = Nvl(!姓名)
        Call zlControl.CboLocate(cbo性别, !性别)
        txt身份证号.Text = Nvl(!身份证号)
        txt社保证号.Text = Nvl(!社保证号)
        Call zlControl.CboLocate(cbo参保类别, !参保类别)
        Call zlControl.CboLocate(cbo报销区县, !缴费地区代码)
        Call zlControl.CboLocate(cbo隶属关系, !公务员待遇, True)
        Call zlControl.CboLocate(cbo公务员, !公务员, True)
        Call zlControl.CboLocate(cbo公务员待遇, !公务员待遇, True)
        Call zlControl.CboLocate(cbo特殊病种, !病种标识, True)
        txt截止日期.Text = Nvl(!特殊病截止日期, "____-__-__")
    End With
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub WriteBill(ByVal rsHostory As ADODB.Recordset)
    Dim lngRow As Long
    Dim objBill As BillEdit
    On Error GoTo errHand
    '将历史就诊记录填写到表格中
    Bill(住院).Redraw = False
    Bill(门诊).Redraw = False
    
    With rsHostory
        '住院
        .Filter = "标志=2"
        Set objBill = Bill(住院)
        Do While Not .EOF
            lngRow = .AbsolutePosition
            objBill.TextMatrix(lngRow, 0) = Nvl(rsHostory!医疗机构)
            objBill.TextMatrix(lngRow, 1) = Nvl(rsHostory!入院类型)
            objBill.TextMatrix(lngRow, 2) = Nvl(rsHostory!入院日期)
            objBill.TextMatrix(lngRow, 3) = Nvl(rsHostory!出院类型)
            objBill.TextMatrix(lngRow, 4) = Nvl(rsHostory!出院日期)
            'objBill.TextMatrix(lngRow, 5) = Format(Nvl(rsHostory!费用总额, 0), "#####0.00;-#####0.00; ;")
            objBill.TextMatrix(lngRow, 5) = Format(Nvl(rsHostory!统筹支付, 0), "#####0.00;-#####0.00; ;")
            objBill.TextMatrix(lngRow, 6) = Format(Nvl(rsHostory!大额支付, 0), "#####0.00;-#####0.00; ;")
            objBill.TextMatrix(lngRow, 7) = Format(Nvl(rsHostory!个人自付, 0), "#####0.00;-#####0.00; ;")
            objBill.TextMatrix(lngRow, 8) = Format(Nvl(rsHostory!个人自费, 0), "#####0.00;-#####0.00; ;")
            objBill.TextMatrix(lngRow, 9) = Format(Nvl(rsHostory!统筹封顶后医保内, 0), "#####0.00;-#####0.00; ;")
            
            lngRow = lngRow + 1
            objBill.Rows = objBill.Rows + 1
            .MoveNext
        Loop
        '门诊
        .Filter = "标志=1"
        Set objBill = Bill(门诊)
        Do While Not .EOF
            lngRow = .AbsolutePosition
            objBill.TextMatrix(lngRow, 0) = Nvl(rsHostory!医疗机构)
            objBill.TextMatrix(lngRow, 1) = Nvl(rsHostory!医疗类别)
            objBill.TextMatrix(lngRow, 2) = Nvl(rsHostory!入院日期)
            'objBill.TextMatrix(lngRow, 3) = Format(Nvl(rsHostory!费用总额, 0), "#####0.00;-#####0.00; ;")
            objBill.TextMatrix(lngRow, 3) = Format(Nvl(rsHostory!统筹支付, 0), "#####0.00;-#####0.00; ;")
            objBill.TextMatrix(lngRow, 4) = Format(Nvl(rsHostory!大额支付, 0), "#####0.00;-#####0.00; ;")
            objBill.TextMatrix(lngRow, 5) = Format(Nvl(rsHostory!个人自付, 0), "#####0.00;-#####0.00; ;")
            objBill.TextMatrix(lngRow, 6) = Format(Nvl(rsHostory!个人自费, 0), "#####0.00;-#####0.00; ;")
            objBill.TextMatrix(lngRow, 7) = Format(Nvl(rsHostory!统筹封顶后医保内, 0), "#####0.00;-#####0.00; ;")
            
            lngRow = lngRow + 1
            objBill.Rows = objBill.Rows + 1
            .MoveNext
        Loop
    End With
errHand:
    rsHostory.Filter = 0
    Bill(住院).Redraw = True
    Bill(门诊).Redraw = True
End Sub

Private Sub Bill_cboClick(Index As Integer, ListIndex As Long)
    With Bill(Index)
'        If .LastRow <> .Row Then Exit Sub
'        .TextMatrix(.Row, .Col) = .CboText
    End With
End Sub

Private Sub Bill_CommandClick(Index As Integer)
    Dim blnReturn As Boolean
    Dim rsTmp As New ADODB.Recordset
    With Bill(Index)
        If Bill(Index).COL = col_医疗机构 Then
            gstrSQL = "" & _
                " SELECT A.医院编码,A.医院名称,zlSpellcode(A.医院名称) As 简码,B.编码||'-'||B.名称 AS 医院等级,C.编码||'-'||C.名称 AS 医院类型" & _
                " FROM 医院等级 A," & _
                "     (SELECT B.编码,B.名称" & _
                "     FROM 指标主表 A,指标体系对照表 B" & _
                "     WHERE A.类别=B.类别 AND A.名称='医院等级') B," & _
                "     (SELECT B.编码,B.名称" & _
                "     FROM 指标主表 A,指标体系对照表 B" & _
                "     WHERE A.类别=B.类别 AND A.名称='医院类型') C" & _
                " WHERE A.医院等级=B.编码(+) AND A.医院类型=C.编码(+) AND A.生效日期<=SYSDATE"
            If rsTmp.State = 1 Then rsTmp.Close
            Call SQLTest(App.Title, "ZL9INSURE\保险参数设置", gstrSQL): rsTmp.Open gstrSQL, gcnBJYB: Call SQLTest
            If rsTmp.RecordCount = 0 Then
                MsgBox "没有找到该医院信息，请重输！", vbInformation, gstrSysName
                Exit Sub
            Else
                '出现选择器
                If rsTmp.RecordCount > 1 Then
                    '对于字段大于3的，即使只有一条记录把该对话框显示出来，以便让用户得到更多的信息
                    blnReturn = frmListSel.ShowSelect(TYPE_北京, rsTmp, "医院编码", "医院等级选择", "请选择医院等级：")
                Else
                    blnReturn = True
                End If
            End If
            If blnReturn Then
                .Text = rsTmp!医院名称
                .TextMatrix(.Row, .COL) = .Text
            End If
        End If
    End With
End Sub

Private Sub Bill_EnterCell(Index As Integer, Row As Long, COL As Long)
    With Bill(Index)
        If COL = 1 Then     'col_入院类型 ,col_医疗类别
            .Clear
            If Index = 住院 Then
                .AddItem "普通"
                .ItemData(.NewIndex) = 0
                .AddItem "特殊病病人"
                .ItemData(.NewIndex) = 1
                .AddItem "精神病"
                .ItemData(.NewIndex) = 3
                .AddItem "中医医院针灸科"
                .ItemData(.NewIndex) = 4
                .ListIndex = 0
            Else
                .AddItem "门诊特殊病"
                .ItemData(.NewIndex) = 12
                .AddItem "家庭病床"
                .ItemData(.NewIndex) = 31
                .ListIndex = 0
            End If
        ElseIf COL = col_出院类型 And Index = 住院 Then
            .Clear
            '0-出院,1-转出院，2-中途结算
            .AddItem "出院"
            .ItemData(.NewIndex) = 0
            .AddItem "转出院"
            .ItemData(.NewIndex) = 1
            .AddItem "中途结算"
            .ItemData(.NewIndex) = 2
            .ListIndex = 0
        ElseIf COL = col_医疗机构 Then
            .TxtCheck = False
        Else
            If .ColData(.COL) = 4 Then
                '默认为是金额输入列
                .TxtCheck = True
                .TextMask = "-0123456789."
            End If
        End If
    End With
End Sub

Private Sub Bill_GotFocus(Index As Integer)
    tabShow.Tab = Index
End Sub

Private Sub Bill_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim StrInput As String
    Dim blnReturn As Boolean
    Dim rsTmp As New ADODB.Recordset
    
    With Bill(Index)
        If KeyCode <> vbKeyReturn Then Exit Sub
        If .TxtVisible = False Then
            StrInput = IIf(.TextMatrix(.Row, .COL) = "", " ", .TextMatrix(.Row, .COL))
            .Text = StrInput
            .TextMatrix(.Row, .COL) = StrInput
        Else
            StrInput = UCase(Trim(.Text))
            If .COL = col_医疗机构 Then
                If Trim(StrInput) = "" Then Exit Sub
                gstrSQL = "SELECT * FROM (" & _
                    " SELECT A.医院编码,A.医院名称,zlSpellcode(A.医院名称) As 简码,B.编码||'-'||B.名称 AS 医院等级,C.编码||'-'||C.名称 AS 医院类型" & _
                    " FROM 医院等级 A," & _
                    "     (SELECT B.编码,B.名称" & _
                    "     FROM 指标主表 A,指标体系对照表 B" & _
                    "     WHERE A.类别=B.类别 AND A.名称='医院等级') B," & _
                    "     (SELECT B.编码,B.名称" & _
                    "     FROM 指标主表 A,指标体系对照表 B" & _
                    "     WHERE A.类别=B.类别 AND A.名称='医院类型') C" & _
                    " WHERE A.医院等级=B.编码(+) AND A.医院类型=C.编码(+) AND A.生效日期<=SYSDATE) A" & _
                    " WHERE (A.医院编码 Like '" & StrInput & "%' Or A.医院名称 Like '" & StrInput & "%' Or A.简码 Like '" & StrInput & "%')"
                If rsTmp.State = 1 Then rsTmp.Close
                Call SQLTest(App.Title, "ZL9INSURE\保险参数设置", gstrSQL): rsTmp.Open gstrSQL, gcnBJYB: Call SQLTest
                If rsTmp.RecordCount = 0 Then
                    MsgBox "没有找到该医院信息，请重输！", vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                Else
                    '出现选择器
                    If rsTmp.RecordCount > 1 Then
                        '对于字段大于3的，即使只有一条记录把该对话框显示出来，以便让用户得到更多的信息
                        blnReturn = frmListSel.ShowSelect(TYPE_北京, rsTmp, "医院编码", "医院等级选择", "请选择医院等级：")
                    Else
                        blnReturn = True
                    End If
                End If
                If blnReturn Then
                    .Text = rsTmp!医院名称
                    .TextMatrix(.Row, .COL) = .Text
                End If
            ElseIf .COL = col_入院日期 Or (.COL = col_出院日期 And Index = 住院) Then
                If Trim(StrInput) <> "" Then
                    If Not IsDate(StrInput) Then
                        MsgBox "不是有效的日期数据，请重输！", vbInformation, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                End If
                
                If Index = 住院 Then
                    If .COL = col_入院日期 Then
                        '不能大于出院日期
                        If Trim(.TextMatrix(.Row, col_出院日期)) <> "" Then
                            If StrInput > .TextMatrix(.Row, col_出院日期) Then
                                MsgBox "入院日期不能大于出院日期！", vbInformation, gstrSysName
                                Cancel = True
                                .TxtSetFocus
                                Exit Sub
                            End If
                        End If
                        If .Row > 1 Then
                            '不能小于上条记录的出院日期
                            If Trim(.TextMatrix(.Row - 1, col_出院日期)) <> "" Then
                                If StrInput < .TextMatrix(.Row - 1, col_出院日期) Then
                                    MsgBox "入院日期不能小于上次就诊的出院日期！", vbInformation, gstrSysName
                                    Cancel = True
                                    .TxtSetFocus
                                    Exit Sub
                                End If
                            End If
                        End If
                        If .Row + 1 <= .Rows - 1 Then
                            '不能大于下一条记录的入院日期
                            If Trim(.TextMatrix(.Row + 1, col_入院日期)) <> "" Then
                                If StrInput > .TextMatrix(.Row + 1, col_入院日期) Then
                                    MsgBox "入院日期不能大于" & .Row + 1 & "行登记的入院日期！", vbInformation, gstrSysName
                                    Cancel = True
                                    .TxtSetFocus
                                    Exit Sub
                                End If
                            End If
                        End If
                    ElseIf .COL = col_出院日期 Then
                        '出院日期不能小于入院日期
                        If Trim(.TextMatrix(.Row, col_入院日期)) <> "" Then
                            If StrInput < .TextMatrix(.Row, col_入院日期) Then
                                MsgBox "出院日期不能小于入院日期！", vbInformation, gstrSysName
                                Cancel = True
                                .TxtSetFocus
                                Exit Sub
                            End If
                        End If
                        If .Row + 1 <= .Rows - 1 Then
                            '出院日期不能大于下一条记录的入院日期
                            If Trim(.TextMatrix(.Row + 1, col_入院日期)) <> "" Then
                                If StrInput > .TextMatrix(.Row + 1, col_入院日期) Then
                                    MsgBox "出院日期不能大于" & .Row + 1 & "行登记的入院日期！", vbInformation, gstrSysName
                                    Cancel = True
                                    .TxtSetFocus
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                Else
                    If .COL = col_入院日期 Then
                        If .Row > 1 Then
                            If Trim(.TextMatrix(.Row - 1, col_入院日期)) <> "" Then
                                If StrInput < .TextMatrix(.Row - 1, col_入院日期) Then
                                    MsgBox "就诊日期不能小于上次就诊日期！", vbInformation, gstrSysName
                                    Cancel = True
                                    .TxtSetFocus
                                    Exit Sub
                                End If
                            End If
                        End If
                        If .Row + 1 <= .Rows - 1 Then
                            '就诊日期不能大于下一条记录的就诊日期
                            If Trim(.TextMatrix(.Row + 1, col_入院日期)) <> "" Then
                                If StrInput > .TextMatrix(.Row + 1, col_入院日期) Then
                                    MsgBox "就诊日期不能大于" & .Row + 1 & "行登记的就诊日期！", vbInformation, gstrSysName
                                    Cancel = True
                                    .TxtSetFocus
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                End If
            Else    '都是金额列，只能输入金额
                .Text = Format(StrInput, "#0.00")
            End If
        End If
    End With
End Sub

Private Sub cbo病人类型_Click()
    Dim blnEnable As Boolean
    blnEnable = (cbo病人类型.ItemData(cbo病人类型.ListIndex) = 1)
    cmd读卡.Enabled = blnEnable
End Sub

Private Sub cbo参保类别_Click()
    Dim bln隶属关系 As Boolean
    Const str隶属关系 = ";离休;在乡二等乙级伤残军人;"
    
    bln隶属关系 = (InStr(1, str隶属关系, ";" & cbo参保类别.Text & ";") <> 0)
    lbl公务员.Visible = Not bln隶属关系
    lbl公务员待遇.Visible = Not bln隶属关系
    cbo公务员.Visible = Not bln隶属关系
    cbo公务员待遇.Visible = Not bln隶属关系
    lbl隶属关系.Visible = bln隶属关系
    cbo隶属关系.Visible = bln隶属关系
End Sub

Private Sub cbo公务员_Click()
    Dim objBill As BillEdit
    On Error Resume Next
    
    Me.cbo公务员待遇.Enabled = (Me.cbo公务员.ItemData(Me.cbo公务员.ListIndex) = 0)
    '只有公务员，才允许输入封顶后医保内金额
    Set objBill = Bill(住院)
    With objBill
        .ColData(.Cols - 1) = 4
        If Me.cbo公务员待遇.Visible And Me.cbo公务员.ItemData(Me.cbo公务员.ListIndex) <> 0 Then
            .ColData(.Cols - 1) = 5
        End If
    End With
    Set objBill = Bill(门诊)
    With objBill
        .ColData(.Cols - 1) = 4
        If Me.cbo公务员待遇.Visible And Me.cbo公务员.ItemData(Me.cbo公务员.ListIndex) <> 0 Then
            .ColData(.Cols - 1) = 5
        End If
    End With
End Sub

Private Sub cbo特殊病种_Click()
    Dim blnEnable As Boolean
    blnEnable = (cbo特殊病种.ListIndex <> 0)
    txt截止日期.Enabled = blnEnable
End Sub

Private Sub cbo医疗类别_Click()
    Dim blnEnable As Boolean
    blnEnable = (cbo医疗类别.ItemData(cbo医疗类别.ListIndex) = 21)
    cbo入院方式.Enabled = blnEnable
    cbo入院类型.Enabled = blnEnable
    txt入院日期.Enabled = blnEnable
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim intPage As Integer
    Dim intCurState As Integer  '当前状态
    Dim lngRow As Long, lngRows As Long
    Dim strBorn As String       '出生日期
    Dim StrInput As String      '待遇审核用
    Dim strInsert As String     '各种金额的组合串
    Dim blnClear As Boolean
    Dim blnTrans As Boolean
    Dim objBill As BillEdit
    Dim strIdentify As String, strAddition As String
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo errHand
    
    If Not ValidData Then Exit Sub
    StrInput = txt确认卡号.Text & "|" & txt社保证号.Text
    If Not 待遇审查_北京(StrInput, True) Then Exit Sub
    
    '先取出该病人的当前状态
    intCurState = 0
    gstrSQL = "Select Nvl(当前状态,0) 当前状态 From 保险帐户 Where 卡号=[1] And 险类=[2]"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CStr(txt确认卡号.Text), TYPE_北京)
     
    If Not ChkRsState(rsTmp) Then intCurState = rsTmp!当前状态
    
    '如果在院，不允许进行身份验证
    If mbytType <> 2 And intCurState = 1 Then
        MsgBox "该参保人当前在院，不允许进行身份验证！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '转换出生日期（身份证是必输入，肯定都是19年的 ）
    If Len(txt身份证号.Text) = 15 Then
        '15位
        strBorn = "19" & Mid(txt身份证号.Text, 7, 2) & "-" & Mid(txt身份证号.Text, 9, 2) & "-" & Mid(txt身份证号.Text, 11, 2)
    Else
        '18位
        strBorn = Mid(txt身份证号.Text, 7, 4) & "-" & Mid(txt身份证号.Text, 11, 2) & "-" & Mid(txt身份证号.Text, 13, 2)
    End If
    
    '产生病人信息
    '构成字符串
    '建立病人档案信息，传入格式：
    '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);
    '8.中心代码;9.顺序号;10人员身份;11帐户余额;12当前状态;13病种ID;14在职(0,1);15退休证号;16年龄段;17灰度级
    '18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计
    strIdentify = txt确认卡号.Text                              '0卡号
    strIdentify = strIdentify & ";" & txt社保证号.Text          '1医保号
    strIdentify = strIdentify & ";"                             '2密码
    strIdentify = strIdentify & ";" & txt姓名.Text              '3姓名
    strIdentify = strIdentify & ";" & cbo性别.Text              '4性别
    strIdentify = strIdentify & ";" & strBorn                   '5出生日期
    strIdentify = strIdentify & ";" & txt身份证号.Text          '6身份证
    strIdentify = strIdentify & ";"                             '7.单位名称(编码)
    strAddition = ";0"                                          '8.中心代码
    strAddition = strAddition & ";"                             '9.顺序号
    strAddition = strAddition & ";" & cbo参保类别.Text          '10人员身份
    strAddition = strAddition & ";" & mdbl帐户余额              '11帐户余额
    strAddition = strAddition & ";" & intCurState               '12当前状态
    strAddition = strAddition & ";0"                            '13病种ID
    strAddition = strAddition & ";1"                            '14在职(1,2,3)
    strAddition = strAddition & ";"                             '15退休证号
    strAddition = strAddition & ";"                             '16年龄段
    strAddition = strAddition & ";"                             '17灰度级
    strAddition = strAddition & ";0"                            '18帐户增加累计
    strAddition = strAddition & ";0"                            '19帐户支出累计
    strAddition = strAddition & ";0"                            '20上年工资总额
    strAddition = strAddition & ";0"                            '21住院次数累计

    mlng病人ID = BuildPatiInfo(0, strIdentify & strAddition, mlng病人ID, TYPE_北京)
    '返回格式:中间插入病人ID
    If mlng病人ID > 0 Then
        mstrReturn = strIdentify & ";" & mlng病人ID & strAddition
    Else
        Exit Sub
    End If
    
    gcnBJYB.BeginTrans
    blnTrans = True
    If mbytType <> 2 Then
        '产生保险病人的基本信息
    '    病人ID,卡号,社保证号,姓名,性别,血型,身份证号,业务类型,入院类别,
    '    入院方式,入院日期,参保类别,缴费地区代码,个人帐户余额,公务员,
    '    公务员待遇 , 定点医院, 病种标识, 特殊病截止日期, 特殊病定点医院
        gstrSQL = "zl_保险帐户_INSERT(" & mlng病人ID & ",'" & txt确认卡号.Text & "','" & txt社保证号.Text & "'," & _
            "'" & txt姓名.Text & "','" & Me.cbo性别.ItemData(Me.cbo性别.ListIndex) & "',NULL,'" & txt身份证号.Text & "'," & _
            "'" & Me.cbo医疗类别.ItemData(Me.cbo医疗类别.ListIndex) & "'," & Me.cbo入院类型.ItemData(Me.cbo入院类型.ListIndex) & "," & _
            "" & Me.cbo入院方式.ItemData(Me.cbo入院方式.ListIndex) & ",TO_DATE('" & txt入院日期.Text & "','yyyy-MM-dd')" & "," & _
            "'" & Me.cbo参保类别.ItemData(Me.cbo参保类别.ListIndex) & "','" & Me.cbo报销区县.ItemData(Me.cbo报销区县.ListIndex) & "'," & _
            "" & mdbl帐户余额 & "," & Me.cbo公务员.ItemData(Me.cbo公务员.ListIndex) & "," & IIf(cbo隶属关系.Visible = False, Me.cbo公务员待遇.ItemData(Me.cbo公务员待遇.ListIndex), Me.cbo隶属关系.ItemData(Me.cbo隶属关系.ListIndex)) & "," & _
            "" & "1,'" & Me.cbo特殊病种.ItemData(Me.cbo特殊病种.ListIndex) & "'," & IIf(Me.txt截止日期.Text = "____-__-__", "NULL", "TO_DATE('" & Me.txt截止日期.Text & "','yyyy-MM-dd')") & ",1)"
        gcnBJYB.Execute gstrSQL, , adCmdStoredProc
    End If
    
    '产生历史手册消费记录
    gcnBJYB.Execute "zl_手册消费记录_DELETEALL('" & txt确认卡号.Text & "')"
    blnClear = True
    For intPage = 0 To Bill.UBound
        Set objBill = Bill(intPage)
        lngRows = objBill.Rows - 1
        For lngRow = 1 To lngRows
            If Trim(objBill.TextMatrix(lngRow, col_医疗机构)) <> "" Then
'               卡号,医疗机构,医疗类别,入院类型,入院日期,出院类型,出院日期,
'               统筹支付,大额支付,个人自付,个人自费,统筹封顶后医保内,交易流水号,清除历史记录
                strInsert = GetMoneySQL(intPage, lngRow)
                gstrSQL = "zl_手册消费记录_INSERT('" & txt确认卡号.Text & "','" & objBill.TextMatrix(lngRow, col_医疗机构) & "'," & _
                    Get医疗类别(intPage, lngRow) & "," & Get住院类型(intPage, lngRow) & ",To_Date('" & objBill.TextMatrix(lngRow, col_入院日期) & "','yyyy-MM-dd')," & _
                    Get住院类型(intPage, lngRow, False) & ",To_Date('" & objBill.TextMatrix(lngRow, IIf(intPage = 住院, col_出院日期, col_入院日期)) & "','yyyy-MM-dd')," & _
                    strInsert & ",NULL," & IIf(blnClear, 1, 0) & ")"
                gcnBJYB.Execute gstrSQL, , adCmdStoredProc
                blnClear = False
            End If
        Next
    Next
    
    '保存
    gcnBJYB.CommitTrans
    
    gComInfo_北京.卡号 = txt确认卡号.Text
    gComInfo_北京.业务类型 = Me.cbo医疗类别.ItemData(Me.cbo医疗类别.ListIndex)
    Unload Me
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTrans Then gcnBJYB.RollbackTrans
End Sub

Private Sub cmd读卡_Click()
    MsgBox "对不起，目前暂不支持持卡病人进行消费！", vbInformation, gstrSysName
    
End Sub

Private Sub cmd功能_Click(Index As Integer)
    Dim lngRow As Long
    
    On Error Resume Next
    Select Case Index
    Case 功能.增加
        Bill(tabShow.Tab).Rows = Bill(tabShow.Tab).Rows + 1
        Bill(tabShow.Tab).Row = Bill(tabShow.Tab).Rows - 1
        Bill(tabShow.Tab).SetFocus
    Case 功能.插入
        lngRow = Bill(tabShow.Tab).Row
        Bill(tabShow.Tab).msfObj.AddItem "", Bill(tabShow.Tab).Row
        Bill(tabShow.Tab).Row = lngRow
        Bill(tabShow.Tab).SetFocus
    Case 功能.删除
        Bill(tabShow.Tab).SetFocus
        SendKeys "{DELETE}", 1
    End Select
End Sub

Private Sub cmd刷新_Click()
    Dim objControl As Control
    mdbl帐户余额 = 0
    Call InitBill
    '将所有数据清空
    For Each objControl In Me.Controls
        If UCase(TypeName(objControl)) = "TEXTBOX" Then
            objControl.Text = ""
        ElseIf UCase(TypeName(objControl)) = "COMBOBOX" Then
            objControl.ListIndex = 0
        End If
    Next
    
    fra基本信息.Enabled = False
    tabShow.Enabled = False
    cmd功能(增加).Enabled = False
    cmd功能(插入).Enabled = False
    cmd功能(删除).Enabled = False
    
    '允许输入卡号或选择病人类型
    Me.cbo病人类型.Enabled = True
    Me.txt卡号.Enabled = True
    Me.txt确认卡号.Enabled = True
    cmdOK.Enabled = False
    If Me.txt入院日期.Text = "____-__-__" Then Me.txt入院日期.Text = Format(zlDatabase.Currentdate(), "yyyy-MM-dd")
    If cbo病人类型.Enabled Then cbo病人类型.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not (Me.ActiveControl.Name = "txt确认卡号" Or Me.ActiveControl.Name = "Bill") Then
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub Form_Load()
    mdbl帐户余额 = 0
    Call InitBill
    Call LoadInitData
    
    Me.txt入院日期.Text = Format(zlDatabase.Currentdate(), "yyyy-MM-dd")
End Sub

Private Sub tabShow_GotFocus()
    If Not tabShow.Enabled Then Exit Sub
    If Bill(tabShow.Tab).Active Then Bill(tabShow.Tab).SetFocus
End Sub

Private Sub txt截止日期_GotFocus()
    With txt截止日期
        .SelStart = 0
        .SelLength = 10
    End With
End Sub

Private Sub txt卡号_GotFocus()
    If Trim(txt卡号.Text) = "" Then
        txt卡号.Text = "S"
    End If
    txt卡号.SelStart = 0
    If Not (Trim(txt卡号.Text) = "" Or Trim(txt卡号.Text) = "S") Then
        txt卡号.SelLength = Len(txt卡号.Text) - 1
    End If
End Sub

Private Sub txt卡号_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub txt社保证号_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub txt身份证号_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub txt确认卡号_GotFocus()
    If Trim(txt确认卡号.Text) = "" Then
        txt确认卡号.Text = "S"
    End If
    txt确认卡号.SelStart = 0
    If Not (Trim(txt确认卡号.Text) = "" Or Trim(txt确认卡号.Text) = "S") Then
        txt确认卡号.SelLength = Len(txt确认卡号.Text) - 1
    End If
End Sub

Private Sub txt确认卡号_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        If Not ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack) Then KeyAscii = 0
        Exit Sub
    End If
    
    KeyAscii = 0
    If Trim(txt卡号.Text) <> Trim(txt确认卡号.Text) Then
        MsgBox "两次输入的手册号不相同，请再次确认手册号！", vbInformation, gstrSysName
        txt卡号.SetFocus
        Exit Sub
    End If
    If Len(txt卡号.Text) < txt卡号.MaxLength Then
        MsgBox "请输入完整的手册号（长度为" & txt卡号.MaxLength & "位）！", vbInformation, gstrSysName
        txt卡号.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(Mid(txt卡号.Text, 1, txt卡号.MaxLength - 1)) Then
        MsgBox "手册号中含有非法字符，请确认！", vbInformation, gstrSysName
        txt卡号.SetFocus
        Exit Sub
    End If
    
    '读取出该病人的基本信息
    Call ReadPatient
    '需要读取该病人在本院的历史就诊记录，减少输入量
    Call ReadBill
    
    '禁止类型选择、卡号输入
    Me.cbo病人类型.Enabled = False
    Me.cmd读卡.Enabled = False
    Me.txt卡号.Enabled = False
    Me.txt确认卡号.Enabled = False
    
    fra基本信息.Enabled = True
    tabShow.Enabled = True
    cmd功能(增加).Enabled = True
    cmd功能(插入).Enabled = True
    cmd功能(删除).Enabled = True
    cmdOK.Enabled = True
    If cbo医疗类别.Enabled Then cbo医疗类别.SetFocus
End Sub

Private Sub txt入院日期_GotFocus()
    With txt入院日期
        .SelStart = 0
        .SelLength = 10
    End With
End Sub

Public Function GetIdentify(ByVal bytType As Byte, Optional lng病人ID As Long) As String
    mlng病人ID = lng病人ID
    mbytType = bytType
    mstrReturn = ""
    Me.Show 1
    GetIdentify = mstrReturn
End Function

Private Function ValidData() As Boolean
    '对合法性进行检查
    Dim strValid As String
    Dim blnValid As Boolean
    Dim objBill As BillEdit
    Dim intPage As Integer, lngRow As Long, lngRows As Long
    On Error GoTo errHand
    '检查必输项是否输入
    '--文本框全部是必输项
    If Not CheckTEXTBOX Then Exit Function
    '--检查身份证的合法性
    If Not (Len(txt身份证号.Text) = 15 Or Len(txt身份证号.Text) = 18) Then
        MsgBox "请输入正确的身份证信息！（位数不够）", vbInformation, gstrSysName
        txt身份证号.SetFocus
        Exit Function
    End If
    '看能否分解出生日期，分析出来不是日期型则认为非法
    If Len(txt身份证号.Text) = 15 Then
        '15位
        strValid = "19" & Mid(txt身份证号.Text, 7, 2) & "-" & Mid(txt身份证号.Text, 9, 2) & "-" & Mid(txt身份证号.Text, 11, 2)
    Else
        '18位
        strValid = Mid(txt身份证号.Text, 7, 4) & "-" & Mid(txt身份证号.Text, 11, 2) & "-" & Mid(txt身份证号.Text, 13, 2)
    End If
    blnValid = IsDate(strValid)
    If Not blnValid Then
        MsgBox "请输入正确的身份证信息！", vbInformation, gstrSysName
        txt身份证号.SetFocus
        Exit Function
    End If
    
    '--再检查公务员或隶属，其它下拉框缺省有值，不必检查
    If cbo隶属关系.Visible Then
        If cbo隶属关系.ItemData(cbo隶属关系.ListIndex) = -1 Then
            MsgBox "请选择隶属关系！", vbInformation, gstrSysName
            cbo隶属关系.SetFocus
            Exit Function
        End If
    Else
        If cbo公务员.ItemData(cbo公务员.ListIndex) = 0 Then
            If cbo公务员待遇.ItemData(cbo公务员待遇.ListIndex) = -1 Then
                MsgBox "请选择公务员待遇！", vbInformation, gstrSysName
                cbo公务员待遇.SetFocus
                Exit Function
            End If
        End If
    End If
    
    '检查历史记录
    For intPage = 0 To Bill.UBound
        Set objBill = Bill(intPage)
        lngRows = objBill.Rows - 1
        For lngRow = 1 To lngRows
            If Trim(objBill.TextMatrix(lngRow, col_医疗机构)) <> "" Then
                '检查入院日期/就诊日期、出院日期是否填写
                If Trim(objBill.TextMatrix(lngRow, col_入院日期)) = "" Then
                    MsgBox "请输入第" & lngRow & "行的" & IIf(intPage = 住院, "入院日期！", "就诊日期！"), vbInformation, gstrSysName
                    tabShow.Tab = intPage
                    tabShow.SetFocus
                    Exit Function
                End If
                '只有住院需要检查出院日期
                If intPage = 住院 Then
                    If Trim(objBill.TextMatrix(lngRow, col_出院日期)) = "" Then
                        MsgBox "请输入第" & lngRow & "行的出院日期！", vbInformation, gstrSysName
                        tabShow.Tab = intPage
                        tabShow.SetFocus
                        Exit Function
                    End If
                End If
                '检查入院类型、出院类型、医疗类别是否输入
                If Trim(objBill.TextMatrix(lngRow, col_入院类型)) = "" Then
                    MsgBox "请选择第" & lngRow & "行的" & IIf(intPage = 住院, "入院类型！", "医疗类别！"), vbInformation, gstrSysName
                    tabShow.Tab = intPage
                    tabShow.SetFocus
                    Exit Function
                End If
                If intPage = 住院 Then
                    If Trim(objBill.TextMatrix(lngRow, col_出院类型)) = "" Then
                        MsgBox "请选择第" & lngRow & "行的出院类型！", vbInformation, gstrSysName
                        tabShow.Tab = intPage
                        tabShow.SetFocus
                        Exit Function
                    End If
                End If
            Else
                If lngRow < objBill.Rows - 1 Then
                    MsgBox "请删除无效空行！（第" & lngRow & "行是空行）", vbInformation, gstrSysName
                    tabShow.Tab = intPage
                    tabShow.SetFocus
                    Exit Function
                End If
            End If
        Next
    Next
    '再次检查日期合法性
    For intPage = 0 To Bill.UBound
        Set objBill = Bill(intPage)
        lngRows = objBill.Rows - 1
        For lngRow = 1 To lngRows
            With objBill
                If Trim(.TextMatrix(lngRow, col_医疗机构)) <> "" Then
                    If intPage = 住院 Then
                        '不能大于出院日期
                        If Trim(.TextMatrix(lngRow, col_出院日期)) <> "" Then
                            If .TextMatrix(lngRow, col_入院日期) > .TextMatrix(lngRow, col_出院日期) Then
                                MsgBox "入院日期不能大于出院日期！", vbInformation, gstrSysName
                                tabShow.Tab = intPage
                                tabShow.SetFocus
                                Exit Function
                            End If
                        End If
                        If lngRow > 1 Then
                            '不能小于上条记录的出院日期
                            If Trim(.TextMatrix(lngRow - 1, col_出院日期)) <> "" Then
                                If .TextMatrix(lngRow, col_入院日期) < .TextMatrix(lngRow - 1, col_出院日期) Then
                                    MsgBox "入院日期不能小于上次就诊的出院日期！", vbInformation, gstrSysName
                                    tabShow.Tab = intPage
                                    tabShow.SetFocus
                                    Exit Function
                                End If
                            End If
                        End If
                        If lngRow + 1 <= .Rows - 1 Then
                            '不能大于下一条记录的入院日期
                            If Trim(.TextMatrix(lngRow + 1, col_入院日期)) <> "" Then
                                If .TextMatrix(lngRow, col_入院日期) > .TextMatrix(lngRow + 1, col_入院日期) Then
                                    MsgBox "入院日期不能大于" & lngRow + 1 & "行登记的入院日期！", vbInformation, gstrSysName
                                    tabShow.Tab = intPage
                                    tabShow.SetFocus
                                    Exit Function
                                End If
                            End If
                        End If
                        '出院日期不能小于入院日期
                        If Trim(.TextMatrix(lngRow, col_入院日期)) <> "" Then
                            If .TextMatrix(lngRow, col_出院日期) < .TextMatrix(lngRow, col_入院日期) Then
                                MsgBox "出院日期不能小于入院日期！", vbInformation, gstrSysName
                                tabShow.Tab = intPage
                                tabShow.SetFocus
                                Exit Function
                            End If
                        End If
                        If lngRow + 1 <= .Rows - 1 Then
                            '出院日期不能大于下一条记录的入院日期
                            If Trim(.TextMatrix(lngRow + 1, col_入院日期)) <> "" Then
                                If .TextMatrix(lngRow, col_出院日期) > .TextMatrix(lngRow + 1, col_入院日期) Then
                                    MsgBox "出院日期不能大于" & lngRow + 1 & "行登记的入院日期！", vbInformation, gstrSysName
                                    tabShow.Tab = intPage
                                    tabShow.SetFocus
                                    Exit Function
                                End If
                            End If
                        End If
                    Else
                        If lngRow > 1 Then
                            If Trim(.TextMatrix(lngRow - 1, col_入院日期)) <> "" Then
                                If .TextMatrix(lngRow, col_入院日期) < .TextMatrix(lngRow - 1, col_入院日期) Then
                                    MsgBox "就诊日期不能小于上次就诊日期！", vbInformation, gstrSysName
                                    tabShow.Tab = intPage
                                    tabShow.SetFocus
                                    Exit Function
                                End If
                            End If
                        End If
                        If lngRow + 1 <= .Rows - 1 Then
                            '就诊日期不能大于下一条记录的就诊日期
                            If Trim(.TextMatrix(lngRow + 1, col_入院日期)) <> "" Then
                                If .TextMatrix(lngRow, col_入院日期) > .TextMatrix(lngRow + 1, col_入院日期) Then
                                    MsgBox "就诊日期不能大于" & lngRow + 1 & "行登记的就诊日期！", vbInformation, gstrSysName
                                    tabShow.Tab = intPage
                                    tabShow.SetFocus
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                End If
            End With
        Next
    Next
    
    ValidData = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function CheckTEXTBOX() As Boolean
    '检查文本框与MaskEdit的输入数据是否合法
    Dim objControl As Control
    For Each objControl In Me.Controls
        Select Case UCase(TypeName(objControl))
        Case "TEXTBOX"
            If objControl.Enabled Then
                If Trim(objControl.Text) = "" Then
                    MsgBox "请输入" & Mid(objControl.Name, 4) & "！", vbInformation, gstrSysName
                    objControl.SetFocus
                    Exit Function
                End If
                If LenB(StrConv(objControl.Text, vbFromUnicode)) > objControl.MaxLength Then
                    MsgBox Mid(objControl.Name, 4) & "超长（最多" & objControl.MaxLength & "个字符）！", vbInformation, gstrSysName
                    objControl.SetFocus
                    Exit Function
                End If
            End If
        Case "MASKEDBOX"
            If objControl.Enabled Then
                If Not IsDate(objControl.Text) Then
                    MsgBox "请输入合法的" & Mid(objControl.Name, 4) & "！", vbInformation, gstrSysName
                    objControl.SetFocus
                    Exit Function
                End If
            End If
        End Select
    Next
    CheckTEXTBOX = True
End Function

Private Function Get医疗类别(ByVal intPage As Integer, ByVal lngRow As Long) As Integer
    Dim str医疗类别 As String
    '获取表格中设置的医疗类别
    If intPage = 住院 Then
        '住院=21
        Get医疗类别 = 21
    Else
        '门诊特殊病=12,家庭病床=31
        str医疗类别 = Bill(intPage).TextMatrix(lngRow, col_医疗类别)
        Select Case str医疗类别
        Case "门诊特殊病"
            Get医疗类别 = 12
        Case "家庭病床"
            Get医疗类别 = 31
        End Select
    End If
End Function

Private Function Get住院类型(ByVal intPage As Integer, ByVal lngRow As Long, Optional ByVal bln入院类型 As Boolean = True) As Integer
    Dim str住院类型 As String
    '只有住院才存在入院类型与出院类型，缺省取入院类型，否则取出院类型
    If intPage <> 住院 Then
        Get住院类型 = 0
        Exit Function
    End If
    If bln入院类型 Then
        str住院类型 = Bill(intPage).TextMatrix(lngRow, col_入院类型)
        Select Case str住院类型
        Case "普通"
            Get住院类型 = 0
        Case "特殊病病人"
            Get住院类型 = 1
        Case "器官移植"
            Get住院类型 = 2
        Case "精神病"
            Get住院类型 = 3
        Case "中医医院针灸科"
            Get住院类型 = 4
        End Select
    Else
        str住院类型 = Bill(intPage).TextMatrix(lngRow, col_出院类型)
        Select Case str住院类型
        Case "正常"
            Get住院类型 = 0
        Case "转出院"
            Get住院类型 = 1
        Case "中途结算"
            Get住院类型 = 2
        End Select
    End If
End Function

Private Function GetMoneySQL(ByVal intPage As Integer, ByVal lngRow As Long) As String
    Dim strReturn As String
    Dim intStart As Integer, intEnd As Integer
    Const int住院 As Integer = 5
    Const int门诊 As Integer = 3
    '获取金额串
    intEnd = Bill(intPage).Cols - 1
    If intPage = 住院 Then
        intStart = int住院
    Else
        intStart = int门诊
    End If
    
    strReturn = ""
    For intStart = intStart To intEnd
        strReturn = strReturn & "," & Val(Bill(intPage).TextMatrix(lngRow, intStart))
    Next
    strReturn = Mid(strReturn, 2)
    GetMoneySQL = strReturn
End Function
