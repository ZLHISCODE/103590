VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.10#0"; "zlIDKind.ocx"
Begin VB.Form frmLabRequest 
   BackColor       =   &H00FDD6C6&
   BorderStyle     =   0  'None
   ClientHeight    =   7290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3480
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Enabled         =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7290
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txt年龄1 
      Height          =   300
      IMEMode         =   2  'OFF
      Left            =   2790
      MaxLength       =   5
      TabIndex        =   38
      Top             =   630
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   0  'None
      DrawStyle       =   5  'Transparent
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   1
      Left            =   1620
      ScaleHeight     =   330
      ScaleWidth      =   1770
      TabIndex        =   36
      Top             =   6450
      Width           =   1770
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   2
         Left            =   45
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   0
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         CalendarTitleBackColor=   -2147483643
         CustomFormat    =   "yy-MM-dd HH:mm:ss"
         Format          =   60424195
         CurrentDate     =   38222
      End
   End
   Begin VB.ComboBox cbo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   2
      Left            =   540
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   6450
      Width           =   1110
   End
   Begin VB.ComboBox cbo医生 
      Height          =   300
      Left            =   2145
      TabIndex        =   11
      Top             =   4995
      Width           =   1155
   End
   Begin VB.ComboBox cbo开单科室 
      Height          =   300
      ItemData        =   "frmLabRequest.frx":0000
      Left            =   540
      List            =   "frmLabRequest.frx":0002
      TabIndex        =   10
      Top             =   4995
      Width           =   1590
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "…"
      Height          =   1110
      Left            =   3030
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "选择项目(*)"
      Top             =   2010
      Width           =   300
   End
   Begin VB.CommandButton cmdExt 
      Caption         =   "…"
      Height          =   255
      Left            =   3015
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "选择检验标本"
      Top             =   5385
      Width           =   255
   End
   Begin VB.ComboBox cbo 
      Height          =   300
      Index           =   0
      Left            =   930
      TabIndex        =   15
      Top             =   5700
      Width           =   2370
   End
   Begin VB.ComboBox cbo 
      Height          =   300
      Index           =   1
      Left            =   540
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   6060
      Width           =   1110
   End
   Begin VB.ComboBox cbo性别 
      Height          =   300
      IMEMode         =   3  'DISABLE
      ItemData        =   "frmLabRequest.frx":0004
      Left            =   540
      List            =   "frmLabRequest.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   630
      Width           =   795
   End
   Begin VB.TextBox txt年龄 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1785
      MaxLength       =   3
      TabIndex        =   2
      Top             =   630
      Width           =   345
   End
   Begin VB.TextBox txtPatientDept 
      Enabled         =   0   'False
      Height          =   300
      Left            =   975
      MaxLength       =   24
      TabIndex        =   6
      Top             =   1365
      Width           =   1785
   End
   Begin VB.TextBox txtID 
      Enabled         =   0   'False
      Height          =   300
      Left            =   735
      Locked          =   -1  'True
      MaxLength       =   18
      TabIndex        =   4
      Top             =   990
      Width           =   1455
   End
   Begin VB.TextBox txtBed 
      Enabled         =   0   'False
      Height          =   300
      Left            =   2730
      MaxLength       =   10
      TabIndex        =   5
      Top             =   990
      Width           =   555
   End
   Begin VB.TextBox txt医嘱内容 
      Height          =   1080
      IMEMode         =   2  'OFF
      Left            =   135
      MaxLength       =   1000
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   2010
      Width           =   2865
   End
   Begin VB.ComboBox cboAge 
      Height          =   300
      IMEMode         =   3  'DISABLE
      ItemData        =   "frmLabRequest.frx":0008
      Left            =   2145
      List            =   "frmLabRequest.frx":001E
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   630
      Width           =   630
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&P"
      Height          =   465
      Left            =   2985
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   90
      Width           =   300
   End
   Begin MSComCtl2.DTPicker dtp 
      Height          =   300
      Index           =   1
      Left            =   930
      TabIndex        =   8
      Top             =   3165
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   60424195
      CurrentDate     =   38222
   End
   Begin zl9LisWork.VsfGrid vsf2 
      Height          =   1125
      Left            =   75
      TabIndex        =   9
      Top             =   3795
      Width           =   3180
      _ExtentX        =   5609
      _ExtentY        =   1984
   End
   Begin VB.TextBox txt姓名 
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   810
      MaxLength       =   64
      TabIndex        =   0
      ToolTipText     =   "“－”打头为病人ID、“＋”住院号、“*”门诊号、“.”挂号单号、“/”收费单据号"
      Top             =   90
      Width           =   2475
   End
   Begin VB.TextBox txt附加 
      Height          =   300
      Left            =   930
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   5355
      Width           =   2370
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   0  'None
      DrawStyle       =   5  'Transparent
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   0
      Left            =   1620
      ScaleHeight     =   330
      ScaleWidth      =   1770
      TabIndex        =   33
      Top             =   6060
      Width           =   1770
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   0
         Left            =   45
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   0
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483643
         CustomFormat    =   "yy-MM-dd HH:mm:ss"
         Format          =   60424195
         CurrentDate     =   38222
      End
   End
   Begin zlIDKind.IDKind IDKind 
      Height          =   420
      Left            =   135
      TabIndex        =   37
      Top             =   105
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   741
      IDKindStr       =   "姓|姓名|0;医|医保号|1;身|身份证号|2;IC|IC卡号|3;门|门诊号|4;就|就诊卡|5"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lbl审核未通过 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   180
      TabIndex        =   40
      Top             =   6870
      Width           =   165
   End
   Begin VB.Label lblRegister 
      BackColor       =   &H00FDD6C6&
      Caption         =   "用于判断登记"
      Height          =   225
      Left            =   2220
      TabIndex        =   39
      Top             =   3540
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "接收"
      Height          =   180
      Index           =   1
      Left            =   135
      TabIndex        =   35
      Top             =   6510
      Width           =   360
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "申请"
      Height          =   180
      Left            =   135
      TabIndex        =   34
      Top             =   5055
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "仪器(及标本分解):"
      Height          =   180
      Left            =   135
      TabIndex        =   32
      Top             =   3570
      Width           =   1530
   End
   Begin VB.Label lbl附加 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "标本种类"
      Height          =   180
      Left            =   135
      TabIndex        =   31
      Top             =   5430
      Width           =   720
   End
   Begin VB.Label lbl医嘱内容 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "申请项目"
      Height          =   180
      Left            =   135
      TabIndex        =   30
      Top             =   1785
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "标本形态"
      Height          =   225
      Index           =   5
      Left            =   135
      TabIndex        =   29
      Top             =   5775
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "核收时间"
      Height          =   180
      Index           =   6
      Left            =   135
      TabIndex        =   28
      Top             =   3240
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "采样"
      Height          =   180
      Index           =   0
      Left            =   135
      TabIndex        =   27
      Top             =   6120
      Width           =   360
   End
   Begin VB.Label lblCash 
      BackColor       =   &H00FDD6C6&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   2970
      TabIndex        =   26
      Top             =   1380
      Width           =   300
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "床号"
      Height          =   180
      Left            =   2325
      TabIndex        =   25
      Top             =   1050
      Width           =   360
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "性别"
      Height          =   180
      Left            =   135
      TabIndex        =   24
      Top             =   690
      Width           =   360
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "年龄"
      Height          =   180
      Left            =   1380
      TabIndex        =   23
      Top             =   690
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "病人科室"
      Height          =   180
      Left            =   135
      TabIndex        =   22
      Top             =   1425
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "标识号"
      Height          =   180
      Left            =   135
      TabIndex        =   21
      Top             =   1050
      Width           =   540
   End
End
Attribute VB_Name = "frmLabRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintEditMode As Integer, ItemDeptID As Long, mlngDefaultDevice As Long
'------------暂时未用------------
Private mlngSampleID As Long, mintSampleType As Integer
'-------------------------------
Private PatientType As Integer, mlng病人ID As Long, mstrNO As String '门诊收费单据号
Private mlngDefaultItemID  As Long  '上级默认的诊疗项目ID
Private mstrAuditer As String
Private iInputType As Integer
Private mblnEmerge As Boolean       '是否使用急诊标本
Private mblnPrice As Boolean        '是否按收费单号核收

'病人姓名当前输入状态，如果一直以该状态可以不输入前导符
'0：就诊卡
'1：病人ID
'2：住院号
'3：门诊号
'4：挂号单
'5：收费单据号
'6：姓名
Private rsRelativeAdvice As ADODB.Recordset '登记的相关医嘱
Private mstrExtData  As String '登记的申请项目信息
Private mlngCapID As Long '采集项目ID
Private mstrKeys As String '当前核收的申请医嘱ID
Private mlngReqDept As Long, mstrReqDoctor As String  '默认的登记科室和医生
Private mstrPrivs As String   '权限

Private mblnBarCode As Boolean
Private mblnSaveAdvice As Boolean '是否需要保存医嘱，用于修改在院病人标本信息

Private mbln微生物项目 As Boolean
Private mlngNoneHomeKey() As Long     '查找到需要被复盖的标本ID
Private mlngSourceKey() As Long       '补填时需要被复盖的标本ID
Private mRsSex As New ADODB.Recordset '性别的记录集
Private mstrNONumber As String        '记录当前录入的最后的一个标本号
Private mblnCheckIn As Boolean        '登记是可以不输入项目
Public mMakeNoRule As String          '标本序号生成规则
Private mintItemRule As Integer       '是否按项目累加的方式来生成标本号
Private mblnCard As Boolean           '是否刷卡
Private mstrMachines As String        '可以操作的仪器ID
Private mSendReport As Integer        '审核后是否自动发送报告 0=发送 1=不发送
Private mstr初审人 As String          '初审人
Private mbln划价单模式 As Boolean     '是否使用划价单模式
Private mblnEdit As Boolean           '是否正在编辑
Private mstr条码仪器 As String        '只能按条码输入的仪器（多个仪器中间使用","分隔)
Private mbln急诊 As Boolean           '当前核收的标本是否是急诊标本
Private mblnLoadLastAdvice As Boolean '是否默认上次登记项目做为后面登记的默认项目
Private mblnShowPwd As Boolean                                          '是否显示密文

'指标列属性
Private Enum ItemCol
    ID = 0
    相关ID
    结果
    标志
    结果参考
    诊疗项目ID
    排列序号
    仪器
    选择
End Enum
Private Enum mCol
    ID = 0
    紧急
    紧急医嘱
    执行状态
    所属情况
    标本类型
    标本号
    姓名
    性别
    年龄
    检验项目
    标识号
    传送
    结果次数
    医嘱id
    仪器id
    转出
    病人ID
    标本时间
    报告时间
    微生物标本
    收费单
    挂号单
    检验人
    审核人
    样本条码
    婴儿
    病人科室
    发送号
    仪器名
    主页ID
    开嘱科室ID
    报告结果
    年龄数字
    年龄单位
    床号
    申请人
    标本形态
    采样人
    采样时间
    检验标本
    NO
    接收人
    接收时间
    审核时间
    病区id
    病区名称
    定位
    执行科室ID
    标本类别
    医嘱紧急
    标本紧急
    申请科室
    申请类型
    复查
    查阅状态
    报告发送
    病人科室ID
    初审人
    初审时间
    单位
    健康号
    审核未通过
    病人来源
    门诊号
    住院号
End Enum

'条码核收的自动保存事件
Public Event ZlAutoSave(ByVal lngSampleID As Long)

'-------------------------------------------- 2007-10-26 加入一卡通支持
Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private mobjICCard As Object
Private Enum IDKinds
    C0姓名 = 0
    C1医保号 = 1
    C2身份证号 = 2
    C3IC卡号 = 3
    C4门诊号 = 4
    C5就诊卡 = 5
End Enum
Private mobjSquareCard As Object                                        '取卡类型

Private Sub Txt姓名Exec()
    Dim strInput As String
    Dim rsTmp As New ADODB.Recordset, i As Integer
    Dim strField As String
    Dim strBarCode As String
    Dim rsDept As ADODB.Recordset, strSQL As String
    Dim intSelect As Integer
    Dim intPatientSource As Integer                     '病人来源
    Dim rsPaInfo As New ADODB.Recordset                 '提取标本记录中的"标识号"
    Dim blnGetPaInfo As Boolean
    Dim strAge As String
    Dim aAge() As String
    Dim str医嘱ID As String
    Dim rs As New ADODB.Recordset
    
    Dim intMainID   As Integer                          '主页id
    Dim strGetSql As String
    Dim rsTest As Recordset
    
    On Error GoTo errH
    If Len(Trim(txt姓名)) = 0 Or Me.txt姓名.Enabled = True Or Me.txt姓名 = Me.txt姓名.Tag Then Exit Sub

    If txt姓名 <> txt姓名.Tag Then mlng病人ID = 0
    If mlng病人ID > 0 Then
        strSQL = "select 病人来源 from 检验标本记录 where id = [1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, mlngSampleID)
        If rsTmp.RecordCount > 0 Then
            intPatientSource = Nvl(rsTmp("病人来源"), PatientType)
        Else
            intPatientSource = PatientType
        End If

        If txt姓名 = txt姓名.Tag And intPatientSource <> 3 Then Exit Sub
    Else
        If txt姓名 = txt姓名.Tag Then Exit Sub
    End If

    mblnSaveAdvice = True
'    Cancel = Not StrIsValid(txt姓名.Text, txt姓名.MaxLength)
    If StrIsValid(txt姓名.Text, txt姓名.MaxLength) = False Then Exit Sub


    '初始病人信息
    '2007-10-26 加入一卡通支付
    If IDKind.Tag = "医保卡" Or IDKind.IDKind = IDKinds.C1医保号 Then
        Set rsTmp = GetPatient("卡" & txt姓名)
        IDKind.Tag = ""
    ElseIf IDKind.Tag = "身份证" Or IDKind.IDKind = IDKinds.C2身份证号 Then
        Set rsTmp = GetPatient("证" & txt姓名)
        IDKind.Tag = ""
    Else
        Set rsTmp = GetPatient(txt姓名)
    End If
    If rsTmp.RecordCount > 0 Then
        If iInputType = 2 Then
            If intMainID = 0 Then
                intMainID = Val(rsTmp("主页ID") & "")
                mlng病人ID = Val(rsTmp("病人ID") & "")
                If intMainID <> 0 Then
                    strGetSql = "Select 主页id, 出院日期 From 病案主页 Where 病人id = [1] and 主页id=[2] "
                    Set rsTest = zlDatabase.OpenSQLRecord(strGetSql, Me.Caption, mlng病人ID, intMainID)
                    If Nvl(rsTest("出院日期")) <> "" Then
                        If MsgBox("该病人已出院，是否继续执行！", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
                            '清空显示
                            mlng病人ID = 0
                            mstrKeys = ""
                            Me.txtPatientDept = "": Me.txtPatientDept.Tag = 0
                            Me.txtID = "": Me.txtBed = ""
                            Me.txt姓名.Text = "": 'Cancel = True
                            Me.txt姓名.Enabled = True:             Me.txt姓名.SetFocus
                            Me.txt医嘱内容 = ""
                            Me.txt附加 = ""
                            Me.txt医嘱内容.Tag = "": Me.txt附加.Tag = "": mlngCapID = 0
                            Me.txt年龄.Text = "": Me.txt年龄1.Text = ""
                            vsf2.Rows = 1
                            vsf2.Rows = 2
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
    End If

    If rsTmp.EOF = True And IDKind.IDKind = IDKinds.C0姓名 Then
        Set rsTmp = GetPatientInfo(txt姓名)
        blnGetPaInfo = True
    End If

    strBarCode = txt姓名
    If rsTmp.RecordCount <= 0 Then
        mlng病人ID = 0
        '登记新病人
        mstrKeys = ""
'        Me.txt年龄 = "": Me.cboAge.ListIndex = 0
        Me.txtPatientDept = "": Me.txtPatientDept.Tag = 0
        Me.txtID = "": Me.txtBed = ""
        '如果想输入院内病人，则不允许继续
        If InStr("+-*./", Left(Me.txt姓名.Text, 1)) > 0 Or mblnBarCode Then
            Me.txt姓名.Text = "": 'Cancel = True
            Me.txt姓名.Enabled = True:             Me.txt姓名.SetFocus
            Exit Sub
        End If
        If mblnBarCode = True Then
            MsgBox "没有找到条码，请查看条码是否已取消绑定！", vbInformation, Me.Caption
            Me.txt姓名.Text = "": 'Cancel = True
            Me.txt姓名.Enabled = True:             Me.txt姓名.SetFocus
            Exit Sub
        End If
        If IDKind.IDKind = IDKinds.C0姓名 Then
            PatientType = 1
            '处理登记的默认科室、医生
            If mlngReqDept > 0 Then
                cbo开单科室.ListIndex = FindComboItem(cbo开单科室, mlngReqDept)
                Me.cbo医生.Text = mstrReqDoctor
            End If
            SetPatientInfoWrite False
            Me.txt姓名.Tag = Me.txt姓名.Text
        Else
            Select Case IDKind.IDKind
                Case IDKinds.C1医保号
                    MsgBox "没有找到医保号为<" & Me.txt姓名.Text & ">的病人！", vbInformation, Me.Caption
                Case IDKinds.C2身份证号
                    MsgBox "没有找到身份证为<" & Me.txt姓名.Text & ">的病人！", vbInformation, Me.Caption
                Case IDKinds.C3IC卡号
                    MsgBox "没有找到IC卡号为<" & Me.txt姓名.Text & ">的病人！", vbInformation, Me.Caption
                Case IDKinds.C4门诊号
                    MsgBox "没有找到门诊号为<" & Me.txt姓名.Text & ">的病人！", vbInformation, Me.Caption
                Case IDKinds.C5就诊卡
                    MsgBox "没有找到就诊卡号为<" & Me.txt姓名.Text & ">的病人！", vbInformation, Me.Caption
            End Select
            Me.txt姓名.Text = "": 'Cancel = True
            Me.txt姓名.Enabled = True:             Me.txt姓名.SetFocus
            Exit Sub
        End If
    Else
        On Error Resume Next
        Me.txt姓名.Text = Nvl(rsTmp("姓名"))
'        Me.txt年龄 = IIf(IsNull(rsTmp("年龄")), "", Val(rsTmp("年龄"))): If Me.txt年龄 = "0" Then Me.txt年龄 = ""
'        Me.cboAge.Text = IIf(IsNull(rsTmp("年龄")), "岁", Replace(rsTmp("年龄"), Val(rsTmp("年龄")), ""))
        If Trim(Nvl(rsTmp("年龄"))) <> "" And Trim(Nvl(rsTmp("年龄1"))) <> "" Then
            If rsTmp("年龄") <> rsTmp("年龄1") Then
'                MsgBox "年龄和出生日期计算的年龄不符！" & _
                        vbCrLf & "出生日期计算年龄为:" & rsTmp("年龄1") & _
                        vbCrLf & "当前年龄为:" & rsTmp("年龄")
                Me.txt年龄.ForeColor = vbRed
            End If
        End If
        Me.txt年龄.Text = "": Me.txt年龄1.Text = ""
        '不使用自动计算的年龄,使用下达医嘱的年龄
        'strAge = IIf(Trim(Nvl(rsTmp("年龄1"))) = "", Nvl(rsTmp("年龄")), Nvl(rsTmp("年龄1")))
        strAge = Nvl(rsTmp("年龄"))
        
        strAge = Replace(strAge, "小时", "时")
        strAge = Replace(strAge, "分钟", "分")
        
        If Trim(Replace(Replace(Replace(Replace(Replace(strAge, "岁", ""), "月", ""), "天", ""), "时", ""), "分", "")) <> "" Then
            If InStr(strAge, "成人") > 0 Or InStr(strAge, "婴儿") > 0 Then
                Me.txt年龄.Text = ""
                Me.cboAge.Text = Trim(strAge)
            Else
                strAge = Replace(Replace(Replace(Replace(Replace(strAge, "岁", "岁;"), "月", "月;"), "天", "天;"), "时", "时;"), "分", "分;")
                'strAge = Replace(strAge, "分钟", "婴儿")
                aAge = Split(strAge, ";")
                If UBound(aAge) = 1 Then
                    Me.txt年龄.Text = Val(aAge(0))
                    Me.cboAge.Text = Replace(Replace(Right(aAge(0), 1), "分", "分钟"), "时", "小时")
                Else
                    Me.txt年龄.Text = Val(aAge(0))
                    Me.cboAge.Text = Replace(Replace(Right(aAge(0), 1), "分", "分钟"), "时", "小时")
                    Me.txt年龄1.Text = Val(aAge(1)) & Replace(Replace(Right(aAge(1), 1), "分", "分钟"), "时", "小时")
                End If
            End If
        Else
            If Val(strAge) <> 0 Then
                Me.txt年龄.Text = Val(strAge)
            End If
            Me.cboAge.ListIndex = 0
        End If
        Me.txt姓名.Tag = Me.txt姓名.Text
'        Me.txt年龄 = IIf(IsNull(rsTmp("年龄1")), "", IIf(IsNumeric(rsTmp("年龄1")), Val(rsTmp("年龄1")), Mid(rsTmp("年龄1"), 1, Len(rsTmp("年龄1")) - 1)))
'        If Me.txt年龄 = "0" Then Me.txt年龄 = ""
'        Me.cboAge.Text = IIf(IsNull(rsTmp("年龄")), "岁", Right(rsTmp("年龄1"), 1))
        If cboAge.ListIndex = -1 Then cboAge.ListIndex = 0
        Me.cbo性别 = Nvl(rsTmp("性别")) ' CombIndex(cbo性别, Nvl(rsTmp("性别")))
        mlng病人ID = Nvl(rsTmp("病人ID"), 0): PatientType = Nvl(rsTmp("PatientType"), 1)

        '设置默认开单科室、医生
        cbo开单科室.ListIndex = FindComboItem(cbo开单科室, Nvl(rsTmp("病人科室"), 0))
'        DoEvents
        gintSelectFocus = 2
        strField = ""
        strField = rsTmp.Fields("医生").Name
        If strField = "医生" Then
            Me.cbo医生.Text = Nvl(rsTmp("医生"))
            For i = 0 To Me.cbo医生.ListCount - 1
                If Me.cbo医生.List(i) Like Nvl(rsTmp("医生")) Then
                    Me.cbo医生.ListIndex = i
                    Exit For
                End If
            Next
        End If
        '显示病人科室
        If IsNumeric(rsTmp("病人科室")) = True Then
            strSQL = "Select 名称 From 部门表 Where ID=[1]"
            Set rsDept = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CLng(Nvl(rsTmp("病人科室"), 0)))
            If rsDept.EOF Then
                Me.txtPatientDept = "": Me.txtPatientDept.Tag = 0
            Else
                Me.txtPatientDept.Text = rsDept("名称"): Me.txtPatientDept.Tag = Nvl(rsTmp("病人科室"), 0)
            End If
        Else
            Me.txtPatientDept = Nvl(rsTmp("病人科室"))
        End If

        Me.txtID = Nvl(rsTmp("住院号")): If Len(Me.txtID) = 0 Then Me.txtID = Nvl(rsTmp("门诊号"))

        Me.txtBed = Nvl(rsTmp("当前床号"))

        '当检验标本记录里有时优先显示检验标本记录里的信息
        If Trim(Me.txtID) = "" Or Trim(Me.txtBed) = "" Or Trim(Me.txtPatientDept) = "" Then
            If blnGetPaInfo = True Then
                Me.txtID = Nvl(rsTmp("标识号"), Me.txtID)
                Me.txtBed = Nvl(rsTmp("床号"), Me.txtBed)
                Me.txtPatientDept = Nvl(rsTmp("病人科室"), Me.txtPatientDept)
            End If
        End If

        '处理登记的默认科室、医生
        If Me.cbo开单科室.ListIndex = -1 And mlngReqDept > 0 Then
            cbo开单科室.ListIndex = FindComboItem(cbo开单科室, mlngReqDept)
            Me.cbo医生.Text = mstrReqDoctor
        End If
    End If
    '核收时选择检验申请
    If mlng病人ID > 0 And Not mintEditMode = 1 And (intPatientSource <> 3 Or mblnBarCode = True) Then
        intSelect = OpenSelect(strBarCode, True)
'        DoEvents
        gintSelectFocus = 2
        Select Case intSelect
            Case 0
                '没有匹配的项目
                mstrKeys = ""
                If mlng病人ID = 0 Or mblnBarCode Then
'                    mintFocusItem = FocusItem.姓名
'                    MsgBox "条码已被核收！", vbInformation, Me.Caption
                    mlng病人ID = 0
                    txt姓名.Text = ""
                    If Me.txt姓名.Enabled = True Then
                        txt姓名.SetFocus
                    End If
'                    Cancel = True
                Else
                    '允许登记
                    SetAdviceEnable True
                    If mlngDefaultItemID > 0 Then
                        strSQL = "select 标本部位 from 诊疗项目目录 where id = [1] "
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, mlngDefaultItemID)
                        AdviceSet检查手术 3, mlngDefaultItemID & ";" & rsTmp("标本部位")
                        mstrExtData = mlngDefaultItemID & ";" & rsTmp("标本部位")
                        '获取采集方式
                        Set rsTmp = SelectCap(Split(Split(mstrExtData, ";")(0), ",")(0))
                        If rsTmp Is Nothing Then
                            MsgBox "没有定义标本采集方式，请到诊疗项目管理中设置。", vbInformation, gstrSysName
                            Exit Sub
                        End If
                        mlngCapID = rsTmp("ID")
                    End If
                    If mblnLoadLastAdvice = True And mstrExtData <> "" Then
                        AdviceSet检查手术 3, mstrExtData
                    End If
                    If rsRelativeAdvice Is Nothing Then
                        Me.txt医嘱内容 = ""
                        Me.txt附加 = ""
                        Me.txt医嘱内容.Tag = "": Me.txt附加.Tag = "": mlngCapID = 0
                    Else
                        txt医嘱内容.Text = Get检查手术名称(2, "")
                        txt医嘱内容.Text = txt医嘱内容.Text & "(" & Split(mstrExtData, ";")(1) & ")"
                        txt医嘱内容.Tag = txt医嘱内容.Text
                        Me.txt附加 = Split(mstrExtData, ";")(1)
                        If mintEditMode = 0 Then
                            Call LoadDefaultData
                            Call SelectDefault
'                            mintFocusItem = FocusItem.标本号

'                            DoEvents
                            vsf2.Col = 2
                            vsf2.ShowCell vsf2.Row, vsf2.Col
                            vsf2.SetFocus
                            gintSelectFocus = 2
                        Else
                            '赋医嘱ID
                            Call SelectDefault
                        End If
                    End If
                End If
            Case 1
                '选取了一个项目
                SetAdviceEnable False   '不允许登记
                '产生仪器和标本号
                If mintEditMode = 0 Then
                    Call LoadDefaultData
                    Call SelectDefault
'                    mintFocusItem = FocusItem.标本号

                    vsf2.Col = 2
                    vsf2.ShowCell vsf2.Row, vsf2.Col
                    vsf2.SetFocus
                Else
                    '赋医嘱ID
                    Call SelectDefault
                    vsf2.Col = 2
                    vsf2.ShowCell vsf2.Row, vsf2.Col
                    vsf2.SetFocus
                End If
                '条码自动保存
                If mblnBarCode Then Me.cbo性别.SetFocus
            Case 2
                '取消了本次选择
'                mintFocusItem = FocusItem.姓名

                mlng病人ID = 0
                mstrKeys = ""
                txt姓名.Text = ""
                If Me.txt姓名.Enabled = True Then
                    txt姓名.SetFocus
                End If
'                Cancel = True
            Case 3
                Me.txt姓名.Enabled = True
                txt姓名.SetFocus
        End Select
    Else
        SetAdviceEnable True
        If mlngDefaultItemID > 0 Then
            strSQL = "select 标本部位 from 诊疗项目目录 where id = [1] "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, mlngDefaultItemID)
            AdviceSet检查手术 3, mlngDefaultItemID & ";" & rsTmp("标本部位")
            mstrExtData = mlngDefaultItemID & ";" & rsTmp("标本部位")
            '获取采集方式
            Set rsTmp = SelectCap(Split(Split(mstrExtData, ";")(0), ",")(0))
            If rsTmp Is Nothing Then
                MsgBox "没有定义标本采集方式，请到诊疗项目管理中设置。", vbInformation, gstrSysName
                Exit Sub
            End If
            mlngCapID = rsTmp("ID")
        End If

        If mblnLoadLastAdvice = True And mstrExtData <> "" Then
            AdviceSet检查手术 3, mstrExtData
        End If

        If rsRelativeAdvice Is Nothing Then
            Me.txt医嘱内容 = ""
            Me.txt附加 = ""
            Me.txt医嘱内容.Tag = "": Me.txt附加.Tag = "": mlngCapID = 0
        Else
            txt医嘱内容.Text = Get检查手术名称(2, "")
            txt医嘱内容.Text = txt医嘱内容.Text & "(" & Split(mstrExtData, ";")(1) & ")"
            txt医嘱内容.Tag = txt医嘱内容.Text
            Me.txt附加 = Split(mstrExtData, ";")(1)
            If mintEditMode <= 1 Then
                Call LoadDefaultData
                Call SelectDefault

                vsf2.Col = 2
                vsf2.ShowCell vsf2.Row, vsf2.Col
                vsf2.SetFocus
            Else
                '赋医嘱ID
                Call SelectDefault
            End If
        End If
    End If

    If mlng病人ID > 0 Then
        txt姓名.Tag = txt姓名.Text
    End If

    If mblnCheckIn = True And Me.txt医嘱内容.Tag = "" And mintEditMode <> 3 Then
        If mlngDefaultDevice > 0 And mintEditMode <> 3 Then
            gstrSql = "select 名称 from 检验仪器 where id = [1] "
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngDefaultDevice)
            vsf2.TextMatrix(1, 1) = Nvl(rsTmp("名称"))
            vsf2.RowData(1) = mlngDefaultDevice
        Else
            vsf2.TextMatrix(1, 1) = "[手工]"
            vsf2.RowData(1) = mlngDefaultDevice
        End If
        '取标本号
        If vsf2.TextMatrix(1, 5) = "-1" Then
            '急诊
            vsf2.TextMatrix(1, 2) = TransSampleNO_PH(Val(CalcNextCode(Val(vsf2.RowData(1)), 1, 1)), vsf2.RowData(1))
        Else
            vsf2.TextMatrix(1, 2) = TransSampleNO_PH(Val(CalcNextCode(Val(vsf2.RowData(1)), 1, 0)), vsf2.RowData(1))
        End If
    End If

    '--------------------------------------------------------------------------------------------------------------------------------
    '当执行完成有自动审核的费用时，对病人费用进行记帐报警。
    gstrSql = " select /*+ rule */ id from 病人医嘱记录 where id in (Select * From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))) " & _
              " Union All " & _
              " select /*+ rule */ id from 病人医嘱记录 where 相关id in (Select * From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))) "
    Set rs = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mstrKeys)
    Do While Not rs.EOF
        str医嘱ID = str医嘱ID & "," & rs("id")
        rs.MoveNext
    Loop
    str医嘱ID = Mid(str医嘱ID, 2)
    If Chk划价费用(Me, str医嘱ID, 0) = False And Trim(str医嘱ID) <> "" Then
        Exit Sub
    End If
    '----------------------------------------------------------------------------------------------------------------------------------

    Exit Sub
errH:
    Call InitEdit
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function zlRefresh(ByVal Row As ReportRow) As Boolean
'显示标本申请信息
'lngSampleID：标本记录ID
    Dim rs As New ADODB.Recordset
    Dim mstrSql As String
    Dim strTmp As String
    Dim strAge As String
    Dim aAge() As String

    On Error GoTo ErrHand

    mblnEdit = False
    ClearItem

    Me.cbo开单科室.ListIndex = -1: Me.cbo医生.ListIndex = -1
    Me.cbo(0).ListIndex = -1: Me.cbo(1).ListIndex = -1

    On Error Resume Next
        Me.txt姓名 = Row.Record(mCol.姓名).Value
        strTmp = "名称='" & CStr(Row.Record(mCol.性别).Value) & "'"
        mRsSex.filter = strTmp
        If mRsSex.EOF = False Then
            Me.cbo性别.Text = mRsSex!编码 & "-" & mRsSex!名称
        End If

'        Me.cbo性别.Text = Row.Record(mCol.性别).Value
'        Me.txt年龄 = IIf(IsNull(rs("年龄")), "", Val(rs("年龄"))): If Me.txt年龄 = "0" Then Me.txt年龄 = ""
'        Me.txt年龄 = IIf(IsNull(rs("年龄")), "", IIf(IsNumeric(rs("年龄")), Val(rs("年龄")), Mid(rs("年龄"), 1, Len(rs("年龄")) - 1))): If Me.txt年龄 = "0" Then Me.txt年龄 = ""
        Me.txt年龄.Text = "": Me.txt年龄1.Text = ""
        strAge = Row.Record(mCol.年龄).Caption
        
        strAge = Replace(strAge, "小时", "时")
        strAge = Replace(strAge, "分钟", "分")
        
        If Trim(Replace(Replace(Replace(Replace(Replace(strAge, "岁", ""), "月", ""), "天", ""), "时", ""), "分", "")) <> "" Then
            If InStr(strAge, "成人") > 0 Or InStr(strAge, "婴儿") > 0 Then
                Me.txt年龄.Text = ""
                Me.cboAge.Text = Trim(strAge)
            Else
                strAge = Replace(Replace(Replace(Replace(Replace(strAge, "岁", "岁;"), "月", "月;"), "天", "天;"), "时", "时;"), "分", "分;")
                aAge = Split(strAge, ";")
                If UBound(aAge) = 1 Then
                    Me.txt年龄.Text = Val(aAge(0))
                    Me.cboAge.Text = Replace(Replace(Right(aAge(0), 1), "分", "分钟"), "时", "小时")
                Else
                    Me.txt年龄.Text = Val(aAge(0))
                    Me.cboAge.Text = Replace(Replace(Right(aAge(0), 1), "分", "分钟"), "时", "小时")
                    Me.txt年龄1.Text = Val(aAge(1)) & Replace(Replace(Right(aAge(1), 1), "分", "分钟"), "时", "小时")
                End If
            End If
        Else
            Me.txt年龄.Text = Val(strAge)
            Me.cboAge.ListIndex = 0
        End If

        Me.txt年龄 = Row.Record(mCol.年龄数字).Value
        Me.cboAge = Row.Record(mCol.年龄单位).Value

        'Me.cboAge.Text = IIf(IsNull(rs("年龄")), "岁", Right(rs("年龄"), 1))
        If cboAge.ListIndex = -1 Then cboAge.ListIndex = 0
        Me.txtPatientDept = Row.Record(mCol.病人科室).Value

        Me.txtID = Row.Record(mCol.标识号).Value
        If Me.txtID.Text = "" Then Me.txtID.Text = Row.Record(mCol.NO).Value
        '处理补填时是否可以修改标识号和科室
        If Row.Record(mCol.申请类型).Value = 1 And Row.Record(mCol.病人来源).Value = 3 And Row.Record(mCol.住院号).Value = "" _
                And Row.Record(mCol.门诊号).Value = "" Then
            Me.txtID.Tag = "可修改"
        End If

        Me.txtBed = Row.Record(mCol.床号).Value

        Me.cbo开单科室.Text = Row.Record(mCol.申请科室).Value

        Me.cbo医生.Text = Row.Record(mCol.申请人).Value
        Me.txt附加 = Row.Record(mCol.检验标本).Value

        Me.DTP(1).Value = Row.Record(mCol.标本时间).Value
        Me.cbo(1).Text = Row.Record(mCol.采样人).Value
        Me.DTP(0).Value = Row.Record(mCol.采样时间).Value

        mstr初审人 = Row.Record(mCol.初审人).Value

        '没有采样人时不显示采样时间
        If Trim(Me.cbo(1).Text) = "" Then
            Me.cbo(1).Visible = False
            Me.DTP(0).Visible = False
            lbl(0).Visible = False
            Me.Picture1(0).Visible = False
        Else
            Me.cbo(1).Visible = True
            Me.DTP(0).Visible = True
            lbl(0).Visible = True
            Me.Picture1(0).Visible = True
        End If

        Me.cbo(0).Text = Row.Record(mCol.标本形态).Value

        If Row.Record(mCol.接收人).Value = "" Then
            Me.cbo(2).Visible = False
            Me.DTP(2).Visible = False
            lbl(1).Visible = False
            Me.Picture1(1).Visible = False
        Else
            Me.cbo(2).Visible = True
            Me.DTP(2).Visible = True
            lbl(1).Visible = True
            Me.Picture1(1).Visible = True
            Me.cbo(2).Text = Row.Record(mCol.接收人).Value
            Me.DTP(2).Value = Row.Record(mCol.接收时间).Value
        End If

        With vsf2
            .Rows = 2
            .RowData(1) = IIf(Row.Record(mCol.仪器id).Value = "", -1, Row.Record(mCol.仪器id).Value)
            .TextMatrix(1, 1) = IIf(Row.Record(mCol.仪器名).Value = "", "手工", Row.Record(mCol.仪器名).Value)
            .TextMatrix(1, 2) = Row.Record(mCol.标本号).Caption
            .TextMatrix(1, 4) = Val(Row.Record(mCol.医嘱id).Value)
            .TextMatrix(1, 5) = IIf(Row.Record(mCol.标本类别).Value = 1, -1, 0) '   IIf(rs("标本类别") = 0, 0, -1)
            '---- 根据是否区分急诊标志，显示急诊列
            If mblnEmerge Then
                .Body.ColWidth(5) = 250
            Else
                .Body.ColWidth(5) = 0
            End If
        End With
        Me.txt医嘱内容 = Row.Record(mCol.检验项目).Value
        Me.txt医嘱内容.Tag = Row.Record(mCol.检验项目).Value

        If lbl(1).Visible = False Then
            Me.lbl审核未通过.Top = lbl(1).Top
        End If

        If lbl(0).Visible = False Then
            Me.lbl审核未通过.Top = lbl(0).Top
        End If
        Me.lbl审核未通过.Caption = Trim(Row.Record(mCol.审核未通过).Value)
        Me.lbl审核未通过.Visible = (Me.lbl审核未通过.Caption <> "")
        Me.lbl审核未通过.Top = Me.lbl(1).Top + Me.lbl(1).Height + 200

        If Row.Record(mCol.所属情况).Value = "院外" Then
            lblCash.Caption = ""
        Else

            Select Case ShowCharge(Row.Record(mCol.ID).Value)
                Case -1     '未有收费单据
                    lblCash.Caption = ""
                Case 0      '划价单
                    lblCash.Caption = "划"
                Case 1, 2     '已完成收费
                    lblCash.Caption = "收"
            End Select
        End If
        Me.lblRegister.Caption = Nvl(Row.Record(mCol.申请类型).Value, 0)
        mbln微生物项目 = IIf(Val(Row.Record(mCol.微生物标本).Value) = 1, True, False)
    SetPatientInfoWrite True
    zlRefresh = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ZlEditStart(ByVal intEditType As Integer, ByVal lngDeptID As Long, ByVal lngDeviceID As Long, _
    Optional ByVal lngSampleID As Long = 0, _
    Optional ByVal lngAdviceID As Long = 0, Optional ByVal intSampleType As Integer = 0, _
    Optional ByVal strAuditer As String = "", Optional ByVal lngDefaultItemID As Long, _
    Optional ByVal lngPatientID As Long) As Boolean
'编辑标本申请信息
'intEditType：＝0－核收、1－登记、2－重新核收、3－补填申请
'lngDeptID：当前检验室
'lngDeviceID：默认检验仪器
'lngSampleID：可选。当前修改的标本ID
'lngAdviceID：可选。当前要核收的检验申请医嘱ID
'intSampleType：可选。标本类别，0－普通、1－急诊
'lngDefaultItemID ： 可选。默认项目ID
    mintEditMode = intEditType: ItemDeptID = lngDeptID
    mlngDefaultDevice = lngDeviceID: mlngSampleID = lngSampleID
    mstrKeys = IIf(mintEditMode = 1, "", lngAdviceID): mintSampleType = intSampleType
    mstrAuditer = strAuditer
    mlngDefaultItemID = lngDefaultItemID

    If Val(mstrKeys) = 0 Then mstrKeys = ""

    mblnSaveAdvice = False
    If mintEditMode = 0 Or mintEditMode = 1 Then
        Me.lblRegister.Caption = 0
    End If

    If InitEdit = False Then
        ZlEditStart = False
        Exit Function
    End If

'    SetActiveWindow Me.Hwnd
    Me.txt姓名.Enabled = True
    Me.txt姓名.SetFocus

    ZlEditStart = True
    mblnEdit = True

    '当从待处理列表选择时处理
    If lngPatientID > 0 Then
        Me.txt姓名.Enabled = False
        Me.txt姓名.Text = "-" & lngPatientID
'        Call txt姓名_Validate(False)
        Call Txt姓名Exec
        Me.txt姓名.Enabled = True
    End If
    gintSelectFocus = 2
End Function

Public Function ZlSave(Optional ByVal intEditState As Integer) As Long
'保存当前标本编辑信息
'intEditMode：当前编辑模式，1－保存后继续编辑、0－保存后结束编辑
    If ValidData = False Then Exit Function
    If SaveData(intEditState) = False Then Exit Function

    On Error Resume Next

    '清除控件内容
    Call ResetVsf(vsf2)

    Me.txt姓名 = "": Me.cbo性别.ListIndex = -1: Me.txt年龄 = "": mstrKeys = "": Me.cboAge.ListIndex = 0: mstrNO = ""
    mblnEdit = False
    txt姓名.SetFocus
    ZlSave = mlngSampleID

End Function

Public Function ZlRefuse() As Boolean
'拒绝当前标本
'intEditMode：当前编辑模式，1－拒绝后继续编辑、0－拒绝后结束编辑
    ZlRefuse = True
    mblnEdit = False
End Function

Public Function ZlCancel() As Boolean
'取消编辑
    Me.Enabled = False

    mintEditMode = -1
    mstrNO = ""
    ClearItem
    mblnEdit = False
    ZlCancel = True
End Function

Private Function InitEdit() As Boolean
    Dim strSQL As String, rs As New ADODB.Recordset, i As Long

    On Error GoTo ErrHand
    mblnBarCode = False
'    mblnSaveAdvice = True

    PatientType = 1: mlng病人ID = 0: Me.txt年龄.ForeColor = vbBlack
    iInputType = -1: mstrNO = ""
    Set rsRelativeAdvice = Nothing

    strSQL = "SELECT 名称,0 AS ID FROM 检验标本形态"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rs.BOF = False Then Call AddComboData(cbo(0), rs)

    '初始编辑项目
    InitDepts
    '根据编辑状态处理界面布局
    cmdOpen.Visible = Not (mintEditMode = 1)
    Me.cbo(1).Enabled = mblnBarCode
    SetAdviceEnable (mintEditMode = 1 Or Me.lblRegister = 0)

    If mintEditMode > 1 Then
        vsf2.Body.Editable = flexEDKbdMouse  'flexEDKbdMouse   'flexEDNone
        DTP(1).Enabled = IIf(mbln微生物项目 = True, True, False)
        If Len(Trim(Me.txt姓名)) = 0 Then
            DTP(0).Value = Format(zlDatabase.Currentdate, DTP(0).CustomFormat)
            If Format(DTP(0).Value, "yyyy-mm-dd") = Format(DTP(1).Value, "yyyy-mm-dd") Then DTP(1).Value = DTP(0).Value
        End If
    Else
        vsf2.Body.Editable = flexEDKbdMouse
        DTP(1).Value = Format(zlDatabase.Currentdate, DTP(1).CustomFormat)
        DTP(1).Enabled = True
        DTP(0).Value = DTP(1).Value
    End If

    Me.Enabled = True

    '初始标本的申请项目、科室
    If mintEditMode > 1 Then
        InitSampleInfo mlngSampleID
        If mlng病人ID = 0 Then
            ClearItem
        End If
    Else
        ClearItem
    End If

    InitEdit = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub SetAdviceEnable(ByVal blnEnable As Boolean)

    If InStr(1, mstrPrivs, "直接申请") = 0 Then blnEnable = False: cmdSel.Enabled = False
    Me.txt医嘱内容.Enabled = blnEnable: Me.cmdSel.Enabled = blnEnable
    If Me.txtID.Tag <> "" Then
        Me.txtBed.Enabled = blnEnable
        Me.txtBed.Locked = Not blnEnable
        Me.txtPatientDept.Enabled = blnEnable
        Me.txtPatientDept.Locked = Not blnEnable
        Me.txtID.Enabled = blnEnable
        Me.txtID.Locked = Not blnEnable
    End If
'    If mbln微生物项目 = False Then
'        Me.txt附加.Enabled = blnEnable
'        Me.cmdExt.Enabled = blnEnable
'    Else
        Me.txt附加.Enabled = True
        Me.cmdExt.Enabled = True
'    End If

'    Me.cbo开单科室.Enabled = blnEnable: Me.cbo医生.Enabled = blnEnable
End Sub

Private Sub AutoSave()
'自动保存当前标本编辑信息（条码方式）
    If ValidData = False Then Exit Sub
    If SaveData = False Then Exit Sub

    '清除控件内容
    Call ResetVsf(vsf2)

    Me.txt姓名 = "": Me.cbo性别.ListIndex = -1: Me.txt年龄 = "": Me.txt年龄1 = "": mstrKeys = "":   Me.cboAge.ListIndex = 0
    txt姓名.SetFocus

    RaiseEvent ZlAutoSave(mlngSampleID)
End Sub

Private Sub ClearItem()
    Me.txt姓名 = "": Me.txt姓名.Tag = "": Me.cbo性别.ListIndex = -1: Me.txt年龄 = "": Me.txt年龄1 = "":  Me.cboAge.ListIndex = 0
    Me.txtPatientDept = "": Me.txtID = "": Me.txtID.Tag = "": Me.txtBed = "": Me.txtPatientDept.Tag = 0
    Me.txt医嘱内容 = "": Me.txt附加 = "": Me.txt年龄.ForeColor = vbBlack
    Me.txt医嘱内容.Tag = "": Me.txt附加.Tag = ""
'    Me.lblCash.Font.Strikethrough = True
    Me.lblCash.Caption = ""
    SetPatientInfoWrite True
    If mintEditMode <= 1 Then ResetVsf vsf2
End Sub

Private Function InitDepts() As Boolean
'功能：初始化住院临床科室
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strOldText As String

    On Error GoTo errH
    strOldText = Me.cbo开单科室.Text
    Me.cbo开单科室.Clear

    strSQL = _
        " Select Distinct A.ID,A.编码,A.名称" & _
        " From 部门表 A,部门性质说明 B " & _
        " Where B.部门ID = A.ID " & _
        " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
        " And (B.工作性质 IN('临床','体检','护理'))" & _
        " Order by A.编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)

    For i = 1 To rsTmp.RecordCount
        cbo开单科室.AddItem rsTmp!名称
        cbo开单科室.ItemData(cbo开单科室.NewIndex) = rsTmp!ID
        If strOldText = rsTmp!名称 Then
            cbo开单科室.ListIndex = cbo开单科室.NewIndex
        End If
        rsTmp.MoveNext
    Next

    On Error Resume Next
'    Me.cbo开单科室.Text = strOldText
    If cbo开单科室.ListCount > 0 And Me.cbo开单科室.ListIndex = -1 Then
        cbo开单科室.ListIndex = 0
    End If

    InitDepts = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cboAge_Click()
    If Me.cboAge.Text = "成人" Or Me.cboAge.Text = "婴儿" Then
        Me.txt年龄.Text = ""
    End If
End Sub

Private Sub cboAge_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SetFocusNextIndex Me.cboAge.TabIndex ' zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo开单科室_GotFocus()
    Call zlControl.TxtSelAll(cbo开单科室)
End Sub

Private Sub cbo开单科室_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cbo开单科室_Validate(False)
        SetFocusNextIndex Me.cbo开单科室.TabIndex  ' zlCommFun.PressKey vbKeyTab
        gintSelectFocus = 2
    End If
End Sub

Private Sub cbo开单科室_Validate(Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intIdx As Long, i As Long
    Dim strInput As String
    Dim vRect As RECT, blnCancel As Boolean

    If cbo开单科室.ListIndex <> -1 Then mlngReqDept = Me.cbo开单科室.ItemData(Me.cbo开单科室.ListIndex): Exit Sub '已选中
    If cbo开单科室.Text = "" Then '无输入
        Exit Sub
    End If

    strInput = UCase(NeedName(cbo开单科室.Text))
    '全院临床科室
    strSQL = _
        " Select Distinct A.ID,A.编码,A.名称,A.简码" & _
        " From 部门表 A,部门性质说明 B " & _
        " Where B.部门ID = A.ID " & _
        " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
        " And (B.工作性质 IN('临床','体检'))" & _
        " And (Upper(A.编码) Like [1] Or Upper(A.名称) Like [2] Or Upper(A.简码) Like [2])" & _
        " Order by A.编码"

    On Error GoTo errH
    vRect = GetControlRect(cbo医生.hWnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "开嘱科室", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, cbo开单科室.Height, blnCancel, False, True, UCase(strInput) & "%", UCase(strInput) & "%")
    If Not rsTmp Is Nothing Then
        If Not zlControl.CboLocate(cbo开单科室, rsTmp!名称) Then
            cbo开单科室.Text = ""
        End If
    Else
        If Not blnCancel Then
            MsgBox "未找到对应的科室。", vbInformation, gstrSysName
        End If
        Cancel = True: Exit Sub
    End If
    If Me.cbo开单科室.ListIndex > -1 Then mlngReqDept = Me.cbo开单科室.ItemData(Me.cbo开单科室.ListIndex)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdExt_Click()
    Dim tmpExtData As String
    Dim lngKey As Long
    Dim vRect As RECT, blnCancel As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSampleType As String

    On Error Resume Next
    If mstrExtData = "" Then
        gstrSql = "select 诊疗项目ID from 病人医嘱记录 where 相关id in (" & mstrKeys & ")"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
        If rsTmp.EOF = True Then Exit Sub   '没有时退出
        Do While Not rsTmp.EOF
            tmpExtData = tmpExtData & "," & Nvl(rsTmp("诊疗项目ID"))
            rsTmp.MoveNext
        Loop
        lngKey = Val(Mid(tmpExtData, 2))
    Else
        lngKey = Val(Split(mstrExtData, ";")(0))
    End If

    If lngKey = 0 Then
        gstrSql = "Select 编码 as ID,名称 From 诊疗检验标本 order by 编码 "
    Else
        gstrSql = "   Select Distinct b.编码 as ID,B.名称  " & _
                "   From 诊疗项目目录 A,诊疗检验标本 B,检验项目参考 C,检验报告项目 D" & _
                "   Where A.ID=D.诊疗项目ID(+) And D.报告项目ID=C.项目ID(+)" & _
                        "       And (C.标本类型 Is Null Or C.标本类型=B.名称) And A.ID In (" & lngKey & ") order by b.编码 "

    End If
    vRect = GetControlRect(Me.txt附加.hWnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, gstrSql, 0, "标本类型", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, Me.txt附加.Height, blnCancel, False, True)

    If Not rsTmp Is Nothing Then
        strSampleType = Nvl(rsTmp!名称)
    Else
        If Not blnCancel Then
            MsgBox "未找到标本类型", vbInformation, gstrSysName
        End If
    End If

    If Trim(strSampleType) <> "" Then
        Me.txt医嘱内容 = Replace(Me.txt医嘱内容, "(" & Me.txt附加 & ")", "(" & strSampleType & ")")
        Me.txt附加 = strSampleType
    End If
    Me.txt附加.SetFocus
    Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdOpen_Click()
    Dim intSelect As Integer

    If mstrKeys = "" Then Exit Sub

    intSelect = OpenSelect("", False): gintSelectFocus = 2 'DoEvents
    Select Case intSelect
        Case 1
            '选取了一个项目
            SetAdviceEnable False   '不允许登记
            '产生仪器和标本号
            If mintEditMode = 0 Then
                Call LoadDefaultData
                Call SelectDefault

                vsf2.Col = 2
                vsf2.ShowCell vsf2.Row, vsf2.Col
                vsf2.SetFocus
            Else
                '赋医嘱ID
                Call SelectDefault
                Me.cbo性别.SetFocus
            End If
    End Select
End Sub

Private Sub cmdSel_Click()
    '检验项目
    Dim rsTmp As New ADODB.Recordset
    If mstrExtData = "" And mintEditMode = 3 And Me.txt医嘱内容.Enabled = True Then
        gstrSql = "select distinct 诊疗项目ID from 检验普通结果 where 检验标本id = [1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngSampleID)
        Do While Not rsTmp.EOF
            mstrExtData = mstrExtData & "," & Nvl(rsTmp("诊疗项目id"))
            rsTmp.MoveNext
        Loop
        If mstrExtData <> "" Then
            mstrExtData = Mid(mstrExtData, 2) & ";" & txt附加
        End If
    End If

    If AdviceInput Then
'        DoEvents
        mblnSaveAdvice = True
        gintSelectFocus = 2
        '显示已缺省设置的值
        txt医嘱内容.Tag = txt医嘱内容.Text
        txt附加.Tag = txt附加.Text

        '处理检验仪器、标本号，对于已有标本的操作（重核或补填申请）则只重新赋医嘱ID
        If mintEditMode <= 1 Then
            Call LoadDefaultData
            Call SelectDefault

            With vsf2
                If .Rows > 1 Then
                    .Row = 1
                End If
                .Col = 2
                .ShowCell vsf2.Row, vsf2.Col
                .SetFocus
            End With
        Else
            '赋医嘱ID
            Call SelectDefault
        End If

        Me.txt医嘱内容.SetFocus
        Call zlCommFun.PressKey(vbKeyTab)
    Else
'        DoEvents
        gintSelectFocus = 2
        '恢复原值
        txt医嘱内容.Text = txt医嘱内容.Tag
        txt附加.Text = txt附加.Tag
        zlControl.TxtSelAll txt医嘱内容

        txt医嘱内容.SetFocus
    End If
    gintSelectFocus = 2
End Sub
'
'Private Sub Form_Activate()
'    On Error Resume Next
'    Select Case mintFocusItem
'        Case FocusItem.标本号
'            vsf2.SetFocus
'        Case FocusItem.开单科室
'            Me.cbo开单科室.SetFocus
'        Case FocusItem.姓名
'            Me.txt姓名.SetFocus
'        Case FocusItem.医嘱内容
'            Me.txt医嘱内容.SetFocus
'        Case FocusItem.医生
'            Me.cbo医生.SetFocus
'    End Select
'    mintFocusItem = 0
'End Sub

Private Sub dtp_Change(Index As Integer)
    If Index = 1 Then
        If Abs(DateDiff("d", DTP(Index).Value, zlDatabase.Currentdate)) > 30 Then
            MsgBox "你选择的检验时间和当前时间相差超过了30天，请注意是否正确！", vbQuestion, Me.Caption
        End If
    End If
End Sub

Public Sub IdKindChange()
    If Me.ActiveControl Is txt姓名 Then
       IDKind.IDKind = IIf(IDKind.IDKind = IDKinds.C5就诊卡, 0, IDKind.IDKind + 1)
    End If
End Sub

Private Sub Form_Load()
    Dim blnEmerge As Boolean
    Dim rs As New ADODB.Recordset, i As Long

    '设置参数
    Call SetPara

    mstrPrivs = gstrPrivs

    mintEditMode = -1
    With vsf2
        .Cols = 0
        .NewColumn "", 0, 4
        .NewColumn "检验仪器", 1600, 1, , 0
        .NewColumn "标本号码", 1200, 1, , 1, 15
        .NewColumn "", 0, 1
        .NewColumn "", 0, 1
        .NewColumn "急", IIf(mblnEmerge, 250, 0), 1, , IIf(mblnEmerge, 1, 0), , flexDTBoolean
        .NoDouble = True
        .FixedCols = 0

        .Cell(flexcpBackColor, 0, 0, 0, .Cols - 1) = &HFDD6C6

        .Body.Appearance = flex3DLight
    End With
    '性别
    Set rs = Nothing
    Set rs = GetDictData("性别")
    Set mRsSex = rs
    cbo性别.Clear
    If Not rs Is Nothing Then
        For i = 1 To rs.RecordCount
            cbo性别.AddItem rs!编码 & "-" & rs!名称
            If rs!缺省 = 1 Then
                cbo性别.ItemData(cbo性别.NewIndex) = 1
                cbo性别.ListIndex = cbo性别.NewIndex
            End If
            rs.MoveNext
        Next
    End If

    gstrSql = "Select Distinct D.ID" & vbNewLine & _
            " From 检验小组成员 A, 检验小组 B, 检验小组仪器 C, 检验仪器 D" & vbNewLine & _
            " Where A.小组id = B.ID And B.ID = C.小组id　and 人员id = [1] And C.仪器id = D.ID And C.更改 = 1"

    Set rs = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, UserInfo.ID)
    Do Until rs.EOF
        mstrMachines = mstrMachines & ";" & rs("ID")
        rs.MoveNext
    Loop
    If mstrMachines <> "" Then mstrMachines = mstrMachines & ";"
'    mintFocusItem = FocusItem.姓名
    '-- 2007-10-26 加入一卡通支持
    Set mobjIDCard = New clsIDCard
    Call mobjIDCard.SetParent(Me.hWnd)

    If mobjSquareCard Is Nothing Then
        Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
        If mobjSquareCard.zlInitComponents(Me, glngModul, glngSys, gstrDBUser, gcnOracle, False) = False Then
            MsgBox "IDKind初始化失败!", vbInformation, gstrSysName
        Else
            IDKind.IDKindStr = mobjSquareCard.zlGetIDKindStr(IDKind.IDKindStr)
        End If
    End If

    IDKind.IDKind = 0 'IC卡

    If mobjLisInsideComm Is Nothing Then
        Dim strErr As String
        Set mobjLisInsideComm = CreateObject("zl9LisInsideComm.clsLisInsideComm")
        If Not mobjLisInsideComm Is Nothing Then
            '初始化LIS接口部件
            If mobjLisInsideComm.InitComponentsHIS(glngSys, glngModul, gcnOracle, strErr) = False Then
                If strErr <> "" Then
                    MsgBox "初始化LIS接口失败！" & vbCrLf & strErr
                End If
                Set mobjLisInsideComm = Nothing
            End If
        End If
    End If


End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
End Sub

Private Function GetDictData(strDict As String) As ADODB.Recordset
'功能：从指定的字典中读取数据
'参数：strDict=字典对应的表名
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String

    On Error GoTo errH

    strSQL = "Select 编码,名称,Nvl(缺省标志,0) as 缺省 From " & strDict & " Order by 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)

    If Not rsTmp.EOF Then Set GetDictData = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cbo性别_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If mblnBarCode = True And KeyAscii <> vbKeyReturn Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then
        If mblnBarCode Then
            AutoSave
        Else
            SetFocusNextIndex Me.cbo性别.TabIndex
'            zlCommFun.PressKey vbKeyTab
        End If
        Exit Sub
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    zlcommfun.OpenIme False
    '2007-10-26 加入一卡通支持
    If Not mobjIDCard Is Nothing Then
        Call mobjIDCard.SetEnabled(False)
        Set mobjIDCard = Nothing
    End If
    Set mobjICCard = Nothing
    Set mobjSquareCard = Nothing
    mstrMachines = ""
End Sub

Private Sub IDKind_Click()
    Dim lng卡类别ID As Long, strOutCardNO As String, strExpand As String, strOutPatiInforXML As String
' 2007-10-26 增加一卡通支持
    Dim blnCancle As Boolean
    IDKind.Tag = ""
    If Not txt姓名.Locked And txt姓名.Text = "" And txt姓名.Tag = "" Then
        If IDKind.IDKind = IDKinds.C3IC卡号 Then
            If mobjICCard Is Nothing Then
                Set mobjICCard = CreateObject("zlICCard.clsICCard")
                Set mobjICCard.gcnOracle = gcnOracle
            End If
            If Not mobjICCard Is Nothing Then
                txt姓名.Text = mobjICCard.Read_Card()
                If txt姓名.Text <> "" Then
                    IDKind.Tag = "医保卡"

'                    Call txt姓名_Validate(blnCancle)
                    Call Txt姓名Exec
                    Me.txt姓名.SetFocus
                    gintSelectFocus = 2
                End If
            End If
        End If
    End If
    lng卡类别ID = Val(IDKind.GetKindItem("卡类别ID"))
    If lng卡类别ID = 0 Then Exit Sub

    If mobjSquareCard.zlReadCard(Me, glngModul, lng卡类别ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txt姓名.Text = strOutCardNO
    If txt姓名.Text <> "" Then Call txt姓名_KeyPress(vbKeyReturn)
End Sub

Private Sub IDKind_ItemClick(Index As Integer)
    mblnShowPwd = Trim(IDKind.GetKindItem(7)) <> ""
    Me.txt姓名 = ""
    If mblnShowPwd = True Then
        Me.txt姓名.PasswordChar = "*"
    Else
        Me.txt姓名.PasswordChar = ""
    End If
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
' 2007-10-26 增加一卡通支持
    Dim lngPreIDKind As Long
    Dim blnCancle As Boolean
    IDKind.Tag = ""
    If Not txt姓名.Locked And txt姓名.Text = "" And txt姓名.Tag = "" And Me.ActiveControl Is txt姓名 Then
        lngPreIDKind = IDKind.IDKind
        IDKind.IDKind = IDKinds.C2身份证号
        txt姓名.Text = strID
        IDKind.Tag = "身份证"
'        Call txt姓名_Validate(blnCancle)
        Call Txt姓名Exec
        IDKind.IDKind = lngPreIDKind
        gintSelectFocus = 2
    End If
End Sub

Private Sub txtBed_GotFocus()
    zlControl.TxtSelAll txtBed
End Sub

Private Sub txtBed_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
'        zlCommFun.PressKey vbKeyTab
        SetFocusNextIndex Me.txtBed.TabIndex
    End If
End Sub

Private Sub txtID_GotFocus()
    zlControl.TxtSelAll txtID
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
'        zlCommFun.PressKey vbKeyTab
        SetFocusNextIndex Me.txtID.TabIndex
    Else
        KeyAscii = FilterKeyAscii(KeyAscii, 99, "0123456789")
    End If
End Sub

Private Sub txtPatientDept_GotFocus()
    zlControl.TxtSelAll txtPatientDept
End Sub

Private Sub txtPatientDept_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Len(Trim(Me.txt医嘱内容)) > 0 Then
            vsf2.Col = 2
            vsf2.ShowCell vsf2.Row, vsf2.Col
            vsf2.SetFocus
        Else
'            zlCommFun.PressKey vbKeyTab
            SetFocusNextIndex Me.txtPatientDept.TabIndex
        End If
        Exit Sub
    End If
End Sub

Private Sub txt附加_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SetFocusNextIndex Me.txt附加.TabIndex ' zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txt年龄_GotFocus()
    zlControl.TxtSelAll txt年龄
End Sub

Private Sub txt年龄_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Len(Trim(Me.txt医嘱内容)) > 0 And Me.txtID.Enabled = False Then
            With vsf2
                If .Rows > 1 Then
                    .Row = 1
                End If
                .Col = 2
                .ShowCell vsf2.Row, vsf2.Col
                .SetFocus
            End With
        Else
'            zlCommFun.PressKey vbKeyTab
            SetFocusNextIndex Me.txt年龄.TabIndex
        End If
        Exit Sub
    Else
        KeyAscii = FilterKeyAscii(KeyAscii, 99, "0123456789.*")
    End If
End Sub

Private Sub cbo开单科室_Click()
    If cbo开单科室.ListIndex > -1 Then InitDoctors cbo开单科室.ItemData(cbo开单科室.ListIndex)
End Sub

Private Sub cbo医生_GotFocus()
    Call zlControl.TxtSelAll(cbo医生)
End Sub

Private Sub cbo医生_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
'        If mbln微生物项目 = True Then
'            Call cbo医生_Validate(False)
'            dtp(0).SetFocus
'        Else
'            zlCommFun.PressKey vbKeyTab
            Call cbo医生_Validate(False)
            SetFocusNextIndex Me.cbo医生.TabIndex + 2
            gintSelectFocus = 2
'        End If
    End If
End Sub

Private Sub cbo医生_Validate(Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intIdx As Long, i As Long
    Dim strInput As String
    Dim vRect As RECT, blnCancel As Boolean
    Dim lngDept As Long

    If cbo医生.ListIndex <> -1 Then mstrReqDoctor = Me.cbo医生.Text: Exit Sub '已选中
    If cbo医生.Text = "" Then '无输入
        Exit Sub
    End If

    lngDept = cbo开单科室.ItemData(cbo开单科室.ListIndex)

    strInput = UCase(NeedName(cbo医生.Text))
    '全院医生
    strSQL = "Select Distinct 部门ID From 部门性质说明 Where 服务对象 IN(1,2,3)"
    strSQL = "Select Distinct A.ID,A.编号,A.姓名,A.简码" & _
        " From 人员表 A,部门人员 B,人员性质说明 C" & _
        " Where A.ID=B.人员ID And A.ID=C.人员ID And C.人员性质='医生'" & _
        " And B.部门ID IN(" & strSQL & ")" & IIf(lngDept > 0, " and b.部门ID=[3] ", "") & _
        " And (Upper(A.编号) Like [1] Or Upper(A.姓名) Like [2] Or Upper(A.简码) Like [2])" & _
        " And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) " & _
        " Order by A.简码"


    On Error GoTo errH
    vRect = GetControlRect(cbo医生.hWnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "开嘱医生", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, cbo医生.Height, blnCancel, False, True, strInput & "%", strInput & "%", lngDept)
    If Not rsTmp Is Nothing Then
        cbo医生.Text = rsTmp!姓名
'        Me.dtp(0).SetFocus
'        SetFocusNextIndex Me.cbo医生.TabIndex


    Else
        If Not blnCancel Then
            MsgBox "未找到对应的医生。", vbInformation, gstrSysName
        End If
        Cancel = True: gintSelectFocus = 2: Exit Sub
    End If
    If Len(Trim(Me.cbo医生.Text)) > 0 Then mstrReqDoctor = Me.cbo医生.Text
    gintSelectFocus = 2
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim lngloop As Long

    If KeyAscii = vbKeyReturn Then
        If mblnBarCode And Index = 1 Then AutoSave: Exit Sub

        For lngloop = 0 To cbo(Index).ListCount - 1
            If InStr(cbo(Index).List(lngloop), "-") > 0 Then
                If Mid(cbo(Index).List(lngloop), 1, InStr(cbo(Index).List(lngloop), "-") - 1) = cbo(Index).Text Then
                    cbo(Index).Text = cbo(Index).List(lngloop)
                    Exit For
                End If
            End If
        Next
'        zlCommFun.PressKey vbKeyTab
        SetFocusNextIndex Me.cbo(Index).TabIndex
    End If
End Sub

Private Sub cbo_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case 0
            Cancel = Not StrIsValid(cbo(Index).Text, 50)
        Case 1, 2
            Cancel = Not StrIsValid(cbo(Index).Text, 50)
    End Select
End Sub

Private Sub dtp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
'        zlCommFun.PressKey vbKeyTab
        SetFocusNextIndex Me.DTP(Index).TabIndex
    End If
End Sub

Private Sub InitDoctors(ByVal lng科室ID As Long)
'功能：读取当前开单科室中包含的所有人员
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strOldDoctor As String

    strOldDoctor = Me.cbo医生.Text
    Me.cbo医生.Clear

    '科室医生或护士
    strSQL = _
        "Select Distinct A.ID,B.部门ID,A.编号,A.姓名,Upper(A.简码) as 简码," & _
        " C.人员性质,Nvl(A.聘任技术职务,0) as 职务" & _
        " From 人员表 A,部门人员 B,人员性质说明 C" & _
        " Where A.ID=B.人员ID And A.ID=C.人员ID" & _
        " And C.人员性质 IN('医生') And B.部门ID=[1] " & _
        " And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) "

    strSQL = strSQL & " Order by 简码,人员性质 Desc"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng科室ID)

    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cbo医生.AddItem rsTmp!姓名
            cbo医生.ItemData(cbo医生.ListCount - 1) = rsTmp!部门ID
            If rsTmp!姓名 = strOldDoctor Then
                cbo医生.ListIndex = cbo医生.NewIndex
            End If

            If rsTmp!ID = UserInfo.ID And cbo医生.ListIndex = -1 Then cbo医生.ListIndex = cbo医生.NewIndex
            rsTmp.MoveNext
        Next

        If cbo医生.ListCount = 1 And cbo医生.ListIndex = -1 Then cbo医生.ListIndex = 0
    End If
End Sub

Private Sub txt姓名_GotFocus()
    txt姓名.SelStart = 0
    txt姓名.SelLength = Len(txt姓名.Text)
    If IDKind.IDKind = IDKinds.C0姓名 Then
        If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
    ElseIf IDKind.IDKind = IDKinds.C3IC卡号 Then
        If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (False)
    End If
    If Val(zlDatabase.GetPara("使用个性化风格")) <> 0 Then
        zlCommFun.OpenIme True
    End If
End Sub

Private Sub txt姓名_KeyDown(KeyCode As Integer, Shift As Integer)
'    Dim blnCard As Boolean
'    '就诊卡
'    blnCard = False
'    If IDKind.IDKind = IDKinds.C5就诊卡 Then
'        If KeyCode <> 8 And Len(txt姓名.Text) = gbytCardNOLen - 1 And txt姓名.SelLength <> Len(txt姓名.Text) Then
'            txt姓名.Text = txt姓名.Text & UCase(Chr(KeyCode))
'            blnCard = True
'            KeyCode = 0
'        End If
'    End If
'
'
''    SetActiveWindow Me.Hwnd
'    If KeyCode <> vbKeyReturn And blnCard = False Then
'        KeyCode = Asc(UCase(Chr(KeyCode)))
'
'    Else
'        KeyCode = 0
''        zlCommFun.PressKey vbKeyTab
'
'        Me.txt姓名.Enabled = False
''        Call txt姓名_Validate(False)
'        Call Txt姓名Exec
'        Debug.Print Me.txt姓名
'        If Me.txt姓名.Text <> "" Then
'            SetFocusNextIndex txt姓名.TabIndex
'        Else
'            Me.txt姓名.Enabled = True
'            If mstr条码仪器 <> "" And frmLabMain.mlngMachineID > 0 Then Me.txt姓名.SetFocus
''            Me.txt姓名.SetFocus
'            If IDKind.IDKind = IDKinds.C5就诊卡 Then
'                Me.txt姓名.Text = ""
'            End If
'        End If
'        gintSelectFocus = 2
'        Me.txt姓名.Enabled = True
'
'        mblnCard = False
'    End If
End Sub

Private Sub txt姓名_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean

    If CheckIsInclude(UCase(Chr(KeyAscii)), "'‘’;；:：?？|,，。""") = True Then KeyAscii = 0

    blnCard = False
    If IDKind.IDKind = IDKinds.C5就诊卡 Then
        gbytCardNOLen = Val(IDKind.GetKindItem("卡号长度", IDKinds.C5就诊卡))
        If KeyAscii <> 8 And Len(txt姓名.Text) = gbytCardNOLen - 1 And txt姓名.SelLength <> Len(txt姓名.Text) Then
            If KeyAscii <> 13 Then
                txt姓名.Text = txt姓名.Text & UCase(Chr(KeyAscii))
            End If
            blnCard = True
            KeyAscii = 0
        End If
    End If

    If IDKind.IDKind = IDKinds.C5就诊卡 Then
'        mblnCard = zlCommFun.InputIsCard(txt姓名, KeyAscii, True)
    End If

    If KeyAscii = vbKeyReturn Or blnCard = True Then

        KeyAscii = 0
'        zlCommFun.PressKey vbKeyTab

        Me.txt姓名.Enabled = False
'        Call txt姓名_Validate(False)
        Call Txt姓名Exec

        If Me.txt姓名.Text <> "" Then
            SetFocusNextIndex txt姓名.TabIndex
        Else
            Me.txt姓名.Enabled = True
            If mstr条码仪器 <> "" And frmLabMain.mlngMachineID > 0 Then Me.txt姓名.SetFocus
'            Me.txt姓名.SetFocus
            If IDKind.IDKind = IDKinds.C5就诊卡 Then
                Me.txt姓名.Text = ""
            End If
        End If
'        Debug.Print Me.txt姓名
        gintSelectFocus = 2
        Me.txt姓名.Enabled = True

        mblnCard = False
    End If

End Sub

Private Sub txt姓名_LostFocus()
    txt姓名.SelStart = 0
    txt姓名.SelLength = Len(txt姓名.Text)
    If txt姓名.Text = "" And Not txt姓名.Locked Then
        If IDKind.IDKind = IDKinds.C0姓名 Then
            If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (True)
        ElseIf IDKind.IDKind = IDKinds.C3IC卡号 Then
            If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (True)
        End If
    End If

End Sub

Private Sub txt姓名_Validate(Cancel As Boolean)
'    Dim strInput As String
'    Dim rsTmp As New ADODB.Recordset, i As Integer
'    Dim strField As String
'    Dim strBarCode As String
'    Dim rsDept As ADODB.Recordset, strsql As String
'    Dim intSelect As Integer
'    Dim intPatientSource As Integer                     '病人来源
'    Dim rsPaInfo As New ADODB.Recordset                 '提取标本记录中的"标识号"
'    Dim blnGetPaInfo As Boolean
'
'    On Error GoTo errH
'    If Len(Trim(txt姓名)) = 0 Or Me.txt姓名.Enabled = True Or Me.txt姓名 = Me.txt姓名.Tag Then Exit Sub
'
'    If txt姓名 <> txt姓名.Tag Then mlng病人id = 0
'    If mlng病人id > 0 Then
'        strsql = "select 病人来源 from 检验标本记录 where id = [1] "
'        Set rsTmp = zlDatabase.OpenSQLRecord(strsql, gstrSysName, mlngSampleID)
'        If rsTmp.RecordCount > 0 Then
'            intPatientSource = Nvl(rsTmp("病人来源"), PatientType)
'        Else
'            intPatientSource = PatientType
'        End If
'
'        If txt姓名 = txt姓名.Tag And intPatientSource <> 3 Then Exit Sub
'    Else
'        If txt姓名 = txt姓名.Tag Then Exit Sub
'    End If
'
'    mblnSaveAdvice = True
'    Cancel = Not StrIsValid(txt姓名.Text, txt姓名.MaxLength)
'
'    '初始病人信息
'    '2007-10-26 加入一卡通支付
'    If IDKind.Tag = "医保卡" Then
'        Set rsTmp = GetPatient("卡" & txt姓名)
'        IDKind.Tag = ""
'    ElseIf IDKind.Tag = "身份证" Then
'        Set rsTmp = GetPatient("证" & txt姓名)
'        IDKind.Tag = ""
'    Else
'        Set rsTmp = GetPatient(txt姓名)
'    End If
'
'    If rsTmp.EOF = True Then
'        Set rsTmp = GetPatientInfo(txt姓名)
'        blnGetPaInfo = True
'    End If
'
'    strBarCode = txt姓名
'    If rsTmp.RecordCount <= 0 Then
'        mlng病人id = 0
'        '登记新病人
'        mstrKeys = ""
'        Me.txt年龄 = "": Me.cboAge.ListIndex = 0
'        Me.txtPatientDept = "": Me.txtPatientDept.Tag = 0
'        Me.txtID = "": Me.txtBed = ""
'        '如果想输入院内病人，则不允许继续
'        If InStr("+-*./", Left(Me.txt姓名.Text, 1)) > 0 Or mblnBarCode Then
'            Me.txt姓名.Text = "": Cancel = True
'            Exit Sub
'        End If
'        If mblnBarCode = True Then
'            MsgBox "没有找到条码，请查看条码是否已取消绑定！", vbInformation, Me.Caption
'            Me.txt姓名.Text = "": Cancel = True
'            Exit Sub
'        End If
'        PatientType = 1
'        '处理登记的默认科室、医生
'        If mlngReqDept > 0 Then
'            cbo开单科室.ListIndex = FindComboItem(cbo开单科室, mlngReqDept)
'            Me.cbo医生.Text = mstrReqDoctor
'        End If
'        SetPatientInfoWrite False
'    Else
'        On Error Resume Next
'        Me.txt姓名.Text = Nvl(rsTmp("姓名"))
''        Me.txt年龄 = IIf(IsNull(rsTmp("年龄")), "", Val(rsTmp("年龄"))): If Me.txt年龄 = "0" Then Me.txt年龄 = ""
''        Me.cboAge.Text = IIf(IsNull(rsTmp("年龄")), "岁", Replace(rsTmp("年龄"), Val(rsTmp("年龄")), ""))
'        Me.txt年龄 = IIf(IsNull(rsTmp("年龄")), "", IIf(IsNumeric(rsTmp("年龄")), Val(rsTmp("年龄")), Mid(rsTmp("年龄"), 1, Len(rsTmp("年龄")) - 1))): If Me.txt年龄 = "0" Then Me.txt年龄 = ""
'        Me.cboAge.Text = IIf(IsNull(rsTmp("年龄")), "岁", Right(rsTmp("年龄"), 1))
'        If cboAge.ListIndex = -1 Then cboAge.ListIndex = 0
'        Me.cbo性别 = Nvl(rsTmp("性别")) ' CombIndex(cbo性别, Nvl(rsTmp("性别")))
'        mlng病人id = Nvl(rsTmp("病人ID"), 0): PatientType = Nvl(rsTmp("PatientType"), 1)
'
'        '设置默认开单科室、医生
'        cbo开单科室.ListIndex = FindComboItem(cbo开单科室, Nvl(rsTmp("病人科室"), 0))
''        DoEvents
'        gintSelectFocus = 2
'        strField = ""
'        strField = rsTmp.Fields("医生").Name
'        If strField = "医生" Then
'            Me.cbo医生.Text = Nvl(rsTmp("医生"))
'            For i = 0 To Me.cbo医生.ListCount - 1
'                If Me.cbo医生.List(i) Like Nvl(rsTmp("医生")) Then
'                    Me.cbo医生.ListIndex = i
'                    Exit For
'                End If
'            Next
'        End If
'        '显示病人科室
'        If IsNumeric(rsTmp("病人科室")) = True Then
'            strsql = "Select 名称 From 部门表 Where ID=[1]"
'            Set rsDept = zlDatabase.OpenSQLRecord(strsql, Me.Caption, CLng(Nvl(rsTmp("病人科室"), 0)))
'            If rsDept.EOF Then
'                Me.txtPatientDept = "": Me.txtPatientDept.Tag = 0
'            Else
'                Me.txtPatientDept.Text = rsDept("名称"): Me.txtPatientDept.Tag = Nvl(rsTmp("病人科室"), 0)
'            End If
'        Else
'            Me.txtPatientDept = Nvl(rsTmp("病人科室"))
'        End If
'
'        Me.txtID = Nvl(rsTmp("住院号")): If Len(Me.txtID) = 0 Then Me.txtID = Nvl(rsTmp("门诊号"))
'
'        Me.txtBed = Nvl(rsTmp("当前床号"))
'
'        '当检验标本记录里有时优先显示检验标本记录里的信息
'        If Trim(Me.txtID) = "" Or Trim(Me.txtBed) = "" Or Trim(Me.txtPatientDept) = "" Then
'            If blnGetPaInfo = True Then
'                Me.txtID = Nvl(rsTmp("标识号"), Me.txtID)
'                Me.txtBed = Nvl(rsTmp("床号"), Me.txtBed)
'                Me.txtPatientDept = Nvl(rsTmp("病人科室"), Me.txtPatientDept)
'            End If
'        End If
'
'        '处理登记的默认科室、医生
'        If Me.cbo开单科室.ListIndex = -1 And mlngReqDept > 0 Then
'            cbo开单科室.ListIndex = FindComboItem(cbo开单科室, mlngReqDept)
'            Me.cbo医生.Text = mstrReqDoctor
'        End If
'    End If
'    '核收时选择检验申请
'    If mlng病人id > 0 And Not mintEditMode = 1 And (intPatientSource <> 3 Or mblnBarCode = True) Then
'        intSelect = OpenSelect(strBarCode, True)
''        DoEvents
'        gintSelectFocus = 2
'        Select Case intSelect
'            Case 0
'                '没有匹配的项目
'                mstrKeys = ""
'                If mlng病人id = 0 Or mblnBarCode Then
''                    mintFocusItem = FocusItem.姓名
'                    MsgBox "条码已被核收！", vbInformation, Me.Caption
'                    mlng病人id = 0
'                    txt姓名.Text = ""
'                    If Me.txt姓名.Enabled = True Then
'                        txt姓名.SetFocus
'                    End If
'                    Cancel = True
'                Else
'                    '允许登记
'                    SetAdviceEnable True
'                    If mlngDefaultItemID > 0 Then
'                        strsql = "select 标本部位 from 诊疗项目目录 where id = [1] "
'                        Set rsTmp = zlDatabase.OpenSQLRecord(strsql, gstrSysName, mlngDefaultItemID)
'                        AdviceSet检查手术 3, mlngDefaultItemID & ";" & rsTmp("标本部位")
'                        mstrExtData = mlngDefaultItemID & ";" & rsTmp("标本部位")
'                        '获取采集方式
'                        Set rsTmp = SelectCap(Split(Split(mstrExtData, ";")(0), ",")(0))
'                        If rsTmp Is Nothing Then
'                            MsgBox "没有定义标本采集方式，请到诊疗项目管理中设置。", vbInformation, gstrSysName
'                            Exit Sub
'                        End If
'                        mlngCapID = rsTmp("ID")
'                    End If
'                    If rsRelativeAdvice Is Nothing Then
'                        Me.txt医嘱内容 = ""
'                        Me.txt附加 = ""
'                        Me.txt医嘱内容.Tag = "": Me.txt附加.Tag = "": mlngCapID = 0
'                    Else
'                        txt医嘱内容.Text = Get检查手术名称(2, "")
'                        txt医嘱内容.Text = txt医嘱内容.Text & "(" & Split(mstrExtData, ";")(1) & ")"
'                        Me.txt附加 = Split(mstrExtData, ";")(1)
'                        If mintEditMode = 0 Then
'                            Call LoadDefaultData
'                            Call SelectDefault
''                            mintFocusItem = FocusItem.标本号
'
''                            DoEvents
'                            vsf2.Col = 2
'                            vsf2.ShowCell vsf2.Row, vsf2.Col
'                            vsf2.SetFocus
'                            gintSelectFocus = 2
'                        Else
'                            '赋医嘱ID
'                            Call SelectDefault
'                        End If
'                    End If
'                End If
'            Case 1
'                '选取了一个项目
'                SetAdviceEnable False   '不允许登记
'                '产生仪器和标本号
'                If mintEditMode = 0 Then
'                    Call LoadDefaultData
'                    Call SelectDefault
''                    mintFocusItem = FocusItem.标本号
'
'                    vsf2.Col = 2
'                    vsf2.ShowCell vsf2.Row, vsf2.Col
'                    vsf2.SetFocus
'                Else
'                    '赋医嘱ID
'                    Call SelectDefault
'                    vsf2.Col = 2
'                    vsf2.ShowCell vsf2.Row, vsf2.Col
'                    vsf2.SetFocus
'                End If
'                '条码自动保存
'                If mblnBarCode Then Me.cbo性别.SetFocus
'            Case 2
'                '取消了本次选择
''                mintFocusItem = FocusItem.姓名
'
'                mlng病人id = 0
'                mstrKeys = ""
'                txt姓名.Text = ""
'                If Me.txt姓名.Enabled = True Then
'                    txt姓名.SetFocus
'                End If
'                Cancel = True
'        End Select
'    Else
'        SetAdviceEnable True
'        If mlngDefaultItemID > 0 Then
'            strsql = "select 标本部位 from 诊疗项目目录 where id = [1] "
'            Set rsTmp = zlDatabase.OpenSQLRecord(strsql, gstrSysName, mlngDefaultItemID)
'            AdviceSet检查手术 3, mlngDefaultItemID & ";" & rsTmp("标本部位")
'            mstrExtData = mlngDefaultItemID & ";" & rsTmp("标本部位")
'            '获取采集方式
'            Set rsTmp = SelectCap(Split(Split(mstrExtData, ";")(0), ",")(0))
'            If rsTmp Is Nothing Then
'                MsgBox "没有定义标本采集方式，请到诊疗项目管理中设置。", vbInformation, gstrSysName
'                Exit Sub
'            End If
'            mlngCapID = rsTmp("ID")
'        End If
'        If rsRelativeAdvice Is Nothing Then
'            Me.txt医嘱内容 = ""
'            Me.txt附加 = ""
'            Me.txt医嘱内容.Tag = "": Me.txt附加.Tag = "": mlngCapID = 0
'        Else
'            txt医嘱内容.Text = Get检查手术名称(2, "")
'            txt医嘱内容.Text = txt医嘱内容.Text & "(" & Split(mstrExtData, ";")(1) & ")"
'            Me.txt附加 = Split(mstrExtData, ";")(1)
'            If mintEditMode <= 1 Then
'                Call LoadDefaultData
'                Call SelectDefault
'
'                vsf2.Col = 2
'                vsf2.ShowCell vsf2.Row, vsf2.Col
'                vsf2.SetFocus
'            Else
'                '赋医嘱ID
'                Call SelectDefault
'            End If
'        End If
'    End If
'
'    If mlng病人id > 0 Then
'        txt姓名.Tag = txt姓名.Text
'    End If
'    Exit Sub
'errH:
'    Call InitEdit
'    If ErrCenter() = 1 Then
'        Resume
'    End If
'    Call SaveErrLog
End Sub
Private Function GetPatient(strCode As String) As ADODB.Recordset
'功能：读取病人信息，并显示该病人存在的医嘱时间
    Dim strSQL As String, i As Long
    Dim strNO As String, str姓名 As String, lng病人ID As Long
    Dim strSeek As String
    Dim objPoint As POINTAPI, lng挂号效期 As Long
    Dim strTmp As String, rsTmp As New Recordset, str挂号单 As String
    Dim str小组 As String, str仪器 As String
    Dim strSQLbak As String
    Dim lng卡类别ID As Long


    On Error GoTo errH

    str小组 = frmLabMain.mstrMachineGroup
    If str小组 <> "所有小组" Then
        str小组 = Mid(str小组, 1, InStr(str小组, "-") - 1)
    End If
    str仪器 = frmLabMain.mlngMachineID
    mstr条码仪器 = ""

'    strsql = "Select 参数值 From 系统参数表 Where 参数号=21"
'    Set rsTmp = zldatabase.OpenSQLRecord(strsql, Me.Caption)
'    If rsTmp.RecordCount > 0 Then
'        lng挂号效期 = Val(0 + rsTmp.Fields("参数值"))
'    End If
    lng挂号效期 = Val(zlDatabase.GetPara(21, glngSys))

    If lng挂号效期 = 0 Then lng挂号效期 = 7 '未设则为最近2天

    If IsNumeric(strCode) And Len(strCode) >= 12 And InStr("*-+./", Mid(strCode, 1, 1)) = 0 Then
        '预置条码单独处理
        mblnBarCode = True
        strSQL = "Select Decode(A.当前科室id,Null,1,2) As PatientType,B.主页ID,B.病人科室id As 病人科室,B.开嘱医生 As 医生," & _
            "a.姓名,decode(d.名称,null,a.性别,d.编码 || '-' || a.性别) as 性别,a.年龄,a.病人id,a.住院号,a.门诊号,a.当前床号,Zl_Age_Calc(A.病人ID) as 年龄1  " & _
            " From 病人信息 A,病人医嘱记录 B,病人医嘱发送 C , 性别 d Where A.病人ID=B.病人ID+0 And B.ID=C.医嘱ID+0 and a.性别 = d.名称(+) " & _
            " And C.样本条码=[1] order by b.开嘱时间 desc  "
        Set GetPatient = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strCode)
'        If GetPatient.EOF = True Then
'            MsgBox "没有找到条码号为<" & strCode & ">的条码!", vbInformation, Me.Caption
'        End If
'        Exit Function
        If GetPatient.EOF = False Then
            Exit Function
        End If
    End If
    mblnBarCode = False
    mblnPrice = False

    If str小组 = "所有小组" Then
        strSQL = "Select distinct 仪器ID,条码输入 From 检验小组 A, 检验小组仪器 B, 检验小组成员 C Where A.ID = B.小组id And A.ID = C.小组id  and c.人员id = [1] and 条码输入 =1" & _
        IIf(str仪器 = 0, "", " and  b.仪器id = [3] ")
    Else
        strSQL = "Select distinct 仪器ID,条码输入 From 检验小组 A, 检验小组仪器 B, 检验小组成员 C Where A.ID = B.小组id And A.ID = C.小组id  and A.编码 = [2] and 条码输入 = 1 " & _
        IIf(str仪器 = 0, "", " and  b.仪器id = [3] ")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID, str小组, Val(str仪器))
    Set GetPatient = rsTmp
    If str仪器 > 0 And rsTmp.RecordCount = 1 Then
        mstr条码仪器 = rsTmp("仪器ID")
        Me.txt姓名.Text = "": Me.txt姓名.Tag = ""
        MsgBox "你必须使用条码输入！如有疑问请于管理员联系！", vbInformation, Me.Caption
        Exit Function
    End If

    '这里记录选择了所有仪器时记录只能按条码输入的仪器
    Do While Not rsTmp.EOF
        mstr条码仪器 = mstr条码仪器 & "," & rsTmp("仪器ID")
        GetPatient.MoveNext
    Loop

    strSeek = strCode
    '判断当前输入模式
    If IsNumeric(strCode) And IsNumeric(Left(strCode, 1)) And Val(IDKind.GetKindItem("卡类别ID")) = 0 Then    '刷卡
        If mobjSquareCard.zlGetPatiID(IDKind.GetKindItem("全名"), strCode, False, lng病人ID) = False Then lng病人ID = 0
        If lng病人ID = 0 Then
            iInputType = 0
            strSeek = strCode
        Else
            iInputType = 1
            strSeek = lng病人ID
        End If
    ElseIf (Left(strCode, 1) = "A" Or Left(strCode, 1) = "-") And IsNumeric(Mid(strCode, 2)) Then '病人ID
        iInputType = 1
        strSeek = Mid(strCode, 2)
    ElseIf (Left(strCode, 1) = "B" Or Left(strCode, 1) = "+") And IsNumeric(Mid(strCode, 2)) Then '住院号
        iInputType = 2
        strSeek = Mid(strCode, 2)
    ElseIf (Left(strCode, 1) = "D" Or Left(strCode, 1) = "*") And IsNumeric(Mid(strCode, 2)) Then '门诊号
        iInputType = 3
        strSeek = Mid(strCode, 2)
    ElseIf Left(strCode, 1) = "G" Or Left(strCode, 1) = "." Then '挂号单
        iInputType = 4
        strSeek = Mid(strCode, 2)
    ElseIf Left(strCode, 1) = "/" Then '收费单据号
        iInputType = 5
        strSeek = Mid(strCode, 2)
        mblnPrice = True
    ElseIf mblnCard Or IDKind.IDKind = IDKinds.C5就诊卡 Then
         iInputType = 7
        strSeek = UCase(strCode)
    ElseIf Not IsNumeric(Mid(strCode, 2)) And Val(IDKind.GetKindItem("卡类别ID")) = 0 Then '当作姓名
        iInputType = 6
        strSeek = Replace(strCode, "(婴儿)", "")
    ElseIf strCode Like "卡*" Then  '医保卡
        strCode = Replace(UCase(strCode), "卡", "")
        If mobjSquareCard.zlGetPatiID(IDKind.GetKindItem("全名"), strCode, False, lng病人ID) = False Then lng病人ID = 0
        If lng病人ID = 0 Then
            iInputType = 8
            strSeek = Replace(UCase(strCode), "卡", "")
        Else
            iInputType = 1
            strSeek = lng病人ID
        End If

    ElseIf strCode Like "证*" Then  '身份证
        strCode = Replace(UCase(strCode), "证", "")
        If mobjSquareCard.zlGetPatiID(IDKind.GetKindItem("全名"), strCode, False, lng病人ID) = False Then lng病人ID = 0
        If lng病人ID = 0 Then
            iInputType = 9
            strSeek = Replace(UCase(strCode), "证", "")
        Else
            iInputType = 1
            strSeek = lng病人ID
        End If

    Else
        If Val(IDKind.GetKindItem("卡类别ID")) <> 0 Then
            lng卡类别ID = Val(IDKind.GetKindItem("卡类别ID"))
            If mobjSquareCard.zlGetPatiID(lng卡类别ID, strCode, False, lng病人ID) = False Then lng病人ID = 0
            If lng病人ID = 0 Then lng病人ID = 0
        Else
            If mobjSquareCard.zlGetPatiID(IDKind.GetKindItem("全名"), strCode, False, lng病人ID) = False Then lng病人ID = 0
        End If
        iInputType = 1
        strSeek = lng病人ID
    End If
    mblnCard = False
    If iInputType = 0 Then '刷卡
        strSQL = _
            "Select Distinct Decode(A.当前科室id, Null, 1, 2) As Patienttype, A.主页id," & vbNewLine & _
            "                Decode(A.当前科室id, Null, Nvl(B.执行部门id, 0), A.当前科室id) As 病人科室, B.执行人 As 医生, A.姓名," & vbNewLine & _
            "                Decode(C.名称, Null, A.性别, C.编码 || '-' || A.性别) As 性别, A.年龄, A.病人id, A.住院号, A.门诊号, A.当前床号,Zl_Age_Calc(A.病人ID) as 年龄1 " & vbNewLine & _
            "From (Select " & gConst_病人信息_列名 & " From 病人信息 a Where 就诊卡号 = [1]" & vbNewLine & _
            "       Union " & vbNewLine & _
            "       Select " & gConst_病人信息_列名 & " From 病人信息 a Where 门诊号 = [2]" & vbNewLine & _
            "       Union " & vbNewLine & _
            "       Select " & gConst_病人信息_列名 & " From 病人信息 a Where 住院号 = [2]) A, 病人挂号记录 B, 性别 C" & vbNewLine & _
            "Where A.病人id = B.病人id(+) and (b.病人id is null or (b.记录状态 =1 and b.记录性质 =1)) And A.门诊号 = B.门诊号(+) And 0 + B.登记时间(+) > Sysdate - [5] And A.性别 = C.名称(+)"

'        strsql = "Select Distinct Decode(a.当前科室id, Null, 1, 2) As Patienttype, Nvl(a.住院次数, 0) As 主页id," & vbNewLine & _
'                "                               Decode(a.当前科室id, Null, Nvl(b.执行部门id, 0), a.当前科室id) As 病人科室, b.执行人 As 医生, a.姓名," & vbNewLine & _
'                "                               Decode(c.名称, Null, a.性别, c.编码 || '-' || a.性别) As 性别, a.年龄, a.病人id, a.住院号, a.门诊号," & vbNewLine & _
'                "                               a.当前床号" & vbNewLine & _
'                "From 病人信息 a, (Select 执行部门id, 执行人, 病人id, 门诊号 From 病人挂号记录 Where 登记时间 > Sysdate - [5]) b, 性别 c" & vbNewLine & _
'                "Where (a.就诊卡号 = [1] Or a.门诊号 =[2] Or a.住院号=[2]) And a.病人id = b.病人id(+) And a.门诊号 = b.门诊号(+) And a.性别 = c.名称(+)"


'            " And (A.当前科室id IS NOT NULL Or NVL(B.执行状态,1) IN (0,2))"
    ElseIf iInputType = 1 Then '病人ID
        strSQL = "Select Decode(A.当前科室id,Null,1,2) As PatientType,A.主页ID,Nvl(A.当前科室id,0) As 病人科室," & _
            "a.姓名,decode(b.名称,null,a.性别,b.编码 || '-' || a.性别) as 性别,a.年龄,a.病人id,a.住院号,a.门诊号,a.当前床号,'' as 医生,Zl_Age_Calc(A.病人ID) as 年龄1  " & _
            " From 病人信息 A , 性别 B Where A.病人ID=[2] And A.性别 = B.名称(+) "
    ElseIf iInputType = 2 Then '住院号
        strSQL = "Select Decode(A.当前科室id,Null,1,2) As PatientType,A.主页ID,Decode(A.当前科室id,Null,Nvl(B.入院科室ID,0),A.当前科室id) As 病人科室,B.住院医师 As 医生," & _
            "a.姓名,decode(c.名称,null,a.性别,c.编码 || '-' || a.性别) as 性别,a.年龄,a.病人id,a.住院号,a.门诊号,a.当前床号,Zl_Age_Calc(A.病人ID) as 年龄1  " & _
            " From 病人信息 A,病案主页 B,性别 C Where A.住院号=[2] And A.主页ID=B.主页ID And A.病人ID=B.病人ID And a.性别 = C.名称(+) " ' And A.当前科室id IS NOT NULL And B.出院日期 Is NULL"
    ElseIf iInputType = 3 Then '门诊号
        strSQL = "Select Distinct Decode(A.当前科室id,Null,1,2) As PatientType,A.主页ID,Decode(A.当前科室id,Null,Nvl(B.执行部门ID,0),A.当前科室id) As 病人科室,B.执行人 As 医生," & _
            "a.姓名,decode(c.名称,null,a.性别,c.编码 || '-' || a.性别) as 性别,a.年龄,a.病人id,a.住院号,a.门诊号,a.当前床号,Zl_Age_Calc(A.病人ID) as 年龄1  " & _
            " From 病人信息 A,(Select NO,执行部门ID,执行人,病人ID,门诊号,记录性质,记录状态 From 病人挂号记录 Where 登记时间>sysdate-[5]) B,性别 C Where A.门诊号=[2] And A.病人ID=B.病人ID(+) and (b.病人ID is null or(b.记录状态 =1 and b.记录性质 =1)) And A.门诊号=B.门诊号(+) And a.性别 = C.名称(+) "
'            " And (A.当前科室id IS NOT NULL Or NVL(B.执行状态,1) IN (0,2))"
    ElseIf iInputType = 4 Then '挂号单
        strNO = GetFullNO(strSeek, 12)
'        strsql = "Select Decode(B.主页id, Null, 1, 2) As Patienttype, Nvl(B.主页id, 0) As 主页id, Nvl(B.执行部门id, 0) As 病人科室," & vbNewLine & _
                "       B.执行人 As 医生, A.姓名, Decode(C.名称, Null, A.性别, C.编码 || '-' || A.性别) As 性别, A.年龄, A.病人id," & vbNewLine & _
                "       A.住院号, A.门诊号, A.当前床号,Zl_Age_Calc(A.病人ID) as 年龄1  " & vbNewLine & _
                "From 病人信息 A, 住院费用记录 B, 性别 C" & vbNewLine & _
                "Where B.记录性质 = 4 And B.病人id = A.病人id And A.性别 = C.名称(+) And B.记录状态 In (1, 3) And B.序号 = 1 And B.NO = [3] "
        strSQL = "Select 1 As Patienttype, 0 As 主页id, Nvl(B.执行部门id, 0) As 病人科室," & vbNewLine & _
                "       B.执行人 As 医生, A.姓名, Decode(C.名称, Null, A.性别, C.编码 || '-' || A.性别) As 性别, A.年龄, A.病人id," & vbNewLine & _
                "       A.住院号, A.门诊号, A.当前床号,Zl_Age_Calc(A.病人ID) as 年龄1  " & vbNewLine & _
                "From 病人信息 A, 门诊费用记录 B, 性别 C" & vbNewLine & _
                "Where B.记录性质 = 4 And B.病人id = A.病人id And A.性别 = C.名称(+) And B.记录状态 In (1, 3) And B.序号 = 1 And B.NO = [3] "
'        strSQLbak = strsql
'        strSQLbak = Replace$(strSQLbak, "住院费用记录", "门诊费用记录")
'        strSQLbak = Replace$(strSQLbak, "Decode(B.主页id, Null, 1, 2) As Patienttype", "1 As Patienttype")
'        strSQLbak = Replace$(strSQLbak, "Nvl(B.主页id, 0)", "0")
'        strsql = strsql & " union all " & strSQLbak

    ElseIf iInputType = 5 Then '收费单据号
        strNO = GetFullNO(strSeek, 13): mstrNO = strNO

        strSQL = "Select 1 As Patienttype, 0 As 主页id," & vbNewLine & _
                "       Nvl(A.当前科室id, B.开单部门id) As 病人科室, B.开单人 As 医生, B.姓名," & vbNewLine & _
                "       Decode(C.名称, Null, B.性别, C.编码 || '-' || B.性别) As 性别, B.年龄, A.病人id, A.单位电话, A.工作单位," & vbNewLine & _
                "       A.单位邮编, A.家庭地址, A.家庭电话, A.家庭地址邮编, A.门诊号, A.身份证号, A.费别, A.医疗付款方式, A.国籍, A.婚姻状况," & vbNewLine & _
                "       A.民族, A.职业,decode(a.病人ID,Null,b.年龄,Zl_Age_Calc(A.病人ID)) as 年龄1 " & vbNewLine & _
                "From 病人信息 A, 门诊费用记录 B, 性别 C" & vbNewLine & _
                "Where B.病人id = A.病人id(+) And B.性别 = C.名称(+) And Mod(B.记录性质,10) = 1 And B.记录状态 In (1, 3) And B.序号 = 1 And" & vbNewLine & _
                "      B.NO = [3] " & vbNewLine

'        strSQLbak = strsql
'        strSQLbak = Replace$(strSQLbak, "住院费用记录", "门诊费用记录")
'        strSQLbak = Replace$(strSQLbak, "Decode(B.主页id, Null, 1, 2) As Patienttype", "1 As Patienttype")
'        strSQLbak = Replace$(strSQLbak, "Nvl(B.主页id, 0)", "0")
'        strsql = strsql & " union all " & strSQLbak & " Order By 病人id "
    ElseIf iInputType = 7 Then '带字母的就诊卡

        strSQL = "Select Distinct Decode(a.当前科室id, Null, 1, 2) As Patienttype, A.主页id," & vbNewLine & _
                "                               Decode(a.当前科室id, Null, Nvl(b.执行部门id, 0), a.当前科室id) As 病人科室, b.执行人 As 医生, a.姓名," & vbNewLine & _
                "                               Decode(c.名称, Null, a.性别, c.编码 || '-' || a.性别) As 性别, a.年龄, a.病人id, a.住院号, a.门诊号," & vbNewLine & _
                "                               a.当前床号,Zl_Age_Calc(A.病人ID) as 年龄1 " & vbNewLine & _
                "From 病人信息 a, (Select 执行部门id, 执行人, 病人id, 门诊号,记录状态,记录性质 From 病人挂号记录 Where 登记时间 > Sysdate - [5]) b, 性别 c" & vbNewLine & _
                "Where a.就诊卡号 = [1]  And a.病人id = b.病人id(+) and (b.病人ID is null or (b.记录状态=1 and b.记录性质 =1)) And a.门诊号 = b.门诊号(+) And a.性别 = c.名称(+)"

    ElseIf iInputType = 8 Then '医保卡
        strSQL = "Select Distinct Decode(a.当前科室id, Null, 1, 2) As Patienttype, A.主页id," & vbNewLine & _
                "                               Decode(a.当前科室id, Null, Nvl(b.执行部门id, 0), a.当前科室id) As 病人科室, b.执行人 As 医生, a.姓名," & vbNewLine & _
                "                               Decode(c.名称, Null, a.性别, c.编码 || '-' || a.性别) As 性别, a.年龄, a.病人id, a.住院号, a.门诊号," & vbNewLine & _
                "                               a.当前床号,Zl_Age_Calc(A.病人ID) as 年龄1 " & vbNewLine & _
                "From 病人信息 a, (Select 执行部门id, 执行人, 病人id, 门诊号,记录状态,记录性质 From 病人挂号记录 Where 登记时间 > Sysdate - [5]) b, 性别 c" & vbNewLine & _
                "Where (a.医保号 = [1] or a.IC卡号= [1]) And a.病人id = b.病人id(+) and (b.病人ID is null or (b.记录性质=1 and b.记录状态=1)) And a.门诊号 = b.门诊号(+) And a.性别 = c.名称(+)"
    ElseIf iInputType = 9 Then '身份证
        strSQL = "Select Distinct Decode(a.当前科室id, Null, 1, 2) As Patienttype, A.主页id," & vbNewLine & _
                "                               Decode(a.当前科室id, Null, Nvl(b.执行部门id, 0), a.当前科室id) As 病人科室, b.执行人 As 医生, a.姓名," & vbNewLine & _
                "                               Decode(c.名称, Null, a.性别, c.编码 || '-' || a.性别) As 性别, a.年龄, a.病人id, a.住院号, a.门诊号," & vbNewLine & _
                "                               a.当前床号,Zl_Age_Calc(A.病人ID) as 年龄1 " & vbNewLine & _
                "From 病人信息 a, (Select 执行部门id, 执行人, 病人id, 门诊号,记录状态,记录性质 From 病人挂号记录 Where 登记时间 > Sysdate - [5]) b, 性别 c" & vbNewLine & _
                "Where a.身份证号 = [1]  And a.病人id = b.病人id(+) And a.门诊号 = b.门诊号(+) and (b.病人ID is null or (b.记录状态 =1 and b.记录性质 =1)) And a.性别 = c.名称(+)"

    Else '当作姓名
        strSQL = "Select Decode(A.当前科室id,Null,1,2) As PatientType,A.主页ID,Nvl(A.当前科室id,Nvl(C.申请科室ID,0)) As 病人科室," & _
            "a.姓名,decode(b.名称,null,a.性别,b.编码 || '-' || a.性别) as 性别,a.年龄,a.病人id,a.住院号,a.门诊号,a.当前床号,'' as 医生,Zl_Age_Calc(A.病人ID) as 年龄1  " & _
            " From 病人信息 A , 性别 B,检验标本记录 C Where A.病人id=[4] And a.性别 = b.名称(+) And a.病人id = c.病人id(+)"
    End If

    Set GetPatient = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strSeek, Val(strSeek), strNO, mlng病人ID, lng挂号效期)
    GetPatient.filter = ""
    If GetPatient.RecordCount > 1 Then
        If iInputType = 3 Or iInputType = 0 Or iInputType >= 7 Then
            '门诊号查询时，一个病人挂了多个号，则需要选择.

            If iInputType = 0 Then
                strSQL = _
                    "Select Distinct Decode(A.当前科室id, Null, 1, 2) As Patienttype, A.主页id," & vbNewLine & _
                    "                Decode(A.当前科室id, Null, Nvl(B.执行部门id, 0), A.当前科室id) As 病人科室, B.执行人 As 医生, A.姓名," & vbNewLine & _
                    "                Decode(C.名称, Null, A.性别, C.编码 || '-' || A.性别) As 性别, A.年龄, A.病人id, A.住院号, A.门诊号, A.当前床号,Zl_Age_Calc(A.病人ID) as 年龄1 " & vbNewLine & _
                    "From (Select " & gConst_病人信息_列名 & " From 病人信息 a Where 就诊卡号 = [1]" & vbNewLine & _
                    "       Union " & vbNewLine & _
                    "       Select " & gConst_病人信息_列名 & " From 病人信息 a Where 门诊号 = [2]" & vbNewLine & _
                    "       Union " & vbNewLine & _
                    "       Select " & gConst_病人信息_列名 & " From 病人信息 a Where 住院号 = [2]) A, 病人挂号记录 B, 性别 C" & vbNewLine & _
                    "Where A.病人id = B.病人id(+) And A.门诊号 = B.门诊号(+) and (b.门诊号 is null or (b.记录状态 =1 and b.记录性质=1)) And 0 + B.登记时间(+) > Sysdate - [5] And  A.性别 = C.名称(+) And B.NO = [6]"

'                strSQL = "Select Distinct Decode(a.当前科室id, Null, 1, 2) As Patienttype, Nvl(a.住院次数, 0) As 主页id," & vbNewLine & _
'                        "                               Decode(a.当前科室id, Null, Nvl(b.执行部门id, 0), a.当前科室id) As 病人科室, b.执行人 As 医生, a.姓名," & vbNewLine & _
'                        "                               Decode(c.名称, Null, a.性别, c.编码 || '-' || a.性别) As 性别, a.年龄, a.病人id, a.住院号, a.门诊号," & vbNewLine & _
'                        "                               a.当前床号" & vbNewLine & _
'                        "From 病人信息 a, (Select NO,执行部门id, 执行人, 病人id, 门诊号 From 病人挂号记录 Where 登记时间 > Sysdate - [5]) b, 性别 c" & vbNewLine & _
'                        "Where (a.就诊卡号 = [1] Or a.门诊号 =[2] Or a.住院号=[2]) And a.病人id = b.病人id(+) And a.门诊号 = b.门诊号(+) And a.性别 = c.名称(+) And B.NO=[6]"

            ElseIf iInputType = 3 Then
                strSQL = "Select Distinct Decode(a.当前科室id, Null, 1, 2) As Patienttype, A.主页id," & vbNewLine & _
                        "                               Decode(a.当前科室id, Null, Nvl(b.执行部门id, 0), a.当前科室id) As 病人科室, b.执行人 As 医生, a.姓名," & vbNewLine & _
                        "                               Decode(c.名称, Null, a.性别, c.编码 || '-' || a.性别) As 性别, a.年龄, a.病人id, a.住院号, a.门诊号," & vbNewLine & _
                        "                               a.当前床号,Zl_Age_Calc(A.病人ID) as 年龄1 " & vbNewLine & _
                        "From 病人信息 a, (Select NO,执行部门id, 执行人, 病人id, 门诊号,记录状态,记录性质 From 病人挂号记录 Where 登记时间 > Sysdate - [5]) b, 性别 c" & vbNewLine & _
                        "Where a.门诊号 = [2] And a.病人id = b.病人id(+) And a.门诊号 = b.门诊号(+) and (b.病人ID is null or(b.记录状态=1 and 记录性质=1)) And a.性别 = c.名称(+) And B.NO=[6]"
            ElseIf iInputType = 7 Then
                strSQL = "Select Distinct Decode(a.当前科室id, Null, 1, 2) As Patienttype, A.主页id," & vbNewLine & _
                        "                               Decode(a.当前科室id, Null, Nvl(b.执行部门id, 0), a.当前科室id) As 病人科室, b.执行人 As 医生, a.姓名," & vbNewLine & _
                        "                               Decode(c.名称, Null, a.性别, c.编码 || '-' || a.性别) As 性别, a.年龄, a.病人id, a.住院号, a.门诊号," & vbNewLine & _
                        "                               a.当前床号,Zl_Age_Calc(A.病人ID) as 年龄1 " & vbNewLine & _
                        "From 病人信息 a, (Select NO,执行部门id, 执行人, 病人id, 门诊号,记录状态,记录性质 From 病人挂号记录 Where 登记时间 > Sysdate - [5]) b, 性别 c" & vbNewLine & _
                        "Where a.就诊卡号 = [1] And a.病人id = b.病人id(+) And a.门诊号 = b.门诊号(+) and (b.病人ID is null or (b.记录状态=1 and b.记录性质 =1)) And a.性别 = c.名称(+) And B.NO=[6]"
            ElseIf iInputType = 8 Then

                strSQL = "Select Distinct Decode(a.当前科室id, Null, 1, 2) As Patienttype, A.主页id," & vbNewLine & _
                        "                               Decode(a.当前科室id, Null, Nvl(b.执行部门id, 0), a.当前科室id) As 病人科室, b.执行人 As 医生, a.姓名," & vbNewLine & _
                        "                               Decode(c.名称, Null, a.性别, c.编码 || '-' || a.性别) As 性别, a.年龄, a.病人id, a.住院号, a.门诊号," & vbNewLine & _
                        "                               a.当前床号,Zl_Age_Calc(A.病人ID) as 年龄1 " & vbNewLine & _
                        "From 病人信息 a, (Select NO,执行部门id, 执行人, 病人id, 门诊号,记录状态,记录性质 From 病人挂号记录 Where 登记时间 > Sysdate - [5]) b, 性别 c" & vbNewLine & _
                        "Where (a.医保号 = [1] or a.IC卡号= [1]) And a.病人id = b.病人id(+) And a.门诊号 = b.门诊号(+) and (b.病人ID is null or (b.记录状态=1 and b.记录性质=1)) And a.性别 = c.名称(+) And B.NO=[6]"
            ElseIf iInputType = 9 Then

                strSQL = "Select Distinct Decode(a.当前科室id, Null, 1, 2) As Patienttype, A.主页id," & vbNewLine & _
                        "                               Decode(a.当前科室id, Null, Nvl(b.执行部门id, 0), a.当前科室id) As 病人科室, b.执行人 As 医生, a.姓名," & vbNewLine & _
                        "                               Decode(c.名称, Null, a.性别, c.编码 || '-' || a.性别) As 性别, a.年龄, a.病人id, a.住院号, a.门诊号," & vbNewLine & _
                        "                               a.当前床号,Zl_Age_Calc(A.病人ID) as 年龄1 " & vbNewLine & _
                        "From 病人信息 a, (Select NO,执行部门id, 执行人, 病人id, 门诊号,记录状态,记录性质 From 病人挂号记录 Where 登记时间 > Sysdate - [5]) b, 性别 c" & vbNewLine & _
                        "Where  a.身份证号 = [1] And a.病人id = b.病人id(+) And a.门诊号 = b.门诊号(+) and (b.病人ID is null or (b.记录状态=1 and b.记录性质=1)) And a.性别 = c.名称(+) And B.NO=[6]"

            End If
            '---- 用于选择器
            If iInputType = 0 Then
                strTmp = _
                    "Select Distinct Rownum As ID, B.NO As 挂号单号, B.执行人 As 医生, A.门诊号, A.就诊卡号, A.住院号, A.姓名," & vbNewLine & _
                    "                Decode(A.当前科室id, Null, D.名称, E.名称) As 病人科室, Decode(C.名称, Null, A.性别, C.编码 || '-' || A.性别) As 性别, A.年龄," & vbNewLine & _
                    "                To_Char(B.登记时间, 'yyyy-MM-dd HH24:MI:SS') As 挂号时间, A.病人id,Zl_Age_Calc(A.病人ID) as 年龄1 " & vbNewLine & _
                    "From (Select " & gConst_病人信息_列名 & " From 病人信息 a Where 就诊卡号 = [1]" & vbNewLine & _
                    "       Union " & vbNewLine & _
                    "       Select " & gConst_病人信息_列名 & " From 病人信息 a Where 门诊号 = [2]" & vbNewLine & _
                    "       Union " & vbNewLine & _
                    "       Select " & gConst_病人信息_列名 & " From 病人信息 a Where 住院号 = [2]) A, 病人挂号记录 B, 性别 C, 部门表 D, 部门表 E" & vbNewLine & _
                    "Where A.病人id = B.病人id(+) and (b.病人id is null or (b.记录状态=1 and b.记录性质=1)) And A.门诊号 = B.门诊号(+) And 0 + B.登记时间(+) > Sysdate - [3] And A.性别 = C.名称(+) And" & vbNewLine & _
                    "      Nvl(B.执行部门id, 0) = D.ID(+) And A.当前科室id = E.ID(+)" & vbNewLine & _
                    "Order By B.NO Desc"

'                strTmp = "Select Distinct rownum As ID, B.NO As 挂号单号, B.执行人 As 医生, A.门诊号, A.就诊卡号, A.住院号, A.姓名," & vbNewLine & _
'                        "                Decode(A.当前科室id, Null, D.名称, E.名称) As 病人科室, Decode(C.名称, Null, A.性别, C.编码 || '-' || A.性别) As 性别, A.年龄," & vbNewLine & _
'                        "                To_Char(B.登记时间, 'yyyy-MM-dd HH24:MI:SS') As 挂号时间, A.病人id" & vbNewLine & _
'                        "From 病人信息 A, (Select 登记时间, NO, 执行部门id, 执行人, 病人id, 门诊号 From 病人挂号记录 Where 登记时间 > Sysdate - [3]) B, 性别 C, 部门表 D, 部门表 E" & vbNewLine & _
'                        "Where (A.就诊卡号 = [1] Or A.门诊号 = [2] Or 住院号=[2]) And A.病人id = B.病人id(+) And A.门诊号 = B.门诊号(+) And A.性别 = C.名称(+) And" & vbNewLine & _
'                        "      Nvl(B.执行部门id, 0) = D.ID(+) And A.当前科室id = E.ID(+)" & vbNewLine & _
'                        "Order By B.NO Desc"

            ElseIf iInputType = 3 Then
                strTmp = "Select Distinct rownum As ID, B.NO As 挂号单号, B.执行人 As 医生, A.门诊号, A.就诊卡号, A.住院号, A.姓名," & vbNewLine & _
                        "                Decode(A.当前科室id, Null, D.名称, E.名称) As 病人科室, Decode(C.名称, Null, A.性别, C.编码 || '-' || A.性别) As 性别, A.年龄," & vbNewLine & _
                        "                To_Char(B.登记时间, 'yyyy-MM-dd HH24:MI:SS') As 挂号时间, A.病人id,Zl_Age_Calc(A.病人ID) as 年龄1 " & vbNewLine & _
                        "From 病人信息 A, (Select 登记时间, NO, 执行部门id, 执行人, 病人id, 门诊号,记录状态,记录性质 From 病人挂号记录 Where 登记时间 > Sysdate - [3]) B, 性别 C, 部门表 D, 部门表 E" & vbNewLine & _
                        "Where A.门诊号 = [2] And A.病人id = B.病人id(+) And A.门诊号 = B.门诊号(+) and (b.病人ID is null or (b.记录状态=1 and b.记录性质 =1)) And A.性别 = C.名称(+) And" & vbNewLine & _
                        "      Nvl(B.执行部门id, 0) = D.ID(+) And A.当前科室id = E.ID(+)" & vbNewLine & _
                        "Order By B.NO Desc"

            ElseIf iInputType = 7 Then
                strTmp = "Select Distinct rownum As ID, B.NO As 挂号单号, B.执行人 As 医生, A.门诊号, A.就诊卡号, A.住院号, A.姓名," & vbNewLine & _
                        "                Decode(A.当前科室id, Null, D.名称, E.名称) As 病人科室, Decode(C.名称, Null, A.性别, C.编码 || '-' || A.性别) As 性别, A.年龄," & vbNewLine & _
                        "                To_Char(B.登记时间, 'yyyy-MM-dd HH24:MI:SS') As 挂号时间, A.病人id,Zl_Age_Calc(A.病人ID) as 年龄1 " & vbNewLine & _
                        "From 病人信息 A, (Select 登记时间, NO, 执行部门id, 执行人, 病人id, 门诊号,记录状态,记录性质 From 病人挂号记录 Where 登记时间 > Sysdate - [3]) B, 性别 C, 部门表 D, 部门表 E" & vbNewLine & _
                        "Where A.就诊卡号 = [1] And A.病人id = B.病人id(+) And A.门诊号 = B.门诊号(+) and (b.病人ID is null or (b.记录状态=1 and b.记录性质=1)) And A.性别 = C.名称(+) And" & vbNewLine & _
                        "      Nvl(B.执行部门id, 0) = D.ID(+) And A.当前科室id = E.ID(+)" & vbNewLine & _
                        "Order By B.NO Desc"
            ElseIf iInputType = 8 Then
                strTmp = "Select Distinct rownum As ID, B.NO As 挂号单号, B.执行人 As 医生, A.门诊号, A.就诊卡号, A.住院号, A.姓名," & vbNewLine & _
                        "                Decode(A.当前科室id, Null, D.名称, E.名称) As 病人科室, Decode(C.名称, Null, A.性别, C.编码 || '-' || A.性别) As 性别, A.年龄," & vbNewLine & _
                        "                To_Char(B.登记时间, 'yyyy-MM-dd HH24:MI:SS') As 挂号时间, A.病人id,Zl_Age_Calc(A.病人ID) as 年龄1 " & vbNewLine & _
                        "From 病人信息 A, (Select 登记时间, NO, 执行部门id, 执行人, 病人id, 门诊号,记录状态,记录性质 From 病人挂号记录 Where 登记时间 > Sysdate - [3]) B, 性别 C, 部门表 D, 部门表 E" & vbNewLine & _
                        "Where (A.医保号 = [1] or a.IC卡号= [1]) And A.病人id = B.病人id(+) And A.门诊号 = B.门诊号(+) and (b.病人ID is null or (b.记录状态=1 and b.记录性质=1)) And A.性别 = C.名称(+) And" & vbNewLine & _
                        "      Nvl(B.执行部门id, 0) = D.ID(+) And A.当前科室id = E.ID(+)" & vbNewLine & _
                        "Order By B.NO Desc"
            ElseIf iInputType = 9 Then
                strTmp = "Select Distinct rownum As ID, B.NO As 挂号单号, B.执行人 As 医生, A.门诊号, A.就诊卡号, A.住院号, A.姓名," & vbNewLine & _
                        "                Decode(A.当前科室id, Null, D.名称, E.名称) As 病人科室, Decode(C.名称, Null, A.性别, C.编码 || '-' || A.性别) As 性别, A.年龄," & vbNewLine & _
                        "                To_Char(B.登记时间, 'yyyy-MM-dd HH24:MI:SS') As 挂号时间, A.病人id,Zl_Age_Calc(A.病人ID) as 年龄1 " & vbNewLine & _
                        "From 病人信息 A, (Select 登记时间, NO, 执行部门id, 执行人, 病人id, 门诊号,记录状态,记录性质 From 病人挂号记录 Where 登记时间 > Sysdate - [3]) B, 性别 C, 部门表 D, 部门表 E" & vbNewLine & _
                        "Where A.身份证号 = [1] And A.病人id = B.病人id(+) And A.门诊号 = B.门诊号(+) and (b.病人ID is null or (b.记录状态=1 and b.记录性质=1)) And A.性别 = C.名称(+) And" & vbNewLine & _
                        "      Nvl(B.执行部门id, 0) = D.ID(+) And A.当前科室id = E.ID(+)" & vbNewLine & _
                        "Order By B.NO Desc"
            End If
            mblnEdit = False
            Call ClientToScreen(txt姓名.hWnd, objPoint)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strTmp, 0, "病人选择", True, "", "", True, True, True, objPoint.X * 15, objPoint.Y * 15, Me.txt姓名.Height, _
                                                    False, True, False, strSeek, Val(strSeek), lng挂号效期)
            mblnEdit = True
            GetPatient.filter = "病人ID=0"
            If Not rsTmp Is Nothing Then
                If rsTmp.State = adStateOpen Then
                    If rsTmp.RecordCount = 1 Then
                        str挂号单 = "" & Nvl(rsTmp.Fields("挂号单号"), 0)
                        If str挂号单 = "0" Then
                            strSQL = Replace(strSQL, " And B.NO = [6]", " And A.病人ID = [6]")
                            str挂号单 = rsTmp.Fields("病人ID")
                            Set GetPatient = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strSeek, Val(strSeek), strNO, mlng病人ID, lng挂号效期, Val(str挂号单))
                        Else
                            Set GetPatient = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strSeek, Val(strSeek), strNO, mlng病人ID, lng挂号效期, str挂号单)
                        End If

                        If GetPatient.RecordCount > 1 Then
                            MsgBox "查到找了多个病人！请加上前缀标志进行查找!"
                            GetPatient.filter = "病人ID=0"
                        End If
                    End If

                End If
            End If
        Else
            MsgBox "查到找了多个病人！请加上前缀标志进行查找!"
            GetPatient.filter = "病人ID=0"
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function OpenSelect(ByVal strText As String, Optional ByVal blnWhere As Boolean = False) As Byte
    '--------------------------------------------------------------------------------------------------------
    '功能:打开列表结构的申请检验单
    '参数:strText       过滤关键字(可能为住院号，门诊号，床位号或条码号)
    '返回:0             取消返回
    '     1             成功返回
    '     2             出错返回
    '     3             发现核收有不满足条件时，1.有两个不相同的标本.2.婴儿医嘱和母亲医嘱在一起
    '--------------------------------------------------------------------------------------------------------
    Dim strInput As String, i As Integer
    Dim rs As New ADODB.Recordset, rsTmp As New ADODB.Recordset
    Dim strLvw As String, strSQL As String
    Dim objPoint As POINTAPI
    Dim strLastKeys As String
    Dim strStart As String, strEnd As String
    Dim strField As String
    Dim blnNoRange As Boolean, blnCheck As Boolean
    Dim mstrSql As String
    Dim lngCurrDevice As Long
    Dim blMachineFind As Boolean
    Dim strTmp As String
    Dim int主页ID As Integer
    Dim intCount As Integer
    Dim strFilter As String
    Dim strDate As String
    Dim lng病人ID As Long
    Dim int执行状态 As Integer
    Dim bln未收费显示 As Boolean
    Dim str执行科室名称 As String
    Dim lng执行科室ID As Long
    Dim strSQLbak As String
    Dim strAge As String, aAge As Variant
    Dim intPatProperty As Integer, rsPatProperty As New ADODB.Recordset   '门诊留观病人

    On Error GoTo ErrHand

    OpenSelect = 2

    If blnCheck Then '显示收费
        strLvw = "姓名,900,0,1;收费,600,0,1;申请时间,1300,0,0;申请项目,1800,0,0;申请科室,1800,0,0;申请人,810,0,0"
    Else
        strLvw = "姓名,900,0,1;申请时间,1300,0,0;申请项目,1800,0,0;申请科室,1800,0,0;申请人,810,0,0"
    End If

    bln未收费显示 = InStr(mstrPrivs, "未收费核收") > 0

    blnNoRange = Val(zlDatabase.GetPara("核收忽略时间", 100, 1208, 1))
    blnCheck = Val(zlDatabase.GetPara("核收显示收费", 100, 1208, 1))
    blMachineFind = Val(zlDatabase.GetPara("按仪器项目核收", 100, 1208, 1))


    strStart = GetDateTime(Split(zlDatabase.GetPara("待核收范围", 100, 1208, "今  天") & ";", ";")(0), 1)
    strEnd = GetDateTime(Split(zlDatabase.GetPara("待核收范围", 100, 1208, "今  天") & ";", ";")(0), 2)

    If strStart = "自定义" Then
        strStart = Format(Split(zlDatabase.GetPara("待核收范围", 100, 1208, "今  天") & ";", ";")(1), "yyyy-mm-dd 00:00:00")
        strEnd = Format(Split(zlDatabase.GetPara("待核收范围", 100, 1208, "今  天") & ";", ";")(2), "yyyy-mm-dd 23:59:59")
    Else
        If strStart = "" Then strStart = GetDateTime("今  天", 1)
        If strEnd = "" Then strEnd = GetDateTime("今  天", 2)
    End If

    lngCurrDevice = 0
    If vsf2.Rows > 1 Then
        lngCurrDevice = Val(vsf2.RowData(1))
    End If

    If mlngDefaultDevice = -1 Then
        If blnCheck Then '显示收费
            If mblnBarCode Then
                mstrSql = "SELECT Decode(SUM(Decode(F.样本条码,[2],1,0)),0,0,1) AS 选择"
            Else
                mstrSql = "SELECT Decode(Nvl(z.医嘱序号,0), 0, 1, Decode(Sum(Nvl(z.数次,0)), 0, 0, 1)) As 选择"
            End If
            mstrSql = mstrSql & ",A.相关ID AS ID,  " & _
                              "C.姓名||Decode(A.婴儿,0,'',Null,'','(婴儿)') As 姓名," & _
                              "C.门诊号," & _
                              "C.住院号," & _
                              "D.名称 AS 申请科室," & _
                              "A.开嘱医生 AS 申请人," & _
                              "f.接收人,f.接收时间, " & _
                              "'Item' AS 图标,Decode(Nvl(z.医嘱序号,0), 0, '√', Decode(Sum(Nvl(z.数次,0)), 0, '×', '√')) As 收费,NVL(A.紧急标志,0) AS 紧急,Y.操作类型,MAX(Decode(H.项目类别,2,2,1)) As 项目类别,MAX(F.采样人) AS 采样人,MAX(F.采样时间) AS 采样时间 " & _
                         "FROM 病人医嘱记录 A," & _
                         "病人信息 C,部门表 D,病人医嘱发送 F,检验报告项目 G,检验项目 H,诊疗项目目录 Y,住院费用记录 Z " & _
                        "WHERE A.诊疗类别 = 'C' " & _
                              "AND A.病人ID=C.病人ID " & _
                              "AND A.开嘱科室ID=D.ID " & _
                              "AND A.相关id IS NOT NULL " & _
                              "AND A.医嘱状态=8 AND A.ID=F.医嘱id " & _
                              "AND A.诊疗项目id=G.诊疗项目id AND G.细菌ID Is Null " & _
                              "AND G.报告项目ID=H.诊治项目ID " & _
                              "AND A.诊疗项目ID=Y.ID " & _
                              IIf(mbln划价单模式 = True, " and a.病人来源 = 2 and c.出院时间 is null ", " ")

            mstrSql = mstrSql & " " & _
                              "AND F.执行状态 =0 AND A.病人ID=[1] AND A.执行科室id+0=[3] " & _
                              IIf(blnNoRange, " ", " AND A.开嘱时间 BETWEEN [5] and [6] ") & _
                              "AND F.NO=Z.NO(+) AND F.记录性质=mod(Z.记录性质(+),10) AND F.医嘱id=Z.医嘱序号(+)+0 " & _
                              IIf(mblnBarCode, " And F.样本条码=[2]  ", "") & _
                              IIf(int主页ID = 0, " ", " And a.主页ID = [9] ") & _
                              " And a.病人来源 = 2 " & _
                              IIf(mblnPrice = True, " And z.No = [10] ", " ") & _
                              "GROUP BY A.相关ID,a.id,C.姓名||Decode(A.婴儿,0,'',Null,'','(婴儿)'),C.门诊号,C.住院号,D.名称,A.开嘱医生,'Item',NVL(A.紧急标志,0),Y.操作类型,f.接收人,f.接收时间 ,z.医嘱序号 "

            mstrSql = mstrSql & " Union all "

            If mblnBarCode Then
                mstrSql = mstrSql & "SELECT Decode(SUM(Decode(F.样本条码,[2],1,0)),0,0,1) AS 选择"
            Else
                mstrSql = mstrSql & "SELECT Decode(Nvl(z.医嘱序号,0), 0, 1, Decode(Sum(Nvl(z.数次,0)), 0, 0, 1)) As 选择"
            End If

            mstrSql = mstrSql & ",A.相关ID AS ID, " & _
                              "C.姓名||Decode(A.婴儿,0,'',Null,'','(婴儿)') As 姓名," & _
                              "C.门诊号," & _
                              "C.住院号," & _
                              "D.名称 AS 申请科室," & _
                              "A.开嘱医生 AS 申请人," & _
                              "f.接收人,f.接收时间, " & _
                              "'Item' AS 图标,Decode(Nvl(z.医嘱序号,0), 0, '√', Decode(Sum(Nvl(z.数次,0)), 0, '×', '√')) As 收费,NVL(A.紧急标志,0) AS 紧急,Y.操作类型,MAX(Decode(H.项目类别,2,2,1)) As 项目类别,MAX(F.采样人) AS 采样人,MAX(F.采样时间) AS 采样时间 " & _
                         "FROM 病人医嘱记录 A," & _
                         "病人信息 C,部门表 D,病人医嘱发送 F,检验报告项目 G,检验项目 H,诊疗项目目录 Y,住院费用记录 Z,检验标本记录 J " & _
                        "WHERE A.诊疗类别 = 'C' " & _
                              "AND A.病人ID=C.病人ID " & _
                              "AND A.开嘱科室ID=D.ID " & _
                              "AND A.相关id IS NOT NULL " & _
                              "AND A.医嘱状态=8 AND A.ID=F.医嘱id " & _
                              "AND A.诊疗项目id=G.诊疗项目id AND G.细菌ID Is Null " & _
                              "AND G.报告项目ID=H.诊治项目ID " & _
                              "AND A.诊疗项目ID=Y.ID " & _
                              IIf(mbln划价单模式 = True, " and a.病人来源 = 2 and c.出院时间 is null ", " ")

            mstrSql = mstrSql & " " & _
                              "AND A.病人ID=[1] AND A.执行科室id+0=[3] " & _
                              IIf(blnNoRange, " ", " AND A.开嘱时间 BETWEEN [5] and [6] ") & _
                              "AND F.NO=Z.NO(+) AND F.记录性质=mod(Z.记录性质(+),10) AND F.医嘱id=Z.医嘱序号(+)+0 And a.相关id = j.医嘱ID(+) And j.id = [8] " & _
                              IIf(mblnBarCode, " And F.样本条码=[2]   ", "") & _
                              IIf(int主页ID = 0, " ", " And a.主页ID = [9] ") & _
                              " And a.病人来源 = 2 " & _
                              IIf(mblnPrice = True, " And z.No = [10] ", " ") & _
                              "GROUP BY A.相关ID,a.id,C.姓名||Decode(A.婴儿,0,'',Null,'','(婴儿)'),C.门诊号,C.住院号,D.名称,A.开嘱医生,'Item',NVL(A.紧急标志,0),Y.操作类型,f.接收人,f.接收时间 ,z.医嘱序号 "
        Else
            If mblnBarCode Then
                mstrSql = mstrSql & "SELECT Distinct Decode 1 AS 选择"
            Else
                mstrSql = mstrSql & "SELECT Distinct 1 AS 选择"
            End If
            mstrSql = mstrSql & ",A.相关ID AS ID, " & _
                              "C.姓名||Decode(A.婴儿,0,'',Null,'','(婴儿)') As 姓名," & _
                              "C.门诊号," & _
                              "C.住院号," & _
                              "D.名称 AS 申请科室," & _
                              "A.开嘱医生 AS 申请人," & _
                              "f.接收人,f.接收时间, " & _
                              "'Item' AS 图标,Decode(Nvl(z.医嘱序号,0), 0, '√', Decode(Sum(Nvl(z.数次,0)), 0, '×', '√')) As 收费,NVL(A.紧急标志,0) AS 紧急,Y.操作类型,Decode(H.项目类别,2,2,1) As 项目类别,F.采样人,F.采样时间 As 采样时间 " & _
                         "FROM 病人医嘱记录 A," & _
                         "病人信息 C,部门表 D,病人医嘱发送 F,检验报告项目 G,检验项目 H,诊疗项目目录 Y,住院费用记录 Z " & _
                        "WHERE A.诊疗类别 = 'C' " & _
                              "AND A.病人ID=C.病人ID " & _
                              "AND A.开嘱科室ID=D.ID " & _
                              "AND A.相关id IS NOT NULL " & _
                              "AND A.医嘱状态=8 AND A.ID=F.医嘱id " & _
                              "AND A.诊疗项目id=G.诊疗项目id AND G.细菌ID Is Null " & _
                              "AND G.报告项目ID=H.诊治项目ID " & _
                              "AND A.诊疗项目ID=Y.ID " & _
                              "And a.病人来源 = 2 " & _
                              IIf(mbln划价单模式 = True, " and a.病人来源 = 2 and c.出院时间 is null ", " ")

            mstrSql = mstrSql & " " & _
                              "AND F.执行状态 = 0 AND A.病人ID=[1] AND A.执行科室id+0=[3] " & _
                              "AND F.NO=Z.NO(+) AND F.记录性质=mod(Z.记录性质(+),10) AND F.医嘱id=Z.医嘱序号(+)+0 " & _
                              IIf(mblnBarCode, " And F.样本条码=[2]   ", "") & _
                              IIf(int主页ID = 0, " ", " And a.主页ID = [9] ") & _
                              IIf(blnNoRange, " ", " AND A.开嘱时间 BETWEEN [5] and [6] ") & _
                              IIf(mblnPrice = True, " And z.No = [10] ", " ") & _
                              " GROUP BY A.相关ID,a.id,C.姓名||Decode(A.婴儿,0,'',Null,'','(婴儿)'),C.门诊号,C.住院号,D.名称,A.开嘱医生,'Item',NVL(A.紧急标志,0),Y.操作类型,f.接收人,f.接收时间,Decode(H.项目类别,2,2,1),F.采样人,F.采样时间 ,z.医嘱序号 "

            mstrSql = mstrSql & " Union all "

            If mblnBarCode Then
                mstrSql = mstrSql & "SELECT Distinct 1 AS 选择"
            Else
                mstrSql = mstrSql & "SELECT Distinct 1 AS 选择"
            End If

            mstrSql = mstrSql & ",A.相关ID AS ID, " & _
                              "C.姓名||Decode(A.婴儿,0,'',Null,'','(婴儿)') As 姓名," & _
                              "C.门诊号," & _
                              "C.住院号," & _
                              "D.名称 AS 申请科室," & _
                              "A.开嘱医生 AS 申请人," & _
                              "f.接收人,f.接收时间, " & _
                              "'Item' AS 图标,Decode(Nvl(z.医嘱序号,0), 0, '√', Decode(Sum(Nvl(z.数次,0)), 0, '×', '√')) As 收费,NVL(A.紧急标志,0) AS 紧急,Y.操作类型,Decode(H.项目类别,2,2,1) As 项目类别,F.采样人,F.采样时间 As 采样时间 " & _
                         "FROM 病人医嘱记录 A," & _
                         "病人信息 C,部门表 D,病人医嘱发送 F,检验报告项目 G,检验项目 H,诊疗项目目录 Y,检验标本记录 j,住院费用记录 Z " & _
                        "WHERE A.诊疗类别 = 'C' " & _
                              "AND A.病人ID=C.病人ID " & _
                              "AND A.开嘱科室ID=D.ID " & _
                              "AND A.相关id IS NOT NULL " & _
                              "AND A.医嘱状态=8 AND A.ID=F.医嘱id " & _
                              "AND A.诊疗项目id=G.诊疗项目id AND G.细菌ID Is Null " & _
                              "AND G.报告项目ID=H.诊治项目ID " & _
                              "AND A.诊疗项目ID=Y.ID " & _
                              "And a.病人来源 = 2 " & _
                              IIf(mbln划价单模式 = True, " and a.病人来源 = 2 and c.出院时间 is null ", " ")

            mstrSql = mstrSql & " " & _
                              "AND A.病人ID=[1] AND A.执行科室id+0=[3] And a.相关id = j.医嘱id(+) and j.id = [8] " & _
                              "AND F.NO=Z.NO(+) AND F.记录性质=mod(Z.记录性质(+),10) AND F.医嘱id=Z.医嘱序号(+)+0 " & _
                              IIf(mblnBarCode, " And F.样本条码=[2]   ", "") & _
                              IIf(int主页ID = 0, " ", " And a.主页ID = [9] ") & _
                              IIf(blnNoRange, " ", " AND A.开嘱时间 BETWEEN [5] and [6] ") & _
                              IIf(mblnPrice = True, " And z.No = [10] ", " ") & _
                              " GROUP BY A.相关ID,a.id,C.姓名||Decode(A.婴儿,0,'',Null,'','(婴儿)'),C.门诊号,C.住院号,D.名称,A.开嘱医生,'Item',NVL(A.紧急标志,0),Y.操作类型,f.接收人,f.接收时间,Decode(H.项目类别,2,2,1),F.采样人,F.采样时间 ,z.医嘱序号 "
        End If
    Else
        If blnCheck Then '显示收费
            If mblnBarCode Then
                mstrSql = mstrSql & "SELECT Decode(SUM(Decode(F.样本条码,[2],1,0)),0,0,1) AS 选择"
            Else
                mstrSql = mstrSql & "SELECT Decode(Nvl(z.医嘱序号,0), 0, 1, Decode(Sum(Nvl(z.数次,0)), 0, 0, 1)) As 选择"
            End If
            mstrSql = mstrSql & ",A.相关ID AS ID, " & _
                              "C.姓名||Decode(A.婴儿,0,'',Null,'','(婴儿)') As 姓名," & _
                              "C.门诊号," & _
                              "C.住院号," & _
                              "D.名称 AS 申请科室," & _
                              "A.开嘱医生 AS 申请人," & _
                              "f.接收人,f.接收时间, " & _
                              "'Item' AS 图标,Decode(Nvl(z.医嘱序号,0), 0, '√', Decode(Sum(Nvl(z.数次,0)), 0, '×', '√')) As 收费,NVL(A.紧急标志,0) AS 紧急,H.操作类型,MAX(Decode(I.项目类别,2,2,1)) As 项目类别,MAX(F.采样人) AS 采样人,MAX(F.采样时间) AS 采样时间 " & _
                         "FROM 病人医嘱记录 A," & _
                         "病人信息 C,部门表 D,病人医嘱发送 F,检验报告项目 G,诊疗项目目录 H,检验项目 I,检验仪器项目 Y,住院费用记录 Z " & _
                        "WHERE A.诊疗类别 = 'C' " & _
                              "AND A.病人ID=C.病人ID " & _
                              "AND A.开嘱科室ID=D.ID " & _
                              "AND A.相关id IS NOT NULL " & _
                              "AND A.医嘱状态=8 AND A.ID=F.医嘱id " & _
                              "AND A.诊疗项目id=G.诊疗项目id AND G.细菌ID Is Null " & _
                              IIf(blMachineFind, "AND G.报告项目id=Y.项目id ", "AND G.报告项目id=Y.项目id(+) ") & _
                              "AND G.报告项目ID=I.诊治项目ID " & _
                              "AND A.诊疗项目ID=H.ID " & _
                              IIf(mbln划价单模式 = True, " and a.病人来源 = 2 and c.出院时间 is null ", " ") & _
                              IIf(mlngDefaultDevice = 0 And lngCurrDevice = 0, "", "AND (Y.仪器ID+0=[7] Or Y.仪器ID Is Null)") & _
                              "AND F.执行状态 = 0 AND A.病人ID=[1] AND A.执行科室id+0=[3] " & _
                              IIf(mblnBarCode, " And F.样本条码=[2]   ", "")

            mstrSql = mstrSql & " " & _
                              IIf(int主页ID = 0, " ", " And a.主页ID = [9] ") & _
                              IIf(blnNoRange, " ", " AND A.开嘱时间 BETWEEN [5] and [6] ") & _
                              "AND F.NO=Z.NO(+) AND F.记录性质=mod(Z.记录性质(+),10) AND F.医嘱id=Z.医嘱序号(+)+0 " & _
                              "And a.病人来源 = 2 " & _
                              IIf(mblnPrice = True, " And z.No = [10] ", " ") & _
                              "GROUP BY A.相关ID,a.id,C.姓名||Decode(A.婴儿,0,'',Null,'','(婴儿)'),C.门诊号,C.住院号,D.名称,A.开嘱医生,'Item',NVL(A.紧急标志,0),H.操作类型,f.接收人,f.接收时间 ,z.医嘱序号 "

            mstrSql = mstrSql & " Union all "

            If mblnBarCode Then
                mstrSql = mstrSql & "SELECT Decode(SUM(Decode(F.样本条码,[2],1,0)),0,0,1) AS 选择"
            Else
                mstrSql = mstrSql & "SELECT Decode(Nvl(z.医嘱序号,0), 0, 1, Decode(Sum(Nvl(z.数次,0)), 0, 0, 1)) As 选择"
            End If

            mstrSql = mstrSql & ",A.相关ID AS ID, " & _
                              "C.姓名||Decode(A.婴儿,0,'',Null,'','(婴儿)') As 姓名," & _
                              "C.门诊号," & _
                              "C.住院号," & _
                              "D.名称 AS 申请科室," & _
                              "A.开嘱医生 AS 申请人," & _
                              "f.接收人,f.接收时间, " & _
                              "'Item' AS 图标,Decode(Nvl(z.医嘱序号,0), Null, '√', Decode(Sum(Nvl(z.数次,0)), 0, '×', '√')) As 收费,NVL(A.紧急标志,0) AS 紧急,H.操作类型,MAX(Decode(I.项目类别,2,2,1)) As 项目类别,MAX(F.采样人) AS 采样人,MAX(F.采样时间) AS 采样时间 " & _
                         "FROM 病人医嘱记录 A," & _
                         "病人信息 C,部门表 D,病人医嘱发送 F,检验报告项目 G,诊疗项目目录 H,检验项目 I,检验仪器项目 Y,住院费用记录 Z,检验标本记录 j,检验项目分布 k " & _
                        "WHERE A.诊疗类别 = 'C' " & _
                              "AND A.病人ID=C.病人ID " & _
                              "AND A.开嘱科室ID=D.ID " & _
                              "AND A.相关id IS NOT NULL " & _
                              "AND A.医嘱状态=8 AND A.ID=F.医嘱id " & _
                              "AND A.诊疗项目id=G.诊疗项目id AND G.细菌ID Is Null " & _
                              IIf(blMachineFind, "AND G.报告项目id=Y.项目id ", "AND G.报告项目id=Y.项目id(+) ") & _
                              "AND G.报告项目ID=I.诊治项目ID " & _
                              "AND A.诊疗项目ID=H.ID " & _
                              IIf(mbln划价单模式 = True, " and a.病人来源 = 2 and c.出院时间 is null ", " ") & _
                              IIf(mlngDefaultDevice = 0 And lngCurrDevice = 0, "", "AND (Y.仪器ID+0=[7] Or Y.仪器ID Is Null)") & _
                              "AND A.病人ID=[1] AND A.执行科室id+0=[3] " & _
                              IIf(mblnBarCode, " And F.样本条码=[2]   ", "")

            mstrSql = mstrSql & " " & _
                              IIf(int主页ID = 0, " ", " And a.主页ID = [9] ") & _
                              IIf(blnNoRange, " ", " AND A.开嘱时间 BETWEEN [5] and [6] ") & _
                              "AND F.NO=Z.NO(+) AND F.记录性质=mod(Z.记录性质(+),10) AND F.医嘱id=Z.医嘱序号(+)+0 and a.相关id = k.医嘱id(+) and j.id =k.标本ID and j.id = [8] " & _
                              "And a.病人来源 = 2 " & _
                              IIf(mblnPrice = True, " And z.No = [10] ", " ") & _
                              "GROUP BY A.相关ID,a.id,C.姓名||Decode(A.婴儿,0,'',Null,'','(婴儿)'),C.门诊号,C.住院号,D.名称,A.开嘱医生,'Item',NVL(A.紧急标志,0),H.操作类型,f.接收人,f.接收时间 ,z.医嘱序号 "
        Else
            If mblnBarCode Then
                mstrSql = mstrSql & "SELECT Distinct 1 AS 选择"
            Else
                mstrSql = mstrSql & "SELECT Distinct 1 AS 选择"
            End If
            mstrSql = mstrSql & ",A.相关ID AS ID, " & _
                              "C.姓名||Decode(A.婴儿,0,'',Null,'','(婴儿)') As 姓名," & _
                              "C.门诊号," & _
                              "C.住院号," & _
                              "D.名称 AS 申请科室," & _
                              "A.开嘱医生 AS 申请人," & _
                              "f.接收人,f.接收时间, " & _
                              "'Item' AS 图标,Decode(Nvl(z.医嘱序号,0), 0, '√', Decode(Sum(Nvl(z.数次,0)), 0, '×', '√')) As 收费,NVL(A.紧急标志,0) AS 紧急,H.操作类型,Decode(I.项目类别,2,2,1) As 项目类别,F.采样人,F.采样时间 As 采样时间 " & _
                         "FROM 病人医嘱记录 A," & _
                         "病人信息 C,部门表 D,病人医嘱发送 F,检验报告项目 G,诊疗项目目录 H,检验项目 I,检验仪器项目 Y,住院费用记录 Z " & _
                        "WHERE A.诊疗类别 = 'C' " & _
                              "AND A.病人ID=C.病人ID " & _
                              "AND A.开嘱科室ID=D.ID " & _
                              "AND A.相关id IS NOT NULL " & _
                              "AND A.医嘱状态=8 AND A.ID=F.医嘱id " & _
                              "AND A.诊疗项目id=G.诊疗项目id AND G.细菌ID Is Null " & _
                              IIf(blMachineFind, "AND G.报告项目id=Y.项目id ", "AND G.报告项目id=Y.项目id(+) ") & _
                              "AND G.报告项目ID=I.诊治项目ID " & _
                              "AND A.诊疗项目ID=H.ID " & _
                              "And a.病人来源 = 2 " & _
                              IIf(mbln划价单模式 = True, " and a.病人来源 = 2 and c.出院时间 is null ", " ")

            mstrSql = mstrSql & " " & _
                              IIf(mlngDefaultDevice = 0 And lngCurrDevice = 0, "", "AND (Y.仪器ID+0=[7] Or Y.仪器ID Is Null)") & _
                              "AND F.执行状态 = 0 AND A.病人ID=[1] AND A.执行科室id+0=[3] " & _
                              "AND F.NO=Z.NO(+) AND F.记录性质=mod(Z.记录性质(+),10) AND F.医嘱id=Z.医嘱序号(+)+0 " & _
                              IIf(mblnBarCode, " And F.样本条码=[2]   ", "") & _
                              IIf(int主页ID = 0, " ", " And a.主页ID = [9] ") & _
                              IIf(blnNoRange, " ", " AND A.开嘱时间 BETWEEN [5] and [6] ") & _
                              IIf(mblnPrice = True, " And z.No = [10] ", " ") & _
                              " GROUP BY A.相关ID,a.id,C.姓名||Decode(A.婴儿,0,'',Null,'','(婴儿)'),C.门诊号,C.住院号,D.名称,A.开嘱医生,'Item',NVL(A.紧急标志,0),H.操作类型,f.接收人,f.接收时间,Decode(I.项目类别,2,2,1),F.采样人,F.采样时间 ,z.医嘱序号 "

            mstrSql = mstrSql & " Union all "

            If mblnBarCode Then
                mstrSql = mstrSql & "SELECT Distinct 1 AS 选择"
            Else
                mstrSql = mstrSql & "SELECT Distinct 1 AS 选择"
            End If

            mstrSql = mstrSql & ",A.相关ID AS ID, " & _
                              "C.姓名||Decode(A.婴儿,0,'',Null,'','(婴儿)') As 姓名," & _
                              "C.门诊号," & _
                              "C.住院号," & _
                              "D.名称 AS 申请科室," & _
                              "A.开嘱医生 AS 申请人," & _
                              "f.接收人,f.接收时间, " & _
                              "'Item' AS 图标,Decode(Nvl(z.医嘱序号,0), 0, '√', Decode(Sum(Nvl(z.数次,0)), 0, '×', '√')) As 收费,NVL(A.紧急标志,0) AS 紧急,H.操作类型,Decode(I.项目类别,2,2,1) As 项目类别,F.采样人,F.采样时间 As 采样时间 " & _
                         "FROM 病人医嘱记录 A," & _
                         "病人信息 C,部门表 D,病人医嘱发送 F,检验报告项目 G,诊疗项目目录 H,检验项目 I,检验仪器项目 Y,检验标本记录 j,住院费用记录 Z " & _
                        "WHERE A.诊疗类别 = 'C' " & _
                              "AND A.病人ID=C.病人ID " & _
                              "AND A.开嘱科室ID=D.ID " & _
                              "AND A.相关id IS NOT NULL " & _
                              "AND A.医嘱状态=8 AND A.ID=F.医嘱id " & _
                              "AND A.诊疗项目id=G.诊疗项目id AND G.细菌ID Is Null " & _
                              IIf(blMachineFind, "AND G.报告项目id=Y.项目id ", "AND G.报告项目id=Y.项目id(+) ") & _
                              "AND G.报告项目ID=I.诊治项目ID " & _
                              "AND A.诊疗项目ID=H.ID " & _
                              "And a.病人来源 = 2 " & _
                              IIf(mbln划价单模式 = True, " and a.病人来源 = 2 and c.出院时间 is null ", " ")

            mstrSql = mstrSql & " " & _
                              IIf(mlngDefaultDevice = 0 And lngCurrDevice = 0, "", "AND (Y.仪器ID+0=[7] Or Y.仪器ID Is Null)") & _
                              "AND A.病人ID=[1] AND A.执行科室id+0=[3] and a.相关id = j.医嘱id(+) and j.id = [8] " & _
                              "AND F.NO=Z.NO(+) AND F.记录性质=mod(Z.记录性质(+),10) AND F.医嘱id=Z.医嘱序号(+)+0 " & _
                              IIf(mblnBarCode, " And F.样本条码=[2]   ", "") & _
                              IIf(int主页ID = 0, " ", " And a.主页ID = [9] ") & _
                              IIf(blnNoRange, " ", " AND A.开嘱时间 BETWEEN [5] and [6] ") & _
                              IIf(mblnPrice = True, " And z.No = [10] ", " ") & _
                              " GROUP BY A.相关ID,a.id,C.姓名||Decode(A.婴儿,0,'',Null,'','(婴儿)'),C.门诊号,C.住院号,D.名称,A.开嘱医生,'Item',NVL(A.紧急标志,0),H.操作类型,f.接收人,f.接收时间,Decode(I.项目类别,2,2,1),F.采样人,F.采样时间 ,z.医嘱序号 "
        End If
    End If
    mstrSql = "Select distinct A.*,TO_CHAR(B.开嘱时间,'YY-MM-DD HH24:MI') AS 申请时间," & _
        "B.医嘱内容 AS 申请项目,B.开嘱科室ID,B.开嘱医生,nvl(b.主页ID,0) as 主页ID,a.收费, " & _
        " B.病人来源,C.名称 as 病人科室 " & _
        " From (" & mstrSql & ") A,病人医嘱记录 B,部门表 C " & _
        " Where A.ID=B.ID And B.病人科室ID = C.id  "

    strSQLbak = mstrSql
    strSQLbak = Replace(strSQLbak, "住院费用记录", "门诊费用记录")
    strSQLbak = Replace(strSQLbak, " and a.病人来源 = 2 and c.出院时间 is null ", " And a.病人来源 <> 2  and nvl(费用状态,0) <> 1 ")
    strSQLbak = Replace(strSQLbak, " And a.病人来源 = 2 ", " And a.病人来源 <> 2 and nvl(费用状态,0) <> 1 ")
    mstrSql = mstrSql & " Union ALL " & strSQLbak

    mstrSql = mstrSql & " order by 主页ID Desc,申请时间 "

    rs.CursorLocation = adUseClient
    Set rs = zlDatabase.OpenSQLRecord(mstrSql, _
        Me.Caption, mlng病人ID, strText, ItemDeptID, "", _
            CDate(Format(strStart, "yyyy-MM-dd hh:mm:ss")), CDate(Format(strEnd, "yyyy-MM-dd hh:mm:ss")), _
            IIf(mlngDefaultDevice > 0, mlngDefaultDevice, lngCurrDevice), mlngSampleID, int主页ID, mstrNO)
    If rs.RecordCount > 0 Then
        If Nvl(rs("病人来源"), 0) = 2 Then
            strSQL = "select nvl(max(主页ID),0) as 主页ID,病人性质 from 病案主页 where 病人id = [1]  group by 主页id ,病人性质 order by  主页id  desc"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID)
            intPatProperty = Val(rsTmp("病人性质"))
        End If
    End If
    If intPatProperty = 1 Then
        If bln未收费显示 = False Then
            rs.MoveFirst
            Do Until rs.EOF
                If Nvl(rs("病人来源"), 0) = 2 And Trim(Nvl(rs("收费"))) = "×" Then
                    mstrSql = Replace(mstrSql, "住院费用记录", "门诊费用记录")
                    Set rsPatProperty = zlDatabase.OpenSQLRecord(mstrSql, _
                    Me.Caption, mlng病人ID, strText, ItemDeptID, "", _
                        CDate(Format(strStart, "yyyy-MM-dd hh:mm:ss")), CDate(Format(strEnd, "yyyy-MM-dd hh:mm:ss")), _
                        IIf(mlngDefaultDevice > 0, mlngDefaultDevice, lngCurrDevice), mlngSampleID, int主页ID, mstrNO)
                    If rsPatProperty.RecordCount > 0 Then
                        Set rs = rsPatProperty
                        Exit Do
                    End If
                End If
                rs.MoveNext
            Loop
        End If
    End If
    If rs.BOF Then
        If mblnBarCode = True Then
            '扫条码的项目给出具体的提示
            gstrSql = "select 执行状态,a.病人ID,a.执行科室ID,c.名称 as 执行科室名称 from 病人医嘱记录 a,病人医嘱发送 b,部门表 c " & vbNewLine & _
                      " where a.id = b.医嘱id and a.相关id is not null and a.执行科室ID = c.id and b.样本条码 = [1]   "
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strText)
            If rsTmp.EOF = True Then
                MsgBox "没有找到条码<" & strText & ">!"
            Else
                str执行科室名称 = rsTmp("执行科室名称")
                lng执行科室ID = rsTmp("执行科室ID")
                lng病人ID = rsTmp("病人ID")
                int执行状态 = rsTmp("执行状态")
                gstrSql = "select 核收时间,核收人,审核时间,审核人,标本序号,仪器ID from 检验标本记录 where 病人id = [1] and 样本条码  = [2] "
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lng病人ID, strText)
                If rsTmp.EOF = False Then
                    If int执行状态 = 1 Then
                        MsgBox "条码<" & strText & ">已被" & rsTmp("审核人") & "在" & rsTmp("审核时间") & "审核,标本号" & _
                            TransSampleNO_PH(rsTmp("标本序号"), Nvl(rsTmp("仪器ID"), -1)) & "."
                    ElseIf int执行状态 = 2 Then
                        MsgBox "条码<" & strText & ">已拒收!"
                    ElseIf int执行状态 = 3 Then
                        MsgBox "条码<" & strText & ">已被" & rsTmp("核收人") & "在" & rsTmp("核收时间") & "核收,标本号" & _
                            TransSampleNO_PH(rsTmp("标本序号"), Nvl(rsTmp("仪器ID"), -1)) & "."
                    End If
                Else

                    If int执行状态 = 1 Then
                        MsgBox "条码<" & strText & ">已完成!"
                    ElseIf int执行状态 = 2 Then
                        MsgBox "条码<" & strText & ">已拒收!"
                    ElseIf int执行状态 = 3 Then
                        MsgBox "条码<" & strText & ">已核收!"
                    Else
                        If ItemDeptID <> lng执行科室ID Then
                            MsgBox "条码<" & strText & ">" & "的执行科室是<" & str执行科室名称 & ">不是当前选择的科室不能核收！"
                        End If
                    End If
                End If
            End If
        End If
        OpenSelect = 0
        Exit Function
    End If

    '住院病人只处理本次住院的医嘱
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        If Nvl(rs("病人来源"), 0) = 2 Then
            int主页ID = Val(Nvl(rsTmp("主页ID")))
        End If
    End If
    strTmp = ""
    Do Until rs.EOF
        If Nvl(rs("病人来源")) = 2 Then
            If int主页ID = Val(Nvl(rs("主页ID"))) Then
                strTmp = strTmp & " or ID=" & Val(Nvl(rs("ID")))
            End If
        Else
            strTmp = strTmp & " or ID=" & Val(Nvl(rs("ID")))
        End If
        rs.MoveNext
    Loop
    If strTmp <> "" Then
        rs.filter = "ID=-1" & strTmp
    End If

    If (rs.RecordCount = 1 And blnWhere) Or mblnBarCode Then GoTo Over
    
    If rs.RecordCount > 0 Then
        '不显示未收费的标本
        If bln未收费显示 = False Then
            strFilter = "收费 <> '×'"
            rs.filter = strFilter
        End If
    End If
    
    Call ClientToScreen(txt姓名.hWnd, objPoint)
    If frmSelectMuli.ShowSelectSP(Me, rs, strLvw, objPoint.X * 15 - 30, objPoint.Y * 15 + txt姓名.Height - 30, _
        8000, 5600, Me.Name & "\待核收标本选择", "请从下表中钩选需一次核收的标本") Then
        GoTo Over
    End If
    Exit Function

Over:
    If rs.EOF Then Exit Function

    '对急诊标志进行判断

    rs.MoveFirst
    Do Until rs.EOF
        If rs("紧急") = 1 Then
            mbln急诊 = True
            Exit Do
        End If
        rs.MoveNext
    Loop

    '对没有收费的病人进行判断
    If bln未收费显示 = False Then
        rs.MoveFirst
        Do Until rs.EOF

            '住院
            If Nvl(rs("病人来源"), 0) = 2 And Trim(Nvl(rs("收费"))) = "×" Then
                MsgBox "核收项目<" & Nvl(rs("申请项目")) & ">有未收费项目或退费项目不能核收", vbInformation, "核收提示"
                OpenSelect = 3
                Exit Function
            End If

            '门诊
            If Nvl(rs("病人来源"), 0) = 1 And Trim(Nvl(rs("收费"))) = "×" Then
                MsgBox "核收项目<" & Nvl(rs("申请项目")) & ">有未收费项目或退费项目不能核收", vbInformation, "核收提示"
                OpenSelect = 3
                Exit Function
            End If

            '体检病人只判断退费
            If Nvl(rs("病人来源"), 0) = 4 And Trim(Nvl(rs("收费"))) = "×" Then
                MsgBox "核收项目<" & Nvl(rs("申请项目")) & ">有未收费项目或退费项目不能核收", vbInformation, "核收提示"
                OpenSelect = 3
                Exit Function
            End If
            rs.MoveNext
        Loop
    End If

    '判断是否超过送检时间限
    rs.MoveFirst
    Do Until rs.EOF
        If (IsDate(Nvl(rs("采样时间"))) = True And Nvl(rs("采样人")) <> "") Then
            gstrSql = "Select Min(送检时限) As 送检时限" & vbNewLine & _
                        "From 病人医嘱记录 A, 检验项目选项 B" & vbNewLine & _
                        "Where A.诊疗项目id = B.诊疗项目id And A.相关id = [1] And Nvl(送检时限, 0) > 0"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(rs("ID")))
            If rsTmp.EOF = False Then
                If Val(Nvl(rsTmp("送检时限"))) > 0 Then
                    strDate = zlDatabase.Currentdate
                    If DateDiff("n", Nvl(rs("采样时间")), strDate) > Val(Nvl(rsTmp("送检时限"))) And Val(Nvl(rsTmp("送检时限"))) > 0 Then
                        strTmp = DateDiff("n", Nvl(rs("采样时间")), strDate)
                        If MsgBox("核收项目超过了送检时限（" & strTmp & "分钟),是否续继?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                            OpenSelect = 3
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
        rs.MoveNext
    Loop

    '项目相同，给予提示
    rs.MoveFirst
    strTmp = ""
    Do While Not rs.EOF
        If Val(rs("选择")) = 1 Then
            If InStr("," & strTmp & ",", "," & Trim(rs("申请项目")) & ",") > 0 Then
                If MsgBox("选择了多个相同的项目，是否继续核收？ ", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                    OpenSelect = 3
                    Exit Function
                Else
                    Exit Do
                End If
            Else
                strTmp = strTmp & "," & Trim(rs("申请项目"))
            End If
        End If
        rs.MoveNext
    Loop

    '病人来源不一制时不能一起核收
    rs.MoveFirst
    strTmp = ""
    Do Until rs.EOF
        If strTmp = "" Then strTmp = Trim(Nvl(rs("病人来源")))
        If Trim(strTmp) <> Trim(Nvl(rs("病人来源"))) Then
            MsgBox "不同的来源病人医嘱不一起核收", vbInformation, "核收提示"
            OpenSelect = 3
            Exit Function
        End If
        rs.MoveNext
    Loop

    '母亲的医嘱和子女的医嘱不能在一起核收
    rs.MoveFirst
    strTmp = ""
    Do Until rs.EOF
        If strTmp = "" Then strTmp = Nvl(rs("姓名"))
        If Trim(strTmp) <> Trim(Nvl(rs("姓名"))) Then
            MsgBox "母亲的医嘱不能和子女的医嘱在一起核收！", vbInformation, "核收提示"
            OpenSelect = 3
            Exit Function
        End If
        rs.MoveNext
    Loop

    '不同的标本不能在一起核收
'    rs.MoveFirst
'    strTmp = ""
'    Do Until rs.EOF
'        If strTmp = "" Then strTmp = Nvl(rs("标本类型"))
'        If Trim(strTmp) <> Trim(Nvl(rs("标本类型"))) Then
'            MsgBox "不同的标本类型不能在一起核收！", vbInformation, "核收提示"
'            OpenSelect = 3
'            Exit Function
'        End If
'        rs.MoveNext
'    Loop

    rs.MoveFirst
    If InStr(Nvl(rs("姓名")), "(婴儿)") > 0 Then
        gstrSql = "Select B.病人id, B.主页id, B.序号, B.婴儿姓名, B.婴儿性别" & vbNewLine & _
                    "From 病人医嘱记录 A, 病人新生儿记录 B" & vbNewLine & _
                    "Where A.病人id = B.病人id And A.主页id = B.主页id And A.婴儿 = B.序号 And A.相关id = [1] And Rownum = 1"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(Nvl(rs("ID"), 0)))
        If rsTmp.EOF = False Then
            Me.txt年龄 = ""
            Me.cboAge = "婴儿"
            txt姓名.Text = Nvl(rsTmp("婴儿姓名"))
            strTmp = "名称='" & CStr(Nvl(rsTmp("婴儿性别"))) & "'"
            mRsSex.filter = strTmp
            If mRsSex.EOF = False Then
                Me.cbo性别.Text = mRsSex!编码 & "-" & mRsSex!名称
            End If
        End If
    Else
        On Error Resume Next
        gstrSql = "select  年龄 from 病人医嘱记录 where 相关id=[1] And Rownum = 1"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(Nvl(rs("ID"), 0)))
        
        strAge = Nvl(rsTmp("年龄"))
        
        strAge = Replace(strAge, "小时", "时")
        strAge = Replace(strAge, "分钟", "分")
        
        If Trim(Replace(Replace(Replace(Replace(Replace(strAge, "岁", ""), "月", ""), "天", ""), "时", ""), "分", "")) <> "" Then
            If InStr(strAge, "成人") > 0 Or InStr(strAge, "婴儿") > 0 Then
                Me.txt年龄.Text = ""
                Me.cboAge.Text = Trim(strAge)
            Else
                strAge = Replace(Replace(Replace(Replace(Replace(strAge, "岁", "岁;"), "月", "月;"), "天", "天;"), "时", "时;"), "分", "分;")
                'strAge = Replace(strAge, "分钟", "婴儿")
                aAge = Split(strAge, ";")
                If UBound(aAge) = 1 Then
                    Me.txt年龄.Text = Val(aAge(0))
                    Me.cboAge.Text = Replace(Replace(Right(aAge(0), 1), "分", "分钟"), "时", "小时")
                Else
                    Me.txt年龄.Text = Val(aAge(0))
                    Me.cboAge.Text = Replace(Replace(Right(aAge(0), 1), "分", "分钟"), "时", "小时")
                    Me.txt年龄1.Text = Val(aAge(1)) & Replace(Replace(Right(aAge(1), 1), "分", "分钟"), "时", "小时")
                End If
            End If
        Else
            If Val(strAge) <> 0 Then
                Me.txt年龄.Text = Val(strAge)
            End If
            Me.cboAge.ListIndex = 0
        End If
    End If
    On Error GoTo ErrHand
    If InStr(txt姓名, "(婴儿)") > 0 Then Me.txt年龄 = ""

    DTP(0).Value = Format(zlCommFun.Nvl(rs("采样时间"), zlDatabase.Currentdate), "YYYY-MM-DD HH:MM:SS")
    mbln微生物项目 = (zlCommFun.Nvl(rs("项目类别"), 1) = 2)
    If Nvl(rs("接收人"), "") = "" Then
        cbo(2).Visible = False
        DTP(2).Visible = False
        lbl(1).Visible = False
    Else
        cbo(2).Visible = True
        DTP(2).Visible = True
        lbl(1).Visible = True
        cbo(2).Text = zlCommFun.Nvl(rs("接收人"))
        'zlControl.CboLocate cbo(2), zlCommFun.Nvl(rs("接收人"))
        DTP(2).Value = Format(zlCommFun.Nvl(rs("接收时间"), zlDatabase.Currentdate), "YYYY-MM-DD HH:MM:SS")
    End If
    txtPatientDept = Nvl(rs("病人科室"))
    Me.cbo开单科室.ListIndex = FindComboItem(Me.cbo开单科室, Nvl(rs("开嘱科室ID")))
    cbo(1).Text = zlCommFun.Nvl(rs("采样人"))
    '没有采样人时不显示采样时间
    If Trim(Me.cbo(1).Text) = "" Then
        Me.cbo(1).Visible = False
        Me.DTP(0).Visible = False
        lbl(0).Visible = False
    Else
        Me.cbo(1).Visible = True
        Me.DTP(0).Visible = True
        lbl(0).Visible = True
    End If
    Me.txtPatientDept.Text = Nvl(rs("病人科室"))
    Select Case Nvl(Nvl(rs("病人来源")))
        Case 1, 3, 4
            txtID.Text = Nvl(rs("门诊号"))
            txtBed.Text = ""
        Case 2
            txtID.Text = Nvl(rs("住院号"))
    End Select
'    zlControl.CboLocate cbo(1), zlCommFun.Nvl(rs("采样人"))
    On Error Resume Next
    strField = ""
    strField = rs.Fields("开嘱医生").Name
    If strField = "开嘱医生" Then
        Me.cbo医生.Text = Nvl(rs("开嘱医生"))
        For i = 0 To Me.cbo医生.ListCount - 1
            If Me.cbo医生.List(i) Like Nvl(rs("开嘱医生")) Then
                Me.cbo医生.ListIndex = i
                Exit For
            End If
        Next
    End If

    gstrSql = "select 标本部位 AS 标本类型 from 病人医嘱记录 where 相关id = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(rs("ID")))
    If rsTmp.EOF = False Then
        Me.txt附加.Text = Nvl(rsTmp("标本类型"))
    End If

    On Error GoTo ErrHand

    If rs.RecordCount = 1 Then
        mstrKeys = zlCommFun.Nvl(rs("ID").Value)

        Me.txt医嘱内容 = rs("申请项目")
        Me.txt医嘱内容.Tag = rs("申请项目")

        If blnCheck Then
'            Me.lblCash.Font.Strikethrough = Not (rs("收费") = "√")
            Me.lblCash.Caption = IIf((rs("收费") = "√"), "收", "")
        Else
'            Me.lblCash.Font.Strikethrough = True
            Me.lblCash.Caption = ""
        End If
    Else
        strLastKeys = mstrKeys: mstrKeys = "": Me.txt医嘱内容 = ""
        Do While Not rs.EOF
            mstrKeys = mstrKeys & "," & zlCommFun.Nvl(rs("ID").Value)
'            If InStr("," & txt医嘱内容 & ",", "," & zlCommFun.Nvl(rs("申请项目").Value & ",")) <= 0 Then
                txt医嘱内容 = txt医嘱内容 & "," & zlCommFun.Nvl(rs("申请项目").Value)
'            End If

            If blnCheck Then
'                Me.lblCash.Font.Strikethrough = Not (rs("收费") = "√")
                Me.lblCash.Caption = IIf((rs("收费") = "√"), "收", "")
            Else
'                Me.lblCash.Font.Strikethrough = True
                Me.lblCash.Caption = ""
            End If

            rs.MoveNext
        Loop
        If mstrKeys = "" Then
            mstrKeys = strLastKeys
        Else
            mstrKeys = Mid(mstrKeys, 2)
            txt医嘱内容 = Mid(txt医嘱内容, 2)
            txt医嘱内容.Tag = Mid(txt医嘱内容, 2)
        End If
    End If

    OpenSelect = 1

    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CalcNextCode(ByVal lngKey As Long, ByVal intRow As Integer, ByVal iType As Integer) As String
    '--------------------------------------------------------------------------------------------------------
    '功能:计算指定仪器在当天内的下一个缺省标本号
    '参数:lngKey                检验仪器ID
    '     iType                 标本类别：0=普通、1=急诊
    '返回:缺省标本号码
    '--------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strToday As String
    Dim strTmp As String
    Dim lng次数 As Long
    Dim strLabNo As String, strLabQCNo As String '检验标本、质控标本
    Dim mstrSql As String, mlngLoop As Long
    Dim strStartDate As String
    Dim strEndDate As String
    Dim lngDefaultItemID As Long
    Dim strItem As String
    Dim rsTmp As New ADODB.Recordset

    '时间,仪器,标本号
    On Error GoTo ErrHand

    strToday = Format(DTP(1).Value, "YYYY-MM-DD")
    strStartDate = GetDateTime(mMakeNoRule, 1, DTP(1).Value)
    strEndDate = GetDateTime(mMakeNoRule, 2, DTP(1).Value)

'    lngDefaultItemID = mlngDefaultItemID

    If mintItemRule = 1 Then
        If mstrKeys <> "" Then
            gstrSql = "select /*+ rule */ 诊疗项目ID from 病人医嘱记录 where 相关ID in  (Select * From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))) "
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mstrKeys)
            If rsTmp.EOF = False Then
                lngDefaultItemID = rsTmp("诊疗项目ID")
            End If
        Else
            If mstrExtData <> "" Then
                lngDefaultItemID = Val(mstrExtData)
            End If
        End If

    End If

    On Error GoTo point1

    mstrSql = "SELECT NVL(MAX(TO_NUMBER(标本序号)),0) AS 最大序号 FROM 检验标本记录 a,检验申请项目 b " & _
                "WHERE 核收时间 BETWEEN [2] and [3] And a.id = b.标本id(+) And nvl(a.是否质控品,0) = 0 " & _
                    IIf(lngKey = -1, " AND 仪器id IS NULL " & _
                        IIf(lngDefaultItemID > 0, " And b.诊疗项目id = [4] ", ""), "AND 仪器id= [1] ") & " And 医嘱ID Is Not Null" & _
                    IIf(mblnEmerge, IIf(iType = 1, " And 标本类别=1", " And Nvl(标本类别,0)<>1"), "")
    Set rs = zlDatabase.OpenSQLRecord(mstrSql, Me.Caption, lngKey, CDate(strStartDate), _
                           CDate(strEndDate), lngDefaultItemID)

    If Not rs.EOF Then strLabNo = zlCommFun.Nvl(rs("最大序号"))

    On Error GoTo ErrHand
    GoTo point2

point1:
    On Error GoTo ErrHand

    mstrSql = "SELECT NVL(MAX(标本序号),'') AS 最大序号 FROM 检验标本记录 a,检验申请项目 b " & _
                "WHERE 核收时间 BETWEEN [2] and [3] And a.id = b.标本id(+) And nvl(a.是否质控品,0) = 0  " & _
                    IIf(lngKey = -1, " AND 仪器id IS NULL " & _
                    IIf(lngDefaultItemID > 0, " And b.诊疗项目id = [4] ", ""), "AND 仪器id= [1] ") & " And 医嘱ID Is Not Null" & _
                    IIf(mblnEmerge, IIf(iType = 1, " And 标本类别=1", " And Nvl(标本类别,0)<>1"), "")
    Set rs = zlDatabase.OpenSQLRecord(mstrSql, Me.Caption, lngKey, CDate(strStartDate), _
                            CDate(strEndDate), lngDefaultItemID)

    If Not rs.EOF Then strLabNo = zlCommFun.Nvl(rs("最大序号"))

point2:
    On Error GoTo point3

    mstrSql = "SELECT NVL(MAX(TO_NUMBER(标本序号)),0) AS 最大序号 FROM 检验标本记录 a,检验申请项目 b " & _
                "WHERE 核收时间 BETWEEN [2] and [3] And a.id = b.标本ID(+) And nvl(a.是否质控品,0) = 0 " & _
                    IIf(lngKey = -1, " AND 仪器id IS NULL " & _
                    IIf(lngDefaultItemID > 0, " And b.诊疗项目id = [4] ", ""), "AND 仪器id= [1] ") & _
                    IIf(mblnEmerge, IIf(iType = 1, " And 标本类别=1", " And Nvl(标本类别,0)<>1"), "")
    Set rs = zlDatabase.OpenSQLRecord(mstrSql, Me.Caption, lngKey, CDate(strStartDate), _
                            CDate(strEndDate), lngDefaultItemID)

    If Not rs.EOF Then strLabQCNo = zlCommFun.Nvl(rs("最大序号"))

    On Error GoTo ErrHand
    GoTo point4

point3:
    On Error GoTo ErrHand

    mstrSql = "SELECT NVL(MAX(标本序号),'') AS 最大序号 FROM 检验标本记录 a,检验申请项目 b " & _
                "WHERE 核收时间 BETWEEN [2] and [3] And a.id = b.标本ID(+) And nvl(a.是否质控品,0) =  0 " & _
                    IIf(lngKey = -1, " AND 仪器id IS NULL " & _
                    IIf(lngDefaultItemID > 0, " And b.诊疗项目id = [4] ", ""), "AND 仪器id=[1] ") & _
                    IIf(mblnEmerge, IIf(iType = 1, " And 标本类别=1", " And Nvl(标本类别,0)<>1"), "")
    Set rs = zlDatabase.OpenSQLRecord(mstrSql, Me.Caption, lngKey, CDate(strStartDate), _
                            CDate(strEndDate), lngDefaultItemID)

    If Not rs.EOF Then strLabQCNo = zlCommFun.Nvl(rs("最大序号"))

point4:
    If Val(strLabNo) >= Val(strLabQCNo) Then
        CalcNextCode = strLabNo
    Else
        CalcNextCode = strLabQCNo
    End If
'    If Val(strLabQCNo) > Val(strLabNo) + 100 Then CalcNextCode = strLabNo

    For mlngLoop = 1 To vsf2.Rows - 1
        If mlngLoop <> intRow Then
            If Val(vsf2.RowData(mlngLoop)) = lngKey Then
                If Val(CalcNextCode) < Val(vsf2.TextMatrix(mlngLoop, 2)) Then
                    CalcNextCode = Val(vsf2.TextMatrix(mlngLoop, 2))
                End If
            End If
        End If
    Next

    If Val(CalcNextCode) <= 0 Then
        CalcNextCode = "1"
        Exit Function
    End If

    CalcNextCode = Val(CalcNextCode) + 1
    Exit Function

ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub LoadDefaultData()
    '--------------------------------------------------------------------------------------------------------
    '功能:产生检验仪器和标本号
    '--------------------------------------------------------------------------------------------------------
    Dim lngloop As Long
    Dim strNO As String
    Dim lngDefaultRec As Long
    Dim strConnectDevIDs As String, lngTmpNO As Long
    Dim strSubQry As String, mstrSql As String, mRs As New ADODB.Recordset, rsTmp As New ADODB.Recordset
    Dim lngCurrDevice As Long '项目的缺省仪器，主窗体未选当前仪器时有效
    Dim blnCurrDevice As Boolean '是否选择了默认仪器
    Dim intNow As Integer
    Dim lngItem As Long


    '获取本机连接的检验仪器
    strConnectDevIDs = GetConnectDevs

    On Error GoTo ErrHand
    '读取相应的检验仪器列表

    If mstrKeys <> "" Then
        If mbln微生物项目 = False Then
            mstrSql = "SELECT ID,名称,缺省仪器,MIN(相关ID) As 医嘱ID FROM " & _
                        "(SELECT DISTINCT NVL(E.ID,-1) AS ID,NVL(E.名称,'[手工]') AS 名称,NVL(D.缺省仪器,-1) AS 缺省仪器,A.相关ID " & _
                            "FROM 病人医嘱记录 A, 检验报告项目 B, 检验仪器项目 D, 检验仪器 E " & _
                            "Where A.诊疗项目ID+0 = B.诊疗项目ID(+) " & _
                            "AND B.报告项目ID = D.项目id(+) AND D.仪器id = E.ID(+) " & _
                            "AND A.病人ID=[1] AND Instr(','||[2]||',',','||A.相关ID||',')>0 " & _
                            "ORDER BY NVL(E.ID,-1)  DESC) " & _
                        "GROUP BY 缺省仪器,ID,名称"
        Else
            mstrSql = "SELECT ID,名称,缺省仪器,MIN(相关ID) As 医嘱ID FROM " & _
                        "(SELECT DISTINCT NVL(E.ID," & mlngDefaultDevice & ") AS ID, " & _
                        "NVL(E.名称," & " (select 名称 from 检验仪器 where id = [3]) " & " ) AS 名称, " & _
                        "NVL(1,-1) AS 缺省仪器,A.相关ID " & _
                            "FROM 病人医嘱记录 A, 检验报告项目 B, 仪器细菌对照 D, 检验仪器 E " & _
                            "Where A.诊疗项目ID+0 = B.诊疗项目ID(+) " & _
                            "AND B.细菌ID = D.细菌id(+) AND D.仪器id = E.ID(+) " & _
                            "AND A.病人ID=[1] AND Instr(','||[2]||',',','||A.相关ID||',')>0 " & _
                            "ORDER BY NVL(E.ID," & mlngDefaultDevice & ")  DESC) " & _
                        "GROUP BY 缺省仪器,ID,名称"
        End If
        Set mRs = zlDatabase.OpenSQLRecord(mstrSql, Me.Caption, mlng病人ID, mstrKeys, mlngDefaultDevice, mstrMachines)
    Else
        strSubQry = ""
        rsRelativeAdvice.MoveFirst
        Do While Not rsRelativeAdvice.EOF
            strSubQry = strSubQry & " Union All " & "Select " & rsRelativeAdvice("ID") & " As ID From Dual"

            rsRelativeAdvice.MoveNext
        Loop
        If Len(strSubQry) > 0 Then strSubQry = Mid(strSubQry, 12)
        rsRelativeAdvice.MoveFirst

        gstrSql = "Select A.诊疗项目id" & vbNewLine & _
                    "From 检验报告项目 A, 检验项目 B, (" & strSubQry & ") S" & vbNewLine & _
                    "Where S.ID = A.诊疗项目id And A.报告项目id = B.诊治项目id And B.项目类别 = 2"

        zlDatabase.OpenRecordset rsTmp, gstrSql, Me.Caption
        mbln微生物项目 = Not rsTmp.EOF

        If mbln微生物项目 = False Then
            gstrSql = "Select A.诊疗项目id" & vbNewLine & _
                    "From 检验报告项目 A, 检验项目 B, (" & strSubQry & ") S" & vbNewLine & _
                    "Where S.ID = A.诊疗项目id And A.报告项目id = B.诊治项目id And B.项目类别 <> 2 "
            zlDatabase.OpenRecordset rsTmp, gstrSql, Me.Caption
            If rsTmp.EOF = False Then
                mstrSql = "SELECT DISTINCT NVL(E.ID,-1) AS ID,NVL(E.名称,'[手工]') AS 名称,NVL(D.缺省仪器,-1) AS 缺省仪器 " & _
                                "FROM 检验报告项目 B, 检验仪器项目 D, 检验仪器 E,(" & strSubQry & ") S " & _
                                "Where S.ID=B.诊疗项目ID(+) " & _
                                "AND B.报告项目ID = D.项目id(+) AND D.仪器id = E.ID(+) " & _
                                "ORDER BY NVL(D.缺省仪器,-1)  DESC"
            Else
                mstrSql = "Select id,名称,1 as 缺省仪器 from 检验仪器 where id = [1] "
            End If
        Else
            mstrSql = "SELECT DISTINCT NVL(E.ID," & mlngDefaultDevice & ") AS ID, " & _
                            "NVL(E.名称," & " (select 名称 from 检验仪器 where id = [1]) " & ") AS 名称,-1 AS 缺省仪器 " & _
                            "FROM 检验报告项目 B, 仪器细菌对照 D, 检验仪器 E,(" & strSubQry & ") S " & _
                            "Where S.ID=B.诊疗项目ID(+) " & _
                            "AND B.细菌ID = D.细菌id(+) AND D.仪器id = E.ID(+)  "
        End If

        Set mRs = zlDatabase.OpenSQLRecord(mstrSql, gstrSysName, mlngDefaultDevice)
        '没有仪器时固定写和当前界面上的仪器ID
        If mRs.RecordCount <= 1 And mlngDefaultDevice > 0 Then
            If mRs("id") = -1 Then
                mstrSql = "select id , 名称 , 0 as 缺省仪器 from 检验仪器 a where id = [1] "
                Set mRs = zlDatabase.OpenSQLRecord(mstrSql, gstrSysName, mlngDefaultDevice)
            End If
        End If
    End If
    '如果主窗体未选当前仪器，则以项目的默认仪器为准
    If mlngDefaultDevice = 0 Then
        If mstrKeys <> "" Then
            mstrSql = "Select B.住院仪器ID,B.门诊仪器ID,B.住院仪器分解,B.门诊仪器分解 " & _
                "From 病人医嘱记录 A,检验项目选项 B Where A.诊疗项目ID=B.诊疗项目ID " & _
                "AND A.病人ID=[1] AND Instr(','||[2]||',',','||A.相关ID||',')>0"
            Set rsTmp = zlDatabase.OpenSQLRecord(mstrSql, Me.Caption, mlng病人ID, mstrKeys)
        Else
            mstrSql = "SELECT B.住院仪器ID,B.门诊仪器ID,B.住院仪器分解,B.门诊仪器分解 " & _
                "FROM 检验项目选项 B,(" & strSubQry & ") S " & _
                "Where S.ID=B.诊疗项目ID"
            Set rsTmp = zlDatabase.OpenSQLRecord(mstrSql, Me.Caption)
        End If
        If rsTmp.EOF Then
            lngCurrDevice = 0
        Else
            lngCurrDevice = IIf(PatientType = 2, Nvl(rsTmp("住院仪器ID"), 0), Nvl(rsTmp("门诊仪器ID"), 0))
        End If

    End If

    If mRs.BOF = False Then
        ResetVsf vsf2
        '如果一次增加所有标本，则增加N条空记录
        vsf2.Rows = mRs.RecordCount + 1

        For lngloop = 1 To vsf2.Rows - 1
            If Val(vsf2.RowData(lngloop)) = 0 Then

                '检验仪器是否已经使用,如已使用,则取一个仪器,如没有下一个,则取最后一个
                lngDefaultRec = -1: mRs.MoveFirst
                blnCurrDevice = False
                Do While Not mRs.EOF
                    If CheckHave(zlCommFun.Nvl(mRs("ID"), 0)) = False Then
                        If zlCommFun.Nvl(mRs("ID"), 0) = mlngDefaultDevice Then
                            '取过滤条件指定的检验仪器
                            lngDefaultRec = mRs.AbsolutePosition
                            Exit Do '不再继续查找
                        Else
                            If zlCommFun.Nvl(mRs("ID"), 0) = lngCurrDevice Then
                                lngDefaultRec = mRs.AbsolutePosition
                                blnCurrDevice = True
                            Else
                                If InStr(";" & strConnectDevIDs & ";", ";" & zlCommFun.Nvl(mRs("ID"), 0) & ";") > 0 Then
                                    '默认取本机连接的检验仪器
                                    If Not blnCurrDevice Then lngDefaultRec = mRs.AbsolutePosition
                                Else
                                    If lngDefaultRec = -1 Then lngDefaultRec = mRs.AbsolutePosition '先将当前仪器选中
                                End If
                            End If
                        End If
                    End If
                    mRs.MoveNext
                Loop
                If lngDefaultRec = -1 Then
                    mRs.MoveLast
                Else
                    mRs.AbsolutePosition = lngDefaultRec
                End If

                If mblnBarCode = False And mstr条码仪器 <> "" Then
                    If InStr("," & mstr条码仪器 & ",", "," & mRs("ID") & ",") > 0 Then
                        MsgBox "仪器<" & mRs("名称") & ">必须使用条码输入!", vbInformation, Me.Caption
                        Exit Sub
                    End If
                End If

                vsf2.TextMatrix(lngloop, 1) = zlCommFun.Nvl(mRs("名称"))
                vsf2.RowData(lngloop) = zlCommFun.Nvl(mRs("ID"), 0)

                If mstrKeys <> "" Then
                    vsf2.TextMatrix(lngloop, 4) = zlCommFun.Nvl(mRs("医嘱ID"), 0)
                    vsf2.TextMatrix(lngloop, 5) = IIf(mbln急诊 And mblnEmerge = True, "-1", "0")
                End If

                intNow = Val(zlDatabase.GetPara("按上次输入的标本号累加", 100, 1208, 0))

                If intNow = 1 And mstrNONumber <> "" Then
                    vsf2.TextMatrix(lngloop, 2) = TransSampleNO_PH(Val(mstrNONumber) + 1, vsf2.RowData(lngloop))
                Else
                    '取标本号
                    If vsf2.TextMatrix(lngloop, 5) = "-1" Then
                        '急诊
                        vsf2.TextMatrix(lngloop, 2) = TransSampleNO_PH(Val(CalcNextCode(Val(vsf2.RowData(lngloop)), lngloop, 1)), vsf2.RowData(lngloop))
                    Else
                        vsf2.TextMatrix(lngloop, 2) = TransSampleNO_PH(Val(CalcNextCode(Val(vsf2.RowData(lngloop)), lngloop, 0)), vsf2.RowData(lngloop))
                    End If
                End If
            End If
        Next
        vsf2.EditMode(1) = 1
        vsf2.ComboList(1) = "..."
    End If
    mbln急诊 = False
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SelectDefault()
    '--------------------------------------------------------------------------------------------------------
    '功能:将指标加入到仪器中
    '--------------------------------------------------------------------------------------------------------
    Dim iRow As Integer, iCurrRow As Integer
    Dim blnChkItem As Boolean '是否有检验指标
    Dim lngItemRow As Long
    Dim aItems() As Variant
    Dim astrKey() As String
    Dim intLoop As Integer
    Dim blnCheck As Boolean
    Dim lngloop As Long


    aItems = ReadData
    If UBound(aItems) = -1 Then Exit Sub
    If vsf2.Rows >= 2 And vsf2.RowData(1) = "" Then Exit Sub
    iCurrRow = vsf2.Row
    For iRow = 1 To vsf2.Rows - 1
        vsf2.Row = iRow
        blnChkItem = SelectValidItem(aItems, vsf2.RowData(iRow))

        '如果该标本没有指标，则删除（微生物的第一个仪器除外）
        If Not blnChkItem And vsf2.Rows > 2 And Not (mbln微生物项目 And iRow = 1) And mintEditMode <= 1 Then
            vsf2.RemoveItem iRow
            iRow = iRow - 1
        End If
        If iRow = vsf2.Rows - 1 Then Exit For
    Next iRow

    astrKey = Split(mstrKeys, ",")
    For intLoop = 0 To UBound(astrKey)
        blnCheck = False
        For iRow = 1 To vsf2.Rows - 1
            If InStr(vsf2.TextMatrix(iRow, 3), Chr(1) & astrKey(intLoop) & Chr(1)) > 0 Then
                blnCheck = True
                Exit For
            End If
        Next
        If blnCheck = False Then
            For lngloop = 0 To UBound(aItems, 2)
                If aItems(1, lngloop) = astrKey(intLoop) Then
                    vsf2.TextMatrix(1, 3) = vsf2.TextMatrix(1, 3) & "|" & aItems(0, lngloop) & Chr(1) & aItems(1, lngloop) & _
                         Chr(1) & aItems(2, lngloop) & Chr(1) & aItems(3, lngloop) & Chr(1) & aItems(4, lngloop) & Chr(1) & aItems(5, lngloop) & _
                         Chr(1) & aItems(6, lngloop)
                    Exit For
                End If
            Next

        End If
    Next

'    For iRow = 1 To vsf2.Rows - 1
'        If Not mbln微生物项目 And Trim(vsf2.TextMatrix(iRow, 3)) = "" Then
'            vsf2.RemoveItem iRow
'            iRow = iRow - 1
'        End If
'        If iRow = vsf2.Rows - 1 Then Exit For
'    Next iRow

    vsf2.Row = iCurrRow
End Sub

Private Function ReadData() As Variant()
    '--------------------------------------------------------------------------------------------------------
    '功能：获取项目检验指标到数组
    '--------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset, rsMicro As New ADODB.Recordset
    Dim strField As String, i As Long
    Dim strSubQry As String, mstrSql As String
    Dim blnOnlyMachine As Boolean   '只核收当前默认的仪器项目
    Dim strWhere As String
    Dim strItems As String

    ReadData = Array()
    On Error GoTo ErrHand
    If mstrKeys = "" And rsRelativeAdvice Is Nothing Then
        Exit Function
    End If

    blnOnlyMachine = zlDatabase.GetPara("只核收当前仪器项目", 100, 1208, 0)

    If blnOnlyMachine = True And mlngDefaultDevice > 0 Then
        strWhere = " And B.报告项目id = E.项目id and  E.仪器ID = [3] "
    Else
        strWhere = " And B.报告项目id = E.项目id(+) "
    End If

    '读取申请的检验项目(检验指标)

    If mstrKeys <> "" Then
        If Not mbln微生物项目 Then
            If GetApplicationFormShowType = True Then
                strWhere = " and m.组合项目 <> 1 "
            End If
'            mstrSQL = "select ID,相关ID,结果,标志,结果参考,诊疗项目ID,RowNum as 排列序号,仪器,选择 From " & _
                "(SELECT ID,相关id,结果,标志,结果参考,诊疗项目ID,排列序号,编码," & _
                "' '||仪器1||' '||仪器2||' '||仪器3||' '||仪器4||' '||仪器5||' '||仪器6||' '||仪器7||' '||仪器8||' '||仪器9||' ' AS 仪器,0 As 选择 " & _
                "FROM " & _
                "   ( " & _
                "    Select a.ID,a.相关id,a.结果,a.标志,a.结果参考,a.诊疗项目id,排列序号,编码, " & _
                "          Max(decode(mod(rownum,9),0,a.仪器id,'')) as 仪器1, " & _
                "          Max(decode(mod(rownum,9),1,a.仪器id,'')) as 仪器2, " & _
                "          Max(decode(mod(rownum,9),2,a.仪器id,'')) as 仪器3, " & _
                "          Max(decode(mod(rownum,9),3,a.仪器id,'')) as 仪器4, " & _
                "          Max(decode(mod(rownum,9),4,a.仪器id,'')) as 仪器5, " & _
                "          Max(decode(mod(rownum,9),5,a.仪器id,'')) as 仪器6, " & _
                "          Max(decode(mod(rownum,9),6,a.仪器id,'')) as 仪器7, " & _
                "          Max(decode(mod(rownum,9),7,a.仪器id,'')) as 仪器8, " & _
                "          Max(decode(mod(rownum,9),8,a.仪器id,'')) as 仪器9 "

'             mstrSQL = mstrSQL & "    From (" & _
                "           SELECT C.ID,A.相关id,Decode(D.结果类型,3,Nvl(D.默认值,'-'),2,D.默认值,'') As 结果,'' As 标志," & _
                "           Trim(REPLACE(REPLACE(' '||zlGetReference(C.ID,A.标本部位,DECODE(F.性别,'男',1,'女',2,0),F.出生日期," & mlngDefaultDevice & "),' .','0.'),'～.','～0.')) AS 结果参考,B.诊疗项目ID,B.排列序号," & _
                "           e.仪器ID,decode(D.排列序号,NULL,M.编码,D.排列序号) as 编码 " & _
                "           FROM 病人医嘱记录 A,检验报告项目 B,诊治所见项目 C,检验项目 D,检验仪器项目 E,病人信息 F,诊疗项目目录 M " & _
                "           WHERE A.相关id>0 " & _
                "               AND A.诊疗项目ID+0=B.诊疗项目ID AND B.细菌ID Is Null " & _
                "               AND B.报告项目ID=C.ID AND A.病人ID=F.病人ID And B.诊疗项目ID = M.ID " & _
                "               AND D.诊治项目ID=C.ID AND B.报告项目ID=E.项目ID(+) AND A.病人ID=[1] AND Instr(','||[2]||',',','||A.相关ID||',')>0 Order by C.ID ) a " & _
                "   Group by  a.ID,a.相关id,a.结果,a.标志,a.结果参考,a.诊疗项目id,排列序号,编码 ) " & _
                "Order by 编码,排列序号) "
            mstrSql = "select ID, 相关id, 结果, 标志, 结果参考, 诊疗项目id, rownum as  排列序号, 仪器id, 0 As 选择" & vbNewLine & _
            "from" & vbNewLine & _
            "(Select ID, 相关id, 结果, 标志, 结果参考, 诊疗项目id,  排列序号, 仪器id, 0 As 选择" & vbNewLine & _
            "From (Select C.ID, A.相关id, Nvl(D.默认值, '') As 结果, '' As 标志," & vbNewLine & _
            "              Trim(Replace(Replace(' ' || Zlgetreference(C.ID, A.标本部位, Decode(F.性别, '男', 1, '女', 2, 0), F.出生日期), ' .'," & vbNewLine & _
            "                                    '0.'), '～.', '～0.')) As 结果参考, B.诊疗项目id, B.排列序号, E.仪器id," & vbNewLine & _
            "              lpad(Decode(D.排列序号, Null, M.编码, D.排列序号),10,'0') As 编码 " & vbNewLine & _
            "       From 病人医嘱记录 A, 检验报告项目 B, 诊治所见项目 C, 检验项目 D, 检验仪器项目 E, 病人信息 F, 诊疗项目目录 M" & vbNewLine & _
            "       Where A.相关id > 0 And A.诊疗项目id + 0 = B.诊疗项目id And B.细菌id Is Null And B.报告项目id = C.ID And A.病人id = F.病人id And" & vbNewLine & _
            "             B.诊疗项目id = M.ID And D.诊治项目id = C.ID  And A.病人id = [1] And" & vbNewLine & _
            "             Instr(',' ||[2]|| ',', ',' || A.相关id || ',') > 0" & vbNewLine & _
            "             " & strWhere & vbNewLine & _
            "       Order By C.ID)" & vbNewLine & _
            "order by 编码,排列序号)"


        Else
'            mstrSQL = "SELECT ID,相关id,结果,标志,结果参考,诊疗项目ID,rownum as 排列序号, " & _
                "' '||仪器1||' '||仪器2||' '||仪器3||' '||仪器4||' '||仪器5||' '||仪器6||' '||仪器7||' '||仪器8||' '||仪器9||' ' AS 仪器,0 As 选择 " & _
                "FROM " & _
                "(SELECT D.ID,A.相关id,'' As 结果,'' As 标志,'' AS 结果参考,B.诊疗项目ID,B.排列序号," & _
                "          Max(decode(mod(rownum,9),0,e.仪器id,'')) as 仪器1, " & _
                "          Max(decode(mod(rownum,9),1,e.仪器id,'')) as 仪器2, " & _
                "          Max(decode(mod(rownum,9),2,e.仪器id,'')) as 仪器3, " & _
                "          Max(decode(mod(rownum,9),3,e.仪器id,'')) as 仪器4, " & _
                "          Max(decode(mod(rownum,9),4,e.仪器id,'')) as 仪器5, " & _
                "          Max(decode(mod(rownum,9),5,e.仪器id,'')) as 仪器6, " & _
                "          Max(decode(mod(rownum,9),6,e.仪器id,'')) as 仪器7, " & _
                "          Max(decode(mod(rownum,9),7,e.仪器id,'')) as 仪器8, " & _
                "          Max(decode(mod(rownum,9),8,e.仪器id,'')) as 仪器9 " & _
                " FROM 病人医嘱记录 A,检验报告项目 B,检验细菌 D,仪器细菌对照 E,病人信息 F " & _
                " WHERE A.相关id>0 " & _
                    "AND A.诊疗项目ID+0=B.诊疗项目ID " & _
                    "AND B.细菌ID=D.ID AND A.病人ID=F.病人ID " & _
                    "AND B.细菌ID=E.细菌ID(+) AND A.病人ID=[1] AND Instr(','||[2]||',',','||A.相关ID||',')>0" & _
                " GROUP BY D.ID,A.相关id,'','','',B.诊疗项目ID,B.排列序号 Order By B.诊疗项目ID,B.排列序号 Desc)"
            mstrSql = "Select ID, 相关id, 结果, 标志, 结果参考, 诊疗项目id, Rownum As 排列序号, 仪器id, 0 As 选择" & vbNewLine & _
                "From (Select D.ID, A.相关id, D.默认结果 As 结果, '' As 标志, '' As 结果参考, B.诊疗项目id, B.排列序号, 仪器id" & vbNewLine & _
                "       From 病人医嘱记录 A, 检验报告项目 B, 检验细菌 D, 仪器细菌对照 E, 病人信息 F" & vbNewLine & _
                "       Where A.相关id > 0 And A.诊疗项目id + 0 = B.诊疗项目id And B.细菌id = D.ID And A.病人id = F.病人id And B.细菌id = E.细菌id(+) And" & vbNewLine & _
                "             A.病人id = [1] And Instr(',' ||[2]|| ',', ',' || A.相关id || ',') > 0" & vbNewLine & _
                "       Order By B.诊疗项目id, B.排列序号 )"

        End If
        Set rs = zlDatabase.OpenSQLRecord(mstrSql, Me.Caption, mlng病人ID, mstrKeys, mlngDefaultDevice, strItems)
        If rs.BOF = False Then
            vsf2.Tag = rs.RecordCount

            ReadData = rs.GetRows
        End If
    Else
        If Not rsRelativeAdvice Is Nothing Then
            strSubQry = ""
            rsRelativeAdvice.MoveFirst
            Do While Not rsRelativeAdvice.EOF
                strSubQry = strSubQry & " Union All " & "Select " & rsRelativeAdvice("ID") & " As ID From Dual"
                If strItems = "" Then
                    strItems = Val(rsRelativeAdvice("ID") & "")
                Else
                    strItems = strItems & "," & Val(rsRelativeAdvice("ID") & "")
                End If
                rsRelativeAdvice.MoveNext
            Loop
            If Len(strSubQry) > 0 Then strSubQry = Mid(strSubQry, 12)
            rsRelativeAdvice.MoveFirst

            '读取申请的检验项目(检验指标)
            If Not mbln微生物项目 Then
'                mstrSQL = "select ID,相关ID,结果,标志,结果参考,诊疗项目ID,RowNum as 排列序号,仪器,选择 From " & _
                    "(SELECT ID,相关ID,结果,标志,结果参考,诊疗项目ID,排列序号,编码," & _
                    "' '||仪器1||' '||仪器2||' '||仪器3||' '||仪器4||' '||仪器5||' '||仪器6||' '||仪器7||' '||仪器8||' '||仪器9||' ' AS 仪器,0 As 选择 " & _
                    "FROM " & _
                    "  (" & _
                    "   Select a.id , a.相关id ,a.结果,a.标志 , a.结果参考,a.诊疗项目id,a.排列序号,编码," & _
                    "          Max(decode(mod(rownum,9),0,a.仪器id,'')) as 仪器1, " & _
                "          Max(decode(mod(rownum,9),1,a.仪器id,'')) as 仪器2, " & _
                "          Max(decode(mod(rownum,9),2,a.仪器id,'')) as 仪器3, " & _
                "          Max(decode(mod(rownum,9),3,a.仪器id,'')) as 仪器4, " & _
                "          Max(decode(mod(rownum,9),4,a.仪器id,'')) as 仪器5, " & _
                "          Max(decode(mod(rownum,9),5,a.仪器id,'')) as 仪器6, " & _
                "          Max(decode(mod(rownum,9),6,a.仪器id,'')) as 仪器7, " & _
                "          Max(decode(mod(rownum,9),7,a.仪器id,'')) as 仪器8, " & _
                "          Max(decode(mod(rownum,9),8,a.仪器id,'')) as 仪器9 " & _
                    "   From ( " & _
                    "       SELECT C.ID,0 As 相关ID,Decode(D.结果类型,3,Nvl(D.默认值,'-'),2,D.默认值,'') As 结果,'' As 标志," & _
                    "       Trim(REPLACE(REPLACE(' '||zlGetReference(C.ID,'" & txt附加.Text & "'," & Decode(cbo性别.Text, "男", 1, "女", 2, 0) & ",NULL," & mlngDefaultDevice & "),' .','0.'),'～.','～0.')) AS 结果参考,B.诊疗项目ID,B.排列序号," & _
                    "       e.仪器id,decode(D.排列序号,NULL,M.编码,D.排列序号) as 编码  " & _
                    "       FROM 检验报告项目 B,诊治所见项目 C,检验项目 D,检验仪器项目 E,(" & strSubQry & ") S , 诊疗项目目录 M " & _
                    "       WHERE B.诊疗项目ID=S.ID AND B.细菌ID Is Null " & _
                    "           AND B.报告项目ID=C.ID And B.诊疗项目Id = M.ID " & _
                    "           AND D.诊治项目ID=C.ID AND B.报告项目ID=E.项目ID(+) order by c.id  ) a " & _
                    "   Group by a.id , a.相关id ,a.结果,a.标志 , a.结果参考,a.诊疗项目id,a.排列序号,a.编码) " & _
                    "Order by 编码,排列序号) "
                mstrSql = "Select ID, 相关id, 结果, 标志, 结果参考, 诊疗项目id, Rownum As 排列序号, 仪器id, 0 As 选择" & vbNewLine & _
                    "From (Select ID, 相关id, 结果, 标志, 结果参考, 诊疗项目id, 排列序号, 仪器id, 0 As 选择" & vbNewLine & _
                    "       From (Select C.ID, 0 As 相关id, Nvl(D.默认值, '') As 结果, '' As 标志," & vbNewLine & _
                    "                     Trim(Replace(Replace(' ' || Zlgetreference(C.ID, '抗凝血', 0, Null), ' .', '0.'), '～.', '～0.')) As 结果参考," & vbNewLine & _
                    "                     B.诊疗项目id, B.排列序号, E.仪器id, lpad(Decode(D.排列序号, Null, M.编码, D.排列序号),10,'0') As 编码 " & vbNewLine & _
                    "              From 检验报告项目 B, 诊治所见项目 C, 检验项目 D, 检验仪器项目 E , 诊疗项目目录 M" & vbNewLine & _
                    "              Where B.诊疗项目id in (Select * From Table(Cast(f_Num2list([4]) As zlTools.t_Numlist))) And B.细菌id Is Null And B.报告项目id = C.ID And B.诊疗项目id = M.ID And D.诊治项目id = C.ID " & vbNewLine & _
                    "              " & strWhere & vbNewLine & _
                    "              Order By C.ID)" & vbNewLine & _
                    "       Order By 编码, 排列序号)"

            Else
'                mstrSQL = "SELECT ID,相关ID,结果,标志,结果参考,诊疗项目ID,rownum as 排列序号," & _
                    "' '||仪器1||' '||仪器2||' '||仪器3||' '||仪器4||' '||仪器5||' '||仪器6||' '||仪器7||' '||仪器8||' '||仪器9||' ' AS 仪器,0 As 选择 " & _
                    "FROM " & _
                    "(SELECT D.ID,0 As 相关ID,'' As 结果,'' As 标志,'' AS 结果参考,B.诊疗项目ID,B.排列序号," & _
                    "          Max(decode(mod(rownum,9),0,e.仪器id,'')) as 仪器1, " & _
                "          Max(decode(mod(rownum,9),1,e.仪器id,'')) as 仪器2, " & _
                "          Max(decode(mod(rownum,9),2,e.仪器id,'')) as 仪器3, " & _
                "          Max(decode(mod(rownum,9),3,e.仪器id,'')) as 仪器4, " & _
                "          Max(decode(mod(rownum,9),4,e.仪器id,'')) as 仪器5, " & _
                "          Max(decode(mod(rownum,9),5,e.仪器id,'')) as 仪器6, " & _
                "          Max(decode(mod(rownum,9),6,e.仪器id,'')) as 仪器7, " & _
                "          Max(decode(mod(rownum,9),7,e.仪器id,'')) as 仪器8, " & _
                "          Max(decode(mod(rownum,9),8,e.仪器id,'')) as 仪器9 " & _
                    " FROM 检验报告项目 B,检验细菌 D,仪器细菌对照 E,(" & strSubQry & ") S " & _
                    " WHERE B.诊疗项目ID=S.ID " & _
                        "AND B.细菌ID=D.ID " & _
                        "AND B.细菌ID=E.细菌ID(+)" & _
                    " GROUP BY D.ID,0,'','','',B.诊疗项目ID,B.排列序号 Order By B.诊疗项目ID,B.排列序号 Desc)"
                mstrSql = "Select ID, 相关id, 结果, 标志, 结果参考, 诊疗项目id, Rownum As 排列序号, 仪器id, 0 As 选择" & vbNewLine & _
                "From (Select D.ID, 0 As 相关id, D.默认结果 As 结果, '' As 标志, '' As 结果参考, B.诊疗项目id, B.排列序号, E.仪器id" & vbNewLine & _
                "       From 检验报告项目 B, 检验细菌 D, 仪器细菌对照 E " & vbNewLine & _
                "       Where B.诊疗项目id in (Select * From Table(Cast(f_Num2list([4]) As zlTools.t_Numlist))) And B.细菌id = D.ID And B.细菌id = E.细菌id(+)" & vbNewLine & _
                "       Order By B.诊疗项目id, B.排列序号 )"


            End If
            Set rs = zlDatabase.OpenSQLRecord(mstrSql, Me.Caption, mlng病人ID, mstrKeys, mlngDefaultDevice, strItems)
'            Call OpenRecord(rs, mstrSQL, Me.Caption)
            If rs.BOF = False Then
                vsf2.Tag = rs.RecordCount

                ReadData = rs.GetRows
            End If
        End If
    End If

    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function SelectValidItem(aItems() As Variant, ByVal lngDeviceID As Long) As Boolean
'将当前仪器可作的指标全部加到列表中
    Dim mlngLoop As Long, lngItemRow  As Long

    SelectValidItem = False
    vsf2.TextMatrix(vsf2.Row, 3) = ""
'    For mlngLoop = UBound(aItems, 2) To 0 Step -1              '不知道以前为什么要返过来循环
    For mlngLoop = 0 To UBound(aItems, 2)
        If Val(aItems(ItemCol.ID, mlngLoop)) > 0 And Val(aItems(ItemCol.选择, mlngLoop)) = 0 Then
            If (InStr(IIf(Trim(aItems(ItemCol.仪器, mlngLoop)) = "", "-1", aItems(ItemCol.仪器, mlngLoop)), lngDeviceID) > 0) Or _
                (lngDeviceID = -1 And mlngDefaultDevice = -1) Or Trim(Nvl(aItems(ItemCol.仪器, mlngLoop))) = "" Then
                '填写医嘱，主要用于已有标本的情况
                If vsf2.TextMatrix(vsf2.Row, 4) = "" Or Val(vsf2.TextMatrix(vsf2.Row, 4)) = 0 Then
                    vsf2.TextMatrix(vsf2.Row, 4) = aItems(ItemCol.相关ID, mlngLoop)
                End If
                aItems(ItemCol.选择, mlngLoop) = 1
                If InStr("|" & vsf2.TextMatrix(vsf2.Row, 3), "|" & aItems(0, mlngLoop) & Chr(1)) = 0 Then
                    SelectValidItem = True

                    vsf2.TextMatrix(vsf2.Row, 3) = vsf2.TextMatrix(vsf2.Row, 3) & "|" & aItems(0, mlngLoop) & Chr(1) & aItems(1, mlngLoop) & _
                         Chr(1) & aItems(2, mlngLoop) & Chr(1) & aItems(3, mlngLoop) & Chr(1) & aItems(4, mlngLoop) & Chr(1) & aItems(5, mlngLoop) & _
                         Chr(1) & aItems(6, mlngLoop)
                    For lngItemRow = 0 To UBound(aItems, 2)
                        If Val(aItems(ItemCol.ID, lngItemRow)) = Val(aItems(ItemCol.ID, mlngLoop)) Then
                            aItems(ItemCol.选择, lngItemRow) = 1
                        End If
                    Next
                End If
            End If
        End If
    Next mlngLoop
    If Len(vsf2.TextMatrix(vsf2.Row, 3)) > 0 Then vsf2.TextMatrix(vsf2.Row, 3) = Mid(vsf2.TextMatrix(vsf2.Row, 3), 2)
End Function

Private Function PreFixNO(Optional curDate As Date = #1/1/1900#) As String
'功能：返回大写的单据号年前缀
    If curDate = #1/1/1900# Then
        PreFixNO = CStr(CInt(Format(zlDatabase.Currentdate, "YYYY")) - 1990)
    Else
        PreFixNO = CStr(CInt(Format(curDate, "YYYY")) - 1990)
    End If
    PreFixNO = IIf(CInt(PreFixNO) < 10, PreFixNO, Chr(55 + CInt(PreFixNO)))
End Function

Private Function GetFullNO(ByVal strNO As String, ByVal intNum As Integer) As String
'功能：由用户输入的部份单号，返回全部的单号。
'参数：intNum=项目序号,为0时固定按年产生
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, intType As Integer
    Dim curDate As Date

    If Len(strNO) >= 8 Then
        GetFullNO = Right(strNO, 8)
        Exit Function
    ElseIf Len(strNO) = 7 Then
        GetFullNO = PreFixNO & strNO
        Exit Function
    ElseIf intNum = 0 Then
        GetFullNO = PreFixNO & Format(strNO, "0000000")
        Exit Function
    End If
    GetFullNO = strNO

    strSQL = "Select 编号规则,Sysdate as 日期 From 号码控制表 Where 项目序号=" & intNum
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If Not rsTmp.EOF Then
        intType = Nvl(rsTmp!编号规则, 0)
        curDate = rsTmp!日期
    End If

    If intType = 1 Then
        '按日编号
        strSQL = Format(CDate("1992-" & Format(rsTmp!日期, "MM-dd")) - CDate("1992-01-01") + 1, "000")
        GetFullNO = PreFixNO & strSQL & Format(Right(strNO, 4), "0000")
    Else
        '按年编号
        GetFullNO = PreFixNO & Format(strNO, "0000000")
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckHave(ByVal lngKey As Long) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '功能:检验是否已经使用过和是否有操作权限
    '参数:
    '返回:
    '--------------------------------------------------------------------------------------------------------
    Dim mlngLoop As Long

    '判断是否有权限操作
    If InStr(mstrMachines, ";" & lngKey & ";") = 0 Then
        CheckHave = True
        Exit Function
    End If

    '判断是否已使用
    For mlngLoop = 1 To vsf2.Rows - 1
        If vsf2.RowData(mlngLoop) = lngKey Then
            CheckHave = True
            Exit Function
        End If
    Next

End Function

Private Function ValidData() As Boolean
    '--------------------------------------------------------------------------------------------------------
    '功能：
    '--------------------------------------------------------------------------------------------------------
    Dim varTmp As Variant
    Dim strTmp As String
    Dim strError As String, mstrSql As String
    Dim lngloop As Long
    Dim lngCount As Long
    Dim rs As New ADODB.Recordset, rsTmp As New ADODB.Recordset
    Dim i As Integer, iType As Integer
    Dim strStartDate As String
    Dim strEndDate As String
    Dim str医嘱ID As String

    On Error GoTo errH

    Call txt姓名_LostFocus

    ValidData = False

    If Me.txt姓名.Text <> Me.txt姓名.Tag Then
        Me.txt姓名.Enabled = False
'        Call txt姓名_Validate(False)
        Call Txt姓名Exec
        If Me.txt姓名.Text <> "" Then
            SetFocusNextIndex txt姓名.TabIndex
        Else
            Me.txt姓名.Enabled = True
            Me.txt姓名.SetFocus
        End If
        gintSelectFocus = 2
        Me.txt姓名.Enabled = True
    End If



    If Len(Trim(Me.txt姓名)) = 0 Then
'        mintFocusItem = FocusItem.姓名
        If Me.txt姓名.Enabled = True Then
            Me.txt姓名.SetFocus
        End If
        Exit Function
    End If

    If Len(Trim(cbo开单科室.Text)) = 0 Then
        If Me.cbo开单科室.Enabled = True Then
            cbo开单科室.SetFocus
        End If
        Exit Function
    End If

    '检查采样时间不能大于核收时间
    If mintEditMode <> 3 Then
        If Trim(Me.cbo(1).Text) <> "" Then
            If CDate(Me.DTP(0).Value) > CDate(Me.DTP(1).Value) Then
                MsgBox "采样时间大于核收时间，请检查核收时间！", vbInformation, Me.Caption
                If Me.DTP(1).Enabled = True And Me.DTP(1).Visible = True Then
                    Me.DTP(1).SetFocus
                End If
                Exit Function
            End If
        End If
    End If

    If mstrKeys = "" And rsRelativeAdvice Is Nothing And mblnCheckIn = False Then
'        mintFocusItem = FocusItem.医嘱内容
        On Error Resume Next
        Me.txt医嘱内容.SetFocus
        Exit Function
    End If

    '1.检查每一个标本指定的检验仪器是否正确
    For i = 1 To vsf2.Rows - 1
        If Trim(vsf2.TextMatrix(i, 2)) = "" Then
            MsgBox "第" & i & "个标本没有标本号！", vbInformation, gstrSysName: gintSelectFocus = 2 'DoEvents
            vsf2.Row = i
            vsf2.Col = 2
            vsf2.SetFocus
            vsf2.ShowCell vsf2.Row, vsf2.Col
            Exit Function
        End If
    Next
    If vsf2.TextMatrix(vsf2.Row, vsf2.Col) <> vsf2.EditText And vsf2.EditText <> "" Then
        vsf2.TextMatrix(vsf2.Row, vsf2.Col) = vsf2.EditText
    End If
    ReDim mlngNoneHomeKey(vsf2.Rows - 1)
    ReDim mlngSourceKey(vsf2.Rows - 1)
    For i = 1 To vsf2.Rows - 1
        iType = IIf(vsf2.TextMatrix(i, 5) = "-1", 1, 0)
        '检查是否有效
        If Val(vsf2.RowData(i)) > 0 Then
            mstrSql = "SELECT ID,标本序号,Nvl(是否质控品,0) as 是否质控品,姓名 FROM 检验标本记录 WHERE   仪器id= [1] " & _
                " AND 核收时间 Between [2] AND [3] AND 标本序号=[4]" & _
                IIf(mblnEmerge, IIf(iType = 1, " And 标本类别=1", " And Nvl(标本类别,0)<>1"), "")
        Else
            mstrSql = "SELECT ID,标本序号,Nvl(是否质控品,0) as 是否质控品,姓名 FROM 检验标本记录 WHERE    仪器id Is Null " & _
                " AND 核收时间 Between [2] AND [3] AND 标本序号=[4]" & _
                IIf(mblnEmerge, IIf(iType = 1, " And 标本类别=1", " And Nvl(标本类别,0)<>1"), "")
        End If

        strStartDate = GetDateTime(mMakeNoRule, 1, DTP(1).Value)
        strEndDate = GetDateTime(mMakeNoRule, 2, DTP(1).Value)

        Set rs = zlDatabase.OpenSQLRecord(mstrSql, Me.Caption, Val(vsf2.RowData(i)), _
            CDate(Format(strStartDate, "yyyy-MM-dd 00:00:00")), _
            CDate(Format(strEndDate, "yyyy-MM-dd 23:59:59")), TransSampleNO(Trim(vsf2.TextMatrix(i, 2))))

        If rs.BOF = False Then
            If rs("是否质控品") = 1 Then
                '是质控品或已核收的标本不能被覆盖
                MsgBox "你设置的标本号是质控品，请重新设定标本号！", vbInformation, Me.Caption
                vsf2.Row = i
                vsf2.Col = 2
                vsf2.SetFocus
                vsf2.ShowCell vsf2.Row, vsf2.Col
                gintSelectFocus = 2
                Exit Function
            End If

            '核收、登记、补填是否复盖无主标本
            If mintEditMode = 3 Then
                mlngNoneHomeKey(i) = mlngSampleID
                rs.filter = "ID<>" & mlngSampleID
                If rs.RecordCount > 0 Then
                    If Trim(Nvl(rs("姓名"))) <> "" Then
                        '是质控品或已核收的标本不能被覆盖
                        MsgBox "你设置的标本号已被核收，请重新设定标本号！", vbInformation, Me.Caption
                        vsf2.Row = i
                        vsf2.Col = 2
                        vsf2.SetFocus
                        vsf2.ShowCell vsf2.Row, vsf2.Col
                        gintSelectFocus = 2
                        Exit Function
                    End If
                    mlngSourceKey(i) = rs("ID")
                End If
            Else
                If Trim(Nvl(rs("姓名"))) <> "" Then
                    MsgBox "你设置的标本号已被核收，请重新设定标本号！", vbInformation, Me.Caption
                    vsf2.Row = i
                    vsf2.Col = 2
                    vsf2.SetFocus
                    vsf2.ShowCell vsf2.Row, vsf2.Col
                    gintSelectFocus = 2
                    Exit Function
                End If
                If MsgBox("你设置的标本号已经存在，是需要否复盖?", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    ValidData = True
                    mlngNoneHomeKey(i) = rs("ID").Value
                    gintSelectFocus = 2
                Else
                    vsf2.Row = i
                    vsf2.Col = 2
                    vsf2.SetFocus
                    vsf2.ShowCell vsf2.Row, vsf2.Col
                    gintSelectFocus = 2
                    Exit Function
                End If
            End If
        Else
            '补填写入一个新的标本号
            mlngNoneHomeKey(i) = mlngSampleID
        End If
    Next

    '--------------------------------------------------------------------------------------------------------------------------------
    '当执行完成有自动审核的费用时，对病人费用进行记帐报警。
    gstrSql = " select /*+ rule */ id from 病人医嘱记录 where id in (Select * From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))) " & _
              " Union All " & _
              " select /*+ rule */ id from 病人医嘱记录 where 相关id in (Select * From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))) "
    Set rs = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mstrKeys)
    Do While Not rs.EOF
        str医嘱ID = str医嘱ID & "," & rs("id")
        rs.MoveNext
    Loop
    str医嘱ID = Mid(str医嘱ID, 2)
    If Chk划价费用(Me, str医嘱ID, 0) = False And Trim(str医嘱ID) <> "" Then
        Exit Function
    End If
    '----------------------------------------------------------------------------------------------------------------------------------
    ValidData = True

    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckMuliQuest(ByVal lng病人ID As Long, ByVal lng仪器id As Long, ByVal strNO As String, ByRef lngKey As Long, ByVal iType As Integer, ByRef blnOther As Boolean) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：
    '参数：     iType                 标本类别：0=普通、1=急诊
    '------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    Dim strStartDate  As String
    Dim strEndDate As String

    On Error GoTo ErrHand

    If lng仪器id > 0 Then
        strSQL = "SELECT A.ID,B.病人ID,C.姓名 FROM 检验标本记录 A,病人医嘱记录 B,病人信息 C WHERE A.医嘱id=B.id AND B.病人ID=C.病人ID" & _
        " AND A.仪器id=[2] AND A.核收时间 Between [3] And [4] AND A.标本序号= [5] " & _
        IIf(mblnEmerge, IIf(iType = 1, " And A.标本类别=1", " And Nvl(A.标本类别,0)<>1"), "")
    Else
        strSQL = "SELECT A.ID,B.病人ID,C.姓名 FROM 检验标本记录 A,病人医嘱记录 B,病人信息 C WHERE A.医嘱id=B.id AND B.病人ID=C.病人ID" & _
        " AND A.仪器id IS NULL AND A.核收时间 Between [3] And [4] AND A.标本序号= [5] " & _
        IIf(mblnEmerge, IIf(iType = 1, " And A.标本类别=1", " And Nvl(A.标本类别,0)<>1"), "")
    End If

    strStartDate = GetDateTime(mMakeNoRule, 1, DTP(1).Value)
    strEndDate = GetDateTime(mMakeNoRule, 2, DTP(1).Value)

    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng仪器id, _
        CDate(Format(strStartDate, "yyyy-MM-dd 00:00:00")), _
        CDate(Format(strEndDate, "yyyy-MM-dd 23:59:59")), strNO)

    If rs.BOF = False Then
        If mintEditMode <= 1 Then
            Call MsgBox("你设置的标本号已经存在，请重新设定标本号!", vbInformation, gstrSysName): gintSelectFocus = 2 'DoEvents
'            mintFocusItem = FocusItem.标本号

            vsf2.Col = 2
            vsf2.SetFocus
            vsf2.ShowCell vsf2.Row, vsf2.Col
            Exit Function
        End If
        If Not IsNull(rs("病人ID")) Then
            lngKey = zlCommFun.Nvl(rs("ID"), 0)
            blnOther = True
        End If
    End If

    CheckMuliQuest = True

    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
End Function

Private Function SaveData(Optional ByVal intEditState As Integer) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '功能：
    '--------------------------------------------------------------------------------------------------------
    Dim blnTran As Boolean
    Dim varTmp As Variant
    Dim lngloop As Long
    Dim strSQL() As String
    Dim blnMuliQuest As Boolean
    Dim lngMuliQuestKey As Long
    Dim mlngKey As Long '医嘱ID
    Dim lngKey As Long '标本ID
    Dim lngResultID As Long '结果ID，用于微生物
    Dim lngResultLoop As Long
    Dim i As Integer, varAdviceIDs As Variant '指标对应的若干医嘱ID
    Dim strItemRecords As String
    Dim AdviceIDs() As Long, SampleIDs() As Long
    Dim strReportCode As String
    Dim strReportParaNo As String
    Dim bytReportParaMode As Byte
    Dim blnAutoPrint As Boolean
    Dim strTmpNO As String '标本号
    Dim blnOther As Boolean '是否其他患者标本
    Dim strTmp As String, rsTmp As ADODB.Recordset
    Dim mlngLoop As Long, blnAuditing As Boolean
    Dim blnNewAdvice As Boolean         '是否新建医嘱
    Dim blnEmergency As Integer         '是否使用急诊 0=不使用急诊 1=使用急诊
    Dim strStartDate  As String
    Dim strEndDate As String
    Dim strItems As String              '诊疗项目ID,多个时使用","分隔
    Dim blnNewpatinet As Boolean        '是否是生成的新病人
    ReDim strSQL(1 To 1)
    ReDim AdviceIDs(0)
    ReDim SampleIDs(0)


    blnAuditing = zlDatabase.GetPara("保存后直接审核", 100, 1208, True, 0)
    blnEmergency = Val(zlDatabase.GetPara("急诊标本", 100, 1208, 0))

    On Error GoTo ErrHand


    '处理并发问题，先生成病人信息
    blnNewpatinet = CreatePatient
    If mlng病人ID = 0 Then
        MsgBox "创建病人失败，请重试", vbInformation, Me.Caption
        Exit Function
    End If

    '登记，处理医嘱
    If mstrKeys = "" And mblnSaveAdvice Then
        If Not ValidAdvice Then
            SaveData = False
            Exit Function
        End If
        mlngKey = SaveAdviceData(blnNewpatinet)
        If mlngKey = -1 Then Exit Function
        blnNewAdvice = True
    Else
        blnNewAdvice = False
    End If


    blnAutoPrint = zlDatabase.GetPara("审核打印", 100, 1208, True, 0)

    strStartDate = GetDateTime(mMakeNoRule, 1, DTP(1).Value)
    strEndDate = GetDateTime(mMakeNoRule, 2, DTP(1).Value)

    For mlngLoop = 1 To vsf2.Rows - 1


        '======================================================================================================================
        '生成标本ID
        If mlngNoneHomeKey(mlngLoop) = 0 Then
            lngKey = zlDatabase.GetNextId("检验标本记录")
        Else
            lngKey = mlngNoneHomeKey(mlngLoop)
        End If
        '======================================================================================================================

        '======================================================================================================================
        '如果自动生成的新病人，给用户一个提示。
        gstrSql = "select distinct a.ID,a.病人ID,a.姓名,c.名称,a.标本序号 from 检验标本记录 a,检验项目分布 b,检验仪器 c where a.id = b.标本id and a.仪器id = c.id and  a.id = [1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngKey)


        If rsTmp.EOF = False Then
            If Val(Nvl(rsTmp("病人ID"))) <> mlng病人ID And Val(Nvl(rsTmp("病人ID"))) <> 0 Then
                '循环读出多个标本信息用于提示
                Do While Not rsTmp.EOF
                    strTmp = strTmp & "仪器：" & rsTmp("名称") & "    标本号:" & rsTmp("标本序号") & vbCrLf
                    rsTmp.MoveNext
                Loop
                rsTmp.MoveFirst
                If MsgBox("你重新登记了新病人，以前的病人<" & Nvl(rsTmp("姓名")) & ">的核收项目将自动回滚!" & vbCrLf & strTmp & _
                    "是否继续?", vbYesNo + vbDefaultButton2) = vbNo Then
                    Exit Function
                End If
            End If
        End If
        '=======================================================================================================================

        '=======================================================================================================================
        '核收前先回滚当前医嘱的所有标本（核收时显示包括了已核收的标本，这里这样处理是为了方便用户核收错时再进行核收的工作
        If intEditState <> 4 Then   '补填时不会滚
            gstrSql = "Select Distinct 医嘱ID From (Select 医嘱ID From 检验项目分布 Where 标本id = [1] " & _
                    "Union All Select 医嘱ID From 检验标本记录 Where ID = [1])"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngKey)
    
            Do While Not rsTmp.EOF
                If Not IsNull(rsTmp(0)) Then
                    strSQL(ReDimArray(strSQL)) = "ZL_检验标本记录_转为无主(" & rsTmp(0) & ")"
                    'zlDatabase.ExecuteProcedure "ZL_检验标本记录_转为无主(" & rsTmp(0) & ")", gstrSysName
                End If
                rsTmp.MoveNext
            Loop
        End If
        '========================================================================================================================


        '===========================================================================================================================================================
        '更新检验标本信息和检验普通结果信息
        If Val(vsf2.TextMatrix(mlngLoop, 4)) <> 0 Then
            mlngKey = Val(vsf2.TextMatrix(mlngLoop, 4))
            gstrSql = "Select 标本部位 From 病人医嘱记录 Where Id = [1] "
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngKey)
'            If rsTmp.EOF = False Then txt附加.Text = rsTmp("标本部位") & ""
        Else
            If mlngKey <= 0 Then mlngKey = Val(vsf2.TextMatrix(mlngLoop, 4))  '核收的默认医嘱ID
        End If

        ReDim Preserve AdviceIDs(UBound(AdviceIDs) + 1)
        ReDim Preserve SampleIDs(UBound(SampleIDs) + 1)
        AdviceIDs(UBound(AdviceIDs)) = mlngKey
        SampleIDs(UBound(SampleIDs)) = lngKey
        If vsf2.Row = mlngLoop And vsf2.Col = 2 And Me.vsf2.EditText <> "" Then
            vsf2.TextMatrix(mlngLoop, 2) = Me.vsf2.EditText
        End If
        strTmpNO = TransSampleNO(vsf2.TextMatrix(mlngLoop, 2))
        mstrNONumber = strTmpNO
        
        strSQL(ReDimArray(strSQL)) = "ZL_检验标本记录_标本核收(" & lngKey & "," & _
                                                                mlngKey & ",'" & IIf(Trim(mstrKeys) = "", mlngKey, mstrKeys & "," & mlngKey) & "'," & _
                                                                mlngSourceKey(mlngLoop) & ",'" & _
                                                                strTmpNO & "'," & _
                                                                IIf(cbo(1).Text <> "", "TO_DATE('" & Format(DTP(0).Value, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),'", "Null,'") & _
                                                                IIf(InStr(cbo(1).Text, "-") > 0, zlCommFun.GetNeedName(cbo(1).Text), cbo(1).Text) & "'," & _
                                                                IIf(Val(vsf2.RowData(mlngLoop)) = -1, 0, Val(vsf2.RowData(mlngLoop))) & "," & _
                                                                "TO_DATE('" & Format(DTP(1).Value, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),'" & _
                                                                IIf(InStr(cbo(0).Text, "-") > 0, zlCommFun.GetNeedName(cbo(0).Text), cbo(0).Text) & "','" & _
                                                                UserInfo.姓名 & "'," & _
                                                                "TO_DATE('" & Format(DTP(1).Value, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')," & IIf(mbln微生物项目, 1, "Null") & "," & _
                                                                IIf(mblnEmerge = True, IIf(vsf2.TextMatrix(mlngLoop, 5) = "-1" And mblnEmerge = True, 1, 0), 0) & ",NULL,'" & _
                                                                txt姓名.Text & "','" & NeedName(cbo性别.Text) & "','" & _
                                                                txt年龄 & Me.cboAge.Text & Me.txt年龄1 & "','" & mstrNO & "','" & _
                                                                txt附加.Text & "'," & Me.cbo开单科室.ItemData(Me.cbo开单科室.ListIndex) & ",'" & Me.cbo医生.Text & "'," & _
                                                                IIf(Trim(txtID.Text) = "", "NULL", IIf(IsNumeric(txtID.Text), txtID.Text, "NULL")) & "," & _
                                                                IIf(Trim(txtBed.Text) = "", "NULL,", "'" & Me.txtBed & "',") & _
                                                                IIf(Trim(txtPatientDept.Text) = "", "NULL,'", "'" & txtPatientDept.Text & "','") & _
                                                                Me.txt医嘱内容 & "'," & CInt(IIf(blnNewAdvice, 1, 0)) & _
                                                                "," & mlng病人ID & "," & ItemDeptID & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
        
        '注意：如果是微生物检验项目，则核收时不填写检验普通结果记录
        If vsf2.TextMatrix(mlngLoop, 3) = "" And mbln微生物项目 = False And mintEditMode > 1 Then
            vsf2.TextMatrix(mlngLoop, 3) = GetSampleData(mlngSampleID)
        End If
        varTmp = Split(vsf2.TextMatrix(mlngLoop, 3), "|")
        strItemRecords = ""
        strItems = ""
        For lngloop = 0 To UBound(varTmp)
            If mstrKeys <> "" Then
                '核收现在申请
                varAdviceIDs = Split(Split(varTmp(lngloop), Chr(1))(ItemCol.相关ID), ";")
                For i = 0 To UBound(varAdviceIDs)
                    mlngKey = Val(varAdviceIDs(i)) '指标对应的医嘱ID
                    If mlngKey > 0 Then
                        strItemRecords = strItemRecords & "|" & mlngKey & "^" & Val(Split(varTmp(lngloop), Chr(1))(ItemCol.ID)) & "^" & _
                            Split(varTmp(lngloop), Chr(1))(ItemCol.结果) & "^" & _
                            IIf(Len(Trim(Split(varTmp(lngloop), Chr(1))(ItemCol.标志))) = 0, 0, Decode(Right(Split(varTmp(lngloop), Chr(1))(ItemCol.标志), 2), "偏高", 3, "偏低", 2, "阳性", 4, 1)) & "^" & Split(varTmp(lngloop), Chr(1))(ItemCol.结果参考) & _
                            "^" & Split(varTmp(lngloop), Chr(1))(ItemCol.诊疗项目ID) & "^" & Split(varTmp(lngloop), Chr(1))(ItemCol.排列序号)
                        '记录诊疗项目
                        If InStr(strItems & ",", "," & Split(varTmp(lngloop), Chr(1))(ItemCol.诊疗项目ID) & ",") <= 0 Then
                            strItems = strItems & "," & Split(varTmp(lngloop), Chr(1))(ItemCol.诊疗项目ID)
                        End If
                    End If
                Next i
            Else
                If Val(Split(varTmp(lngloop), Chr(1))(ItemCol.相关ID)) > 0 Then
                    mlngKey = Val(Split(varTmp(lngloop), Chr(1))(ItemCol.相关ID))
                End If
                strItemRecords = strItemRecords & "|" & mlngKey & "^" & Val(Split(varTmp(lngloop), Chr(1))(ItemCol.ID)) & "^" & _
                    Split(varTmp(lngloop), Chr(1))(ItemCol.结果) & "^" & _
                    IIf(Len(Trim(Split(varTmp(lngloop), Chr(1))(ItemCol.标志))) = 0, 0, Decode(Right(Split(varTmp(lngloop), Chr(1))(ItemCol.标志), 2), "偏高", 3, "偏低", 2, "阳性", 4, 1)) & "^" & Split(varTmp(lngloop), Chr(1))(ItemCol.结果参考) & _
                    "^" & Split(varTmp(lngloop), Chr(1))(ItemCol.诊疗项目ID) & "^" & Split(varTmp(lngloop), Chr(1))(ItemCol.排列序号)
                '记录诊疗项目
                If InStr(strItems & ",", "," & Split(varTmp(lngloop), Chr(1))(ItemCol.诊疗项目ID) & ",") <= 0 Then
                    strItems = strItems & "," & Split(varTmp(lngloop), Chr(1))(ItemCol.诊疗项目ID)
                End If
            End If
        Next lngloop
        If Len(strItemRecords) > 0 Then
            strItemRecords = Mid(strItemRecords, 2)
            strSQL(ReDimArray(strSQL)) = "Zl_检验普通结果_Write(" & lngKey & "," & _
                IIf(Val(vsf2.RowData(mlngLoop)) = -1, 0, Val(vsf2.RowData(mlngLoop))) & ",'" & _
                strItemRecords & "',0," & IIf(mbln微生物项目, 1, 0) & ")"

            If mbln微生物项目 = False Then
                '删除当前核收项目里没有的空项目
                strSQL(ReDimArray(strSQL)) = "Zl_检验普通结果_DeleteItem(" & lngKey & ",'" & Mid(strItems, 2) & "'," & IIf(mbln微生物项目, 1, 0) & ")"
            End If
        Else
            '修改参考值和标志
            strSQL(ReDimArray(strSQL)) = "Zl_检验普通结果_Write(" & lngKey & "," & _
                IIf(Val(vsf2.RowData(mlngLoop)) = -1, 0, Val(vsf2.RowData(mlngLoop))) & ",'',0," & IIf(mbln微生物项目, 1, 0) & ",'" & IIf(Trim(mstrKeys) = "", mlngKey, mstrKeys & "," & mlngKey) & "')"
        End If
        strSQL(ReDimArray(strSQL)) = "Zl_重新计算结果_Cale(" & lngKey & ")"
        '===========================================================================================================================================================================

    Next

    If mlngSampleID = 0 Then mlngSampleID = lngKey
    gcnOracle.BeginTrans
    blnTran = True
    '集中执行SQL
    For mlngLoop = 1 To UBound(strSQL)
        If strSQL(mlngLoop) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(mlngLoop), Me.Caption)
    Next
    gcnOracle.CommitTrans
    If mstrKeys <> "" Then
        ModifyApplyToLIS mstrKeys, 1
    End If
    '检验签名
    If Signature(lngKey, gstrDBUser, "核收") = False Then
        Exit Function
    End If



    '设置病人信息为不能写入
    SetPatientInfoWrite True

    '审核完成自动审核
    If blnAuditing And mintEditMode > 1 And Len(mstrAuditer) > 0 Then
        '核收登记的不能立即审核
        For mlngLoop = 1 To vsf2.Rows - 1
            '检验审核规则判断
            If VerifyAuditingRule(lngKey) = 1 Then
                If MsgBox("验单有结果超出警示值!是否续继?", _
                    vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then
                    Exit Function
                End If
            End If
            If mSendReport = 1 And mstr初审人 = "" Then
                '初审
                gstrSql = "Zl_检验标本记录_初审报告(" & lngKey & ",1,'" & UserInfo.姓名 & "')"
                zlDatabase.ExecuteProcedure gstrSql, Me.Caption
            Else
                Call zlDatabase.ExecuteProcedure("ZL_检验标本记录_报告审核(" & SampleIDs(mlngLoop) & ",'" & mstrAuditer & "','" & UserInfo.编号 & _
                                                "','" & UserInfo.姓名 & "')", Me.Caption)
                If blnAutoPrint Then
                    If GetReportCode(AdviceIDs(mlngLoop), 0, strReportCode, strReportParaNo, bytReportParaMode) Then
                        Call ReportOpen(gcnOracle, glngSys, strReportCode, Me, "NO=" & strReportParaNo, "性质=" & bytReportParaMode, "医嘱ID=" & AdviceIDs(mlngLoop), "标本ID=" & lngKey, "病人ID=" & mlng病人ID, 2)
                    End If
                End If
            End If
        Next
    End If
    SaveData = True

    mblnSaveAdvice = False

    '检验上次结果是否超标（函数内部通过参数控制是否检查，检查会影响性能）
    Call chkLastRual(lngKey)

    Exit Function
ErrHand:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If

End Function

Private Sub txt医嘱内容_GotFocus()
    Call zlControl.TxtSelAll(txt医嘱内容)
    Me.txt医嘱内容.IMEMode = 2
End Sub

Private Sub txt医嘱内容_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset

    On Error GoTo errH

    If KeyAscii = vbKeyReturn Then
        mblnSaveAdvice = True
        KeyAscii = 0
        If txt医嘱内容.Text = txt医嘱内容.Tag Then
'            zlCommFun.PressKey vbKeyTab
            SetFocusNextIndex Me.txt医嘱内容.TabIndex
            gintSelectFocus = 2
            Exit Sub
        End If

        With txt医嘱内容
            Set rsTmp = SelectDiagItem()
        End With

        If rsTmp Is Nothing Then '取消或无数据
            '恢复原值
            txt医嘱内容.Text = txt医嘱内容.Tag
            zlControl.TxtSelAll txt医嘱内容
            txt医嘱内容.SetFocus: gintSelectFocus = 2: Exit Sub
        End If
        '新项目的录入

        '根据选择项目设置缺省医嘱信息
        If AdviceInput(rsTmp) Then
'            DoEvents

            gintSelectFocus = 2
            '显示已缺省设置的值
            txt医嘱内容.Tag = txt医嘱内容.Text
            txt附加.Tag = txt附加.Text

            '处理检验仪器、标本号，对于已有标本的操作（重核或补填申请）则只重新赋医嘱ID
            If mintEditMode <= 1 Then
                Call LoadDefaultData
                Call SelectDefault

                With vsf2
                    If .Rows > 1 Then
                        .Row = 1
                    End If
                    .Col = 2
                    .ShowCell vsf2.Row, vsf2.Col
                    .SetFocus
                End With
            Else
                '赋医嘱ID
                Call SelectDefault

                Me.cbo开单科室.SetFocus
            End If
        Else
'            DoEvents
            gintSelectFocus = 2
            '恢复原值
            txt医嘱内容.Text = txt医嘱内容.Tag
            txt附加.Text = txt附加.Tag
            zlControl.TxtSelAll txt医嘱内容

            txt医嘱内容.SetFocus: gintSelectFocus = 2: Exit Sub
        End If
    ElseIf KeyAscii = Asc("*") Then
        KeyAscii = 0
    ElseIf KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txt医嘱内容_Validate(Cancel As Boolean)
    '恢复人为的改变
    If txt医嘱内容.Text <> txt医嘱内容.Tag Then
        txt医嘱内容.Text = txt医嘱内容.Tag
    End If
End Sub

Private Function SelectDiagItem() As ADODB.Recordset
'选择检验项目
    Dim strSQL As String
    Dim objPoint As POINTAPI

    strSQL = "Select Distinct A.ID,A.编码,A.名称,nvl(A.计算单位,'次') As 计算单位,nvl(A.标本部位,' ') As 标本部位," + _
        "Decode(A.类别,'H',Decode(A.操作类型,'1','护理等级','护理常规')," + _
        "'E',Decode(A.操作类型,'1','过敏试验','2','给药途径','3','中药煎法',4,'中药用法','其它')," + _
        "'Z',Decode(A.操作类型,'1','留观','2','住院','3','转科','4','术后','5','出院','6','转院','其它'),A.操作类型) As 项目特性,A.类别 As 类别ID,A.ID As 诊疗项目ID,nvl(执行频率,0) As 执行频率ID,nvl(计算方式,0) As 计算方式ID,nvl(执行安排,0) As 执行安排ID,nvl(计价性质,0) As 计价性质ID,nvl(执行科室,0) As 执行科室ID "
    strSQL = strSQL + "From 诊疗项目目录 A,诊疗项目别名 C,诊疗执行科室 D Where A.ID=C.诊疗项目ID And A.ID=D.诊疗项目ID And A.类别='C' And D.执行科室ID=" & ItemDeptID
    strSQL = strSQL + " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL) " + _
        "And A.服务对象 IN(" & PatientType & ",3,4)  And Nvl(A.适用性别,0) IN (" + _
        IIf(Me.cbo性别.Text Like "*男*", "1,0)", "2,0)") + _
        " And Nvl(A.执行频率,0) IN(0,1)" + _
        " And (upper(A.编码) Like '" + gstrMatch + UCase(txt医嘱内容) + "%' Or Upper(A.名称) Like '" + txt医嘱内容 + "%' Or Upper(C.简码) Like '" + UCase(txt医嘱内容) + "%')"

    Call ClientToScreen(txt医嘱内容.hWnd, objPoint)
    Set SelectDiagItem = zlDatabase.ShowSelect(Me, strSQL, 0, "选择申请项目", True, Me.txt医嘱内容.Text, "", True, True, True, objPoint.X * 15, objPoint.Y * 15, Me.txt医嘱内容.Height, False, True)
End Function

Private Function AdviceInput(Optional rsInput As ADODB.Recordset = Nothing) As Boolean
'功能：根据新输的诊疗项目(新增或更换)设置缺省的医嘱数据
'参数：rsInput=输入或选择返回的记录集
'返回：本次录入是否有效
    Dim rsTmp As ADODB.Recordset
    Dim strHelpText As String
    Dim strSQL As String
    Dim strExtData As String
    Dim blnOk As Boolean
    Dim t_Pati As TYPE_PatiInfoEx

    On Error GoTo errH

    '项目附加数据输入及输入合法性检查
    '---------------------------------------------------------------------------------------------------------------
    If Not rsInput Is Nothing Then txt医嘱内容.Text = rsInput!名称    '暂时显示

    '需要输入更多数据的一些项目
    '---------------------------------------------------------------------------------------------------------------
    '检验项目选择检验标本
    strHelpText = "检验项目"

    If Not rsInput Is Nothing Then
        strExtData = rsInput!诊疗项目ID & ";" & rsInput!标本部位    '新输入项目
    Else
        If Trim(Me.txt医嘱内容.Text) = "" Then mstrExtData = ""
        strExtData = mstrExtData    '新输入项目
    End If

    With t_Pati
        .str性别 = NeedName(cbo性别.Text)
    End With
    On Error Resume Next
    '接口改造：bytUseType 以前没传，现在传为0
    blnOk = frmAdviceEditEx.ShowMe(Me, Me.vsf2.hWnd, t_Pati, 2, 4, 0, 1, PatientType, , , , 0, strExtData, , , , , True)
    On Error GoTo errH

    If Not blnOk Then Exit Function
    If strExtData = "" Or Mid(strExtData, 1, 1) = ";" Then Exit Function

    '获取采集方式
    Set rsTmp = SelectCap(Split(Split(strExtData, ";")(0), ",")(0))
    If rsTmp Is Nothing Then
        MsgBox "没有定义标本采集方式，请到诊疗项目管理中设置。", vbInformation, gstrSysName
        Exit Function
    End If
    mlngCapID = rsTmp("ID")

    strSQL = "Select C.项目类别 From 诊疗项目目录 A,检验报告项目 B,检验项目 C " & _
        "Where A.ID=B.诊疗项目ID And B.报告项目ID=C.诊治项目ID And A.ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Split(Split(strExtData, ";")(0), ",")(0))
'    If rsTmp.EOF Then
'        mbln微生物项目 = False
'    Else
'        mbln微生物项目 = IIf(Nvl(rsTmp("项目类别"), 0) = 2, True, False)
'    End If

    mstrExtData = strExtData
    If Not rsInput Is Nothing Then Me.txt附加 = Trim(rsInput("标本部位"))

    Call AdviceSet检查手术(3, mstrExtData)
    txt医嘱内容.Text = Get检查手术名称(2, "")
    txt医嘱内容.Text = txt医嘱内容.Text & "(" & Split(mstrExtData, ";")(1) & ")"
    Me.txt附加 = Split(mstrExtData, ";")(1)

    '开嘱医生
    On Error Resume Next
    If Me.cbo医生.Text = "" Then Me.cbo医生.ListIndex = 0

    AdviceInput = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub InitSampleInfo(ByVal lngSampleID As Long)
'功能：根据标本ID，初始检验项目、申请科室医生等信息
'参数：rsInput=输入或选择返回的记录集
'返回：本次录入是否有效
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer, strTmp As String

    On Error GoTo errH

    strSQL = "Select 医嘱ID,申请科室ID,申请人,检验备注,病人ID,病人来源,标本形态,采样人,采样时间,接收人,接收时间 From 检验标本记录 Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngSampleID)
    If rsTmp.EOF Then Exit Sub
    mlng病人ID = Nvl(rsTmp("病人ID"), 0): If mlng病人ID > 0 Then Me.txt姓名.Tag = Me.txt姓名
'    If Not IsNull(rsTmp("医嘱ID")) Then ' And Nvl(rsTmp("病人来源"), 3) <> 3 Then
'        mblnSaveAdvice = False: SetAdviceEnable False
'    End If

    On Error Resume Next
'    If IsNull(rsTmp("医嘱ID")) And Not IsNull(rsTmp("申请科室ID")) Then
    If Not IsNull(rsTmp("申请科室ID")) Then
        Me.cbo开单科室.ListIndex = FindComboItem(Me.cbo开单科室, Nvl(rsTmp("申请科室ID"), 0))
        Me.cbo医生.Text = Nvl(rsTmp("申请人"))
    End If

    Me.cbo(0).Text = Nvl(rsTmp("标本形态"))

    Me.cbo(1).Text = Nvl(rsTmp("采样人"))
    DTP(0).Value = Format(zlCommFun.Nvl(rsTmp("采样时间"), zlDatabase.Currentdate), "YYYY-MM-DD HH:MM:SS")
    If Nvl(rsTmp("接收人")) = "" Then
        Me.cbo(2).Visible = False
        Me.DTP(2).Visible = False
        lbl(1).Visible = False
    Else
        Me.cbo(2).Visible = True
        Me.DTP(2).Visible = True
        lbl(1).Visible = True
        Me.cbo(2).Text = Nvl(rsTmp("接收人"))
        Me.DTP(2).Value = Nvl(rsTmp("接收时间"))
    End If
    On Error GoTo errH

    If IsNull(rsTmp("医嘱ID")) Then
        '无主标本
        strSQL = "Select 诊疗项目ID From 检验申请项目 Where 标本ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngSampleID)
    Else
        '有医嘱的
'        strSQL = "Select Distinct b.诊疗项目id From 检验项目分布 a , 病人医嘱记录 b  " & _
'                 " Where b.相关id = a.医嘱id(+) And a.标本id = [1] "
'        Set rsTmp =zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngSampleID)
        strSQL = "Select 诊疗项目ID From 病人医嘱记录 Where 相关ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CLng(rsTmp("医嘱ID")))
    End If
    If rsTmp.EOF Then Exit Sub

    i = 0: mstrExtData = ""

    Do While Not rsTmp.EOF
        i = i + 1
        mstrExtData = mstrExtData & "," & Nvl(rsTmp("诊疗项目ID"), 0)
'        If i = 3 Then Exit Do '最多显示3个项目

        rsTmp.MoveNext
    Loop


    If Len(mstrExtData) > 0 Then
        mstrExtData = Mid(mstrExtData, 2)
    Else
        Exit Sub
    End If

    If mlngSampleID > 0 Then
        strSQL = "select 标本类型 from 检验标本记录 where id = [1] and 标本类型 is not null "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, mlngSampleID)
    End If
    If rsTmp.EOF = True Then
        strSQL = "Select 标本类型,Sum(1) From (" & _
            "   Select A.ID,C.标本类型" & _
            "   From 诊疗项目目录 A,检验项目参考 C,检验报告项目 D" & _
            "   Where A.ID=D.诊疗项目ID And D.报告项目ID=C.项目ID" & _
            "   And A.ID In (" & mstrExtData & ")" & _
            " ) Group By 标本类型 Order By Sum(1) Desc "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    End If

    If rsTmp.EOF Then
        mstrExtData = mstrExtData & ";血清"
    Else
        mstrExtData = mstrExtData & ";" & rsTmp("标本类型")
    End If

    '获取采集方式
    Set rsTmp = SelectCap(Split(Split(mstrExtData, ";")(0), ",")(0))
    If rsTmp Is Nothing Then
        MsgBox "没有定义标本采集方式，请到诊疗项目管理中设置。", vbInformation, gstrSysName
        Exit Sub
    End If
    mlngCapID = rsTmp("ID")

'    strsql = "Select C.项目类别 From 诊疗项目目录 A,检验报告项目 B,检验项目 C " & _
'        "Where A.ID=B.诊疗项目ID And B.报告项目ID=C.诊治项目ID And A.ID=[1]"
'    Set rsTmp = zlDatabase.OpenSQLRecord(strsql, Me.Caption, Split(Split(mstrExtData, ";")(0), ",")(0))
'    If rsTmp.EOF Then
'        mbln微生物项目 = False
'    Else
'        mbln微生物项目 = IIf(Nvl(rsTmp("项目类别"), 0) = 2, True, False)
'    End If

    Call AdviceSet检查手术(3, mstrExtData)
'    txt医嘱内容.Text = Get检查手术名称(2, "")
'    txt医嘱内容.Text = txt医嘱内容.Text & "(" & Split(mstrExtData, ";")(1) & ")"
    Me.txt附加 = Split(mstrExtData, ";")(1)

    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function SelectCap(Optional ByVal lngItemID As Long = 0) As ADODB.Recordset
'获取采集方式
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim tmpRect As RECT

    On Error GoTo DBError

    strSQL = "Select Distinct A.ID,A.编码,A.名称 " + _
        "From 诊疗项目目录 A,诊疗用法用量 D Where A.ID=D.用法ID" + _
        " And A.类别='E' And A.操作类型='6'" & _
        " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL) " + _
        " And A.服务对象 IN(" & PatientType & ",3) And Nvl(A.适用性别,0) IN (" + _
        IIf(Me.cbo性别.Text Like "*男*", "1,0)", "2,0)") + _
        " And Nvl(A.执行频率,0) IN(0,1)" + _
        " And D.项目ID=" & lngItemID
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rsTmp.EOF Then
        strSQL = "Select Distinct A.ID,A.编码,A.名称 " + _
            "From 诊疗项目目录 A Where " + _
            " A.类别='E' And A.操作类型='6'" & _
            " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL) " + _
            " And A.服务对象 IN(" & PatientType & ",3) And Nvl(A.适用性别,0) IN (" + _
            IIf(Me.cbo性别.Text Like "*男*", "1,0)", "2,0)") + _
            " And Nvl(A.执行频率,0) IN(0,1)"
    End If
    If rsTmp.State = adStateOpen Then rsTmp.Close: Set rsTmp = New ADODB.Recordset
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If Not rsTmp.EOF Then Set SelectCap = rsTmp

    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub AdviceSet检查手术(ByVal int类型 As Integer, ByVal strDataIDs As String)
'功能：1.重新设置指定检查组合项目的部位行,用于新输入检查组合项目或修改部位
'      2.重新设置指定手术项目的附加手术及麻醉项目行,用于新输入手术项目或手术项目的附加手术及麻醉项目
'参数：int类型=1=处理检查部位项目,2=处理附加手术及麻醉项目
'      strDataIDs=检查:包含检查部位信息,手术:包含附加手术及麻醉项目信息,其中可能没有附加手术和麻醉
    Dim strSQL As String, i As Long
    Dim arrIDs As Variant

    On Error GoTo errH

    '处理检验项目
    strDataIDs = Mid(strDataIDs, 1, InStr(strDataIDs, ";") - 1)

    If strDataIDs <> "" Then
        If Not rsRelativeAdvice Is Nothing Then
            rsRelativeAdvice.Close
        Else
            Set rsRelativeAdvice = New ADODB.Recordset
        End If
        strSQL = "Select ID,编码,名称,nvl(标本部位,' ') As 标本部位," + _
        "类别,nvl(计价性质,0) As 计价性质,nvl(执行科室,0) As 执行科室,操作类型 From 诊疗项目目录 Where ID IN(" & strDataIDs & ")"
        Set rsRelativeAdvice = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Else
        If Not rsRelativeAdvice Is Nothing Then rsRelativeAdvice.Close: Set rsRelativeAdvice = Nothing
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Get检查手术名称(ByVal int类型 As Integer, ByVal txtMainAdvice As String) As String
'功能：重新生成检查手术内容的医嘱内容
'参数：int类型=1=处理检查部位项目,2=处理附加手术及麻醉项目
    Dim lngBegin As Long, i As Long
    Dim str麻醉 As String, strTmp As String
    Dim strDate As String

    If rsRelativeAdvice Is Nothing Or int类型 = 1 Then Get检查手术名称 = txtMainAdvice: Exit Function

    rsRelativeAdvice.MoveFirst
    Do While Not rsRelativeAdvice.EOF
        If Len(Trim(rsRelativeAdvice("名称"))) > 0 Then
            strTmp = strTmp & "," & rsRelativeAdvice("名称")
        End If

        rsRelativeAdvice.MoveNext
    Loop

    If strTmp <> "" Then
        Get检查手术名称 = IIf(Len(Trim(txtMainAdvice)) = 0, "", txtMainAdvice & " 及 ") & Mid(strTmp, 2)
    Else
        Get检查手术名称 = txtMainAdvice
    End If
End Function

Private Sub vsf2_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strPh As String, strMsg As String
    Dim mlngLoop As Long, lngItemRow  As Long

    Select Case Col
        Case 2
            If vsf2.RowData(Row) = -1 Then
                '手工标本号
                If gblnManualPH Then
                    strPh = ValidPH(vsf2.TextMatrix(Row, Col), strMsg)
                    If Len(strMsg) > 0 Then
                        MsgBox strMsg, vbOKOnly + vbInformation, gstrSysName
                        vsf2.TextMatrix(Row, Col) = ""
                    Else
                        vsf2.TextMatrix(Row, Col) = strPh
                    End If
                End If
            End If
        Case 5
            If Val(vsf2.RowData(Row)) = 0 Then Exit Sub

            If vsf2.TextMatrix(Row, Col) = "-1" Then
            '急诊
                vsf2.TextMatrix(Row, 2) = TransSampleNO_PH(Val(CalcNextCode(Val(vsf2.RowData(Row)), Row, 1)), vsf2.RowData(Row))
            Else
                vsf2.TextMatrix(Row, 2) = TransSampleNO_PH(Val(CalcNextCode(Val(vsf2.RowData(Row)), Row, 0)), vsf2.RowData(Row))
            End If
    End Select
End Sub

Private Sub vsf2_BeforeDeleteCell(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub vsf2_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub vsf2_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
'    zlCommFun.PressKey vbKeyTab
    SetFocusNextIndex Me.vsf2.TabIndex
    Cancel = True
End Sub

Private Sub vsf2_KeyPress(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer, Cancel As Boolean)
    On Error GoTo errH
    If mblnBarCode And KeyAscii = vbKeyReturn Then
        KeyAscii = 0: Me.cbo开单科室.SetFocus: gintSelectFocus = 2
    Else
        If KeyAscii = vbKeyReturn Then
            If Row + 1 = vsf2.Rows Then
                KeyAscii = 0: Me.cbo开单科室.SetFocus: gintSelectFocus = 2
            Else
                vsf2.Row = Row + 1
                vsf2.Col = 1
            End If
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsf2_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    On Error GoTo errH
    Select Case Col
        Case 2
            If KeyAscii = 13 Then
'                KeyAscii = 0: Me.DTP(0).SetFocus: gintSelectFocus = 2: Exit Sub
'                KeyAscii = 0: cbo开单科室.SetFocus: gintSelectFocus = 2: Exit Sub
                If Row + 1 = vsf2.Rows Then
                    KeyAscii = 0: Me.cbo开单科室.SetFocus: gintSelectFocus = 2
                Else
                    vsf2.Row = Row + 1
                    vsf2.Col = 2
                End If
            End If
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            If vsf2.RowData(vsf2.Row) <> -1 Then
                KeyAscii = FilterKeyAscii(KeyAscii, 99, "0123456789")
            Else
                '手工标本号
                If gblnManualPH Then
                    KeyAscii = FilterKeyAscii(KeyAscii, 99, "0123456789-")
                Else
                    KeyAscii = FilterKeyAscii(KeyAscii, 99, "0123456789")
                End If
            End If
    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
'检查医嘱内容的合法性
Private Function ValidAdvice() As Boolean
    ValidAdvice = True

    On Error Resume Next
    If txt姓名.Text = "" Then
        ValidAdvice = False
        MsgBox "请输入病人的姓名！", vbInformation, gstrSysName: gintSelectFocus = 2 'DoEvents
'        mintFocusItem = FocusItem.姓名
        txt姓名.SetFocus: Exit Function
    End If

    If Len(Trim(Me.txt医嘱内容)) = 0 And mblnCheckIn = False Then
        ValidAdvice = False
        MsgBox "必须输入申请项目！", vbInformation, gstrSysName: gintSelectFocus = 2 'DoEvents
'        mintFocusItem = FocusItem.医嘱内容
        Me.txt医嘱内容.SetFocus: Exit Function
    End If
    If Me.cbo开单科室.ListIndex = -1 And mblnCheckIn = False Then
        ValidAdvice = False
        MsgBox "请指定开单科室！", vbInformation, gstrSysName: gintSelectFocus = 2 'DoEvents
'        mintFocusItem = FocusItem.开单科室
        Me.cbo开单科室.SetFocus: Exit Function
    End If
    If Len(Trim(Me.cbo医生.Text)) = 0 And mblnCheckIn = False Then
        ValidAdvice = False
        MsgBox "请指定开单医生！", vbInformation, gstrSysName: gintSelectFocus = 2 'DoEvents
'        mintFocusItem = FocusItem.医生
        Me.cbo医生.SetFocus: Exit Function
    End If
End Function

Private Function SaveAdviceData(blnNewPatient As Boolean) As Long
    '参数                   blnNewPatient 是否是新病人
    Dim strSQL As String, strDate As String, strNO As String
    Dim lngAdviceID As Long, lngTmpID As Long, lngSendNO As Long
    Dim lngMaxSeq As Long, iSendSeq As Integer
    Dim rsTmp As New ADODB.Recordset
    Dim lng开嘱科室ID As Long, lng病人ID As Long, strDoctor As String, i As Integer
    Dim str执行科室ID As String, str执行科室ID1 As String, lngDept As Long
    Dim rsCard As ADODB.Recordset
    Dim tmpstr类别 As String, tmplngClinicID As Long, tmpint计价特性 As Integer, tmpint执行性质 As Integer
    Dim rsDept As ADODB.Recordset
    Dim lngPatientHomePage As Long
    Dim blnNewpatinet As Boolean    '新病人
    Dim blnPatientType As Boolean   '是标识为外来病人

    On Error GoTo ErrHand
    blnPatientType = zlDatabase.GetPara("所有登记病人标识为外来", 100, 1208, 0)


    '保存病人信息
    strDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    blnNewpatinet = blnNewPatient
    '保存医嘱并发送
    lngAdviceID = zlDatabase.GetNextId("病人医嘱记录")
    '得到最大医嘱序号
    gstrSql = "select max(序号) as 序号 from 病人医嘱记录 where 病人id = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.ControlBox, mlng病人ID)
    If rsTmp.EOF = False Then
        lngMaxSeq = Val(Nvl(rsTmp("序号"), 0))
    Else
        lngMaxSeq = 0
    End If

    lng开嘱科室ID = Me.cbo开单科室.ItemData(Me.cbo开单科室.ListIndex)
    strDoctor = NeedName(Me.cbo医生.Text)
    If ItemDeptID = 0 Then
        MsgBox "你没有选择一个执行科室不能进行保存，请在主界面上选择一个科室再进行保存！", vbInformation, Me.Caption
        SaveAdviceData = -1
        Exit Function
    Else
        str执行科室ID = ItemDeptID
    End If

    iSendSeq = 1
    '检验项目将采集方式作为主医嘱
    tmplngClinicID = mlngCapID
    '取采集方式的执行部门
    str执行科室ID1 = "NULL"

    lngSendNO = zlDatabase.GetNextNo(10)
    strNO = zlDatabase.GetNextNo(IIf(PatientType = 2, 14, 13))

    gstrSql = "select nvl(max(主页ID),0) as 主页ID from 病案主页 where 病人ID = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, gstrSysName, mlng病人ID)
    lngPatientHomePage = rsTmp("主页ID")

    '门诊和住院病人登记是否标识为外来病人
    If blnPatientType = False Then
        If blnNewpatinet = True Then
            PatientType = 3
        End If
    Else
        PatientType = 3
    End If

    If mblnCheckIn = True And rsRelativeAdvice Is Nothing Then Exit Function

    '保存相关医嘱
    If Not rsRelativeAdvice Is Nothing Then
        lngMaxSeq = lngMaxSeq + 1
        rsRelativeAdvice.MoveFirst
        Do While Not rsRelativeAdvice.EOF
            lngTmpID = zlDatabase.GetNextId("病人医嘱记录")
            With rsRelativeAdvice
                strSQL = "ZL_病人医嘱记录_Insert(" & lngTmpID & "," & lngAdviceID & "," & _
                    lngMaxSeq & "," & PatientType & "," & mlng病人ID & "," & IIf(lngPatientHomePage = 0, "NULL", lngPatientHomePage) & "," & _
                    "0,1," & _
                    "1,'" & .Fields("类别") & "'," & _
                    .Fields("ID") & ",NULL,NULL,NULL,1," & _
                    "'" & Replace(.Fields("名称"), "'", "''") & "',''," & _
                    "'" & Me.txt附加 & "','一次性',NULL,NULL,'',NULL," & _
                    .Fields("计价性质") & "," & _
                    str执行科室ID & "," & _
                    .Fields("执行科室") & ",0," & strDate & ",NULL," & _
                    IIf(Val(Me.txtPatientDept.Tag) = 0, lng开嘱科室ID, Val(Me.txtPatientDept.Tag)) & "," & lng开嘱科室ID & ",'" & strDoctor & "'," & _
                    "Sysdate,''," & lngAdviceID & ")"
                    zlDatabase.ExecuteProcedure strSQL, Me.Caption

                iSendSeq = iSendSeq + 1
                strSQL = "ZL_病人医嘱发送_Insert(" & _
                    lngTmpID & "," & lngSendNO & "," & PatientType & ",'" & strNO & "'," & _
                    iSendSeq & ",1,NULL,NULL," & _
                    "Sysdate+1/(24*3600)," & _
                    "0," & str执行科室ID & ",0,0)"
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
                .MoveNext
            End With
        Loop
    End If
    '检验申请的采集方式放到最后
    lngMaxSeq = lngMaxSeq + 1
    strSQL = "ZL_病人医嘱记录_Insert(" & lngAdviceID & ",NULL," & _
        lngMaxSeq & "," & PatientType & "," & mlng病人ID & "," & IIf(lngPatientHomePage = 0, "NULL", lngPatientHomePage) & "," & _
        "0,1," & _
        "1,'E'," & mlngCapID & ",NULL,NULL,NULL,1," & _
        "'" & Replace(Me.txt医嘱内容, "'", "''") & "',''," & _
        "'" & Me.txt附加 & "','一次性',NULL,NULL,'',NULL,2," & _
        str执行科室ID & ",3,0," & strDate & ",NULL," & _
        IIf(Val(Me.txtPatientDept.Tag) = 0, lng开嘱科室ID, Val(Me.txtPatientDept.Tag)) & "," & lng开嘱科室ID & ",'" & strDoctor & "'," & _
        "Sysdate,''," & lngAdviceID & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption

    iSendSeq = iSendSeq + 1
    '发送主医嘱
    strSQL = "ZL_病人医嘱发送_Insert(" & _
        lngAdviceID & "," & lngSendNO & "," & PatientType & ",'" & strNO & "'," & _
        iSendSeq & ",1,NULL,NULL," & _
        "Sysdate+1/(24*3600)," & _
        "0," & str执行科室ID & ",0,1)"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption

    SaveAdviceData = lngAdviceID

    Exit Function
ErrHand:

    Err.Raise Err.Number, "标本核收"
End Function

Private Sub SetPatientInfoWrite(blnTrue As Boolean)
    '功能         设置病人信息是否可以写入
    '参数         是否可写入
    Dim blnModifyInfo As Boolean                        '是否能修改病人信息

    blnModifyInfo = zlDatabase.GetPara("登记时可直接输入病人信息", 100, 1208, 0)
    If blnModifyInfo = 0 Then blnTrue = True

    Me.txtID.Locked = blnTrue
    Me.txtID.Enabled = Not blnTrue
    Me.txtBed.Locked = blnTrue
    Me.txtBed.Enabled = Not blnTrue
    Me.txtPatientDept.Locked = blnTrue
    Me.txtPatientDept.Enabled = Not blnTrue

End Sub
Private Function GetPatientInfo(lngID As String) As ADODB.Recordset

    gstrSql = "Select 1 As Patienttype, 0 As 主页id, A.病人科室, A.姓名, Decode(B.名称, Null, A.性别, B.编码 || '-' || A.性别) As 性别," & vbNewLine & _
                "       A.年龄, A.病人id, C.住院号, C.门诊号, A.床号 as 当前床号,a.标识号,a.申请人 as 医生,Zl_Age_Calc(A.病人ID) as 年龄1 " & vbNewLine & _
                " From 检验标本记录 A, 性别 B, 病人信息 C" & vbNewLine & _
                " Where A.性别 = B.名称(+) And A.病人id = C.病人id and " & IIf(IsNumeric(lngID) = False, " 1 = 2 and 标识号 = [1] ", " 标识号 = [1] ")

    Set GetPatientInfo = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(lngID))
End Function
Public Function zlRefresh_bak(ByVal lngSampleID As Long) As Boolean
'显示标本申请信息
'lngSampleID：标本记录ID
    Dim rs As New ADODB.Recordset
    Dim mstrSql As String

    On Error GoTo ErrHand


    mstrSql = "Select A.主页id, A.病人id, A.姓名, L.编码 || '-' || A.性别 As 性别, A.年龄," & vbNewLine & _
                "       Decode(A.病人来源, 3, To_Char(A.NO), 1, To_Char(A.门诊号), 2, To_Char(A.住院号), 4, To_Char(A.门诊号)) As 病人号," & vbNewLine & _
                "       Decode(A.住院号, Null, Null, A.床号) As 床号, A.标本类型, A.核收时间," & vbNewLine & _
                "       Decode(A.病人来源, 1, '门诊', 2, '住院', 3, '外来', 4, '体检') As 病人来源," & vbNewLine & _
                "       Decode(A.仪器id, Null," & vbNewLine & _
                "               To_Char(Trunc(A.标本序号 / 10000) + 1, '0000') || '-' || To_Char(Mod(A.标本序号, 10000), '0000')," & vbNewLine & _
                "               A.标本序号) As 标本序号, A.申请人 As 开嘱医生, A.申请时间 As 开嘱时间," & vbNewLine & _
                "       Nvl(A.病人科室, '未知') As 病人科室, Nvl(B.名称, '未知') As 开嘱科室, Nvl(C.名称, '手工') As 检验仪器," & vbNewLine & _
                "       A.检验项目 As 医嘱内容, A.采样人, A.采样时间, A.标本形态, A.仪器id, Nvl(A.标本类别, 0) As 标本类别," & vbNewLine & _
                "       Nvl(A.标本类别, 0) As 标本类别, A.检验时间, A.标识号, A.床号 As 床号1" & vbNewLine & _
                "From 检验标本记录 A, 部门表 B, 检验仪器 C, 性别 L" & vbNewLine & _
                "Where A.性别 = L.名称(+) And A.申请科室id = B.ID(+) And A.仪器id = C.ID(+) and a.id = [1] "

    Set rs = zlDatabase.OpenSQLRecord(mstrSql, Me.Caption, lngSampleID)

    On Error Resume Next
    If rs.EOF Then
        ClearItem

        Me.cbo开单科室.ListIndex = -1: Me.cbo医生.ListIndex = -1
        Me.cbo(0).ListIndex = -1: Me.cbo(1).ListIndex = -1
    Else
        Me.txt姓名 = Nvl(rs("姓名"))
        Me.cbo性别.Text = Nvl(rs("性别"))
'        Me.txt年龄 = IIf(IsNull(rs("年龄")), "", Val(rs("年龄"))): If Me.txt年龄 = "0" Then Me.txt年龄 = ""
'        Me.txt年龄 = IIf(IsNull(rs("年龄")), "", IIf(IsNumeric(rs("年龄")), Val(rs("年龄")), Mid(rs("年龄"), 1, Len(rs("年龄")) - 1))): If Me.txt年龄 = "0" Then Me.txt年龄 = ""

        If IsNull(rs("年龄")) Then
            Me.txt年龄 = ""
        Else
            Me.txt年龄 = Val(rs("年龄"))
            If Me.txt年龄 = 0 Then Me.txt年龄 = ""
        End If

        If IsNull(rs("年龄")) = True Then
            Me.cboAge.Text = "岁"
        Else
            If Val(rs("年龄")) = 0 Then
                Me.cboAge.Text = rs("年龄")
            Else
                Me.cboAge.Text = Mid(rs("年龄"), Len(CStr(Val(rs("年龄")))) + 1)
            End If
        End If
        'Me.cboAge.Text = IIf(IsNull(rs("年龄")), "岁", Right(rs("年龄"), 1))
        If cboAge.ListIndex = -1 Then cboAge.ListIndex = 0
        Me.txtPatientDept = Nvl(rs("病人科室"))
        Me.txtID = Nvl(rs("病人号"), Nvl(rs("标识号")))
        Me.txtBed = Nvl(rs("床号"), Nvl(rs("床号1")))

        Me.cbo开单科室.Text = Nvl(rs("开嘱科室"))
        Me.cbo医生.Text = Nvl(rs("开嘱医生"))
        Me.txt附加 = Nvl(rs("标本类型"))

        Me.DTP(1).Value = rs("核收时间")
        Me.cbo(1).Text = Nvl(rs("采样人"))
        Me.DTP(0).Value = rs("采样时间")

        Me.cbo(0).Text = Nvl(rs("标本形态"))

        With vsf2
            .Rows = 2
            .RowData(1) = Nvl(rs("仪器ID"), -1)
            .TextMatrix(1, 1) = Nvl(rs("检验仪器"))
            .TextMatrix(1, 2) = Nvl(rs("标本序号"))
            .TextMatrix(1, 5) = IIf(rs("标本类别") = 0, 0, -1)
        End With
        Me.txt医嘱内容 = ""
        Do While Not rs.EOF
            Me.txt医嘱内容 = Me.txt医嘱内容 & "," & Nvl(rs("医嘱内容"))

            rs.MoveNext
        Loop
        rs.MoveFirst
        If Len(Me.txt医嘱内容) > 0 Then Me.txt医嘱内容 = Mid(Me.txt医嘱内容, 2)
        If Nvl(rs!病人来源, 0) = 3 Then
            lblCash.Caption = ""
        Else
            lblCash.Caption = IIf(CheckChargeState(lngSampleID, False, False), "收", "")
        End If
    End If

    SetPatientInfoWrite True
    zlRefresh_bak = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub SetFocusNextIndex(TabIndex As Integer)
    '功能：         模拟按下Tab键后跳到下一个控件
    '参数:          TabIndex 当前控件的TabIndex
    Dim objThis As Object
    Dim intLoop As Integer
    On Error Resume Next
    For intLoop = TabIndex + 1 To Me.Count - 1
        For Each objThis In Me.Controls
            If objThis.TabIndex = intLoop Then
                If TypeName(objThis) = "VsfGrid" Then
                    objThis.Col = 2
                    objThis.ShowCell vsf2.Row, vsf2.Col
                    objThis.SetFocus
                    gintSelectFocus = 2
                    Exit Sub
                Else
                    If objThis.Enabled = True And objThis.Visible = True Then
                        objThis.SetFocus
                        gintSelectFocus = 2
                        Exit Sub
                    End If
                End If

            End If
        Next
    Next
End Sub

Private Function ShowCharge(ByVal lngKey As Long) As Integer
    '功能: 根据医嘱显示收费状态
    '参数: 医嘱ID
    '返回: -1=没有收费单 0=划价单 1=已收费
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    Dim strReplace As String

    strSQL = "select 病人来源 from 检验标本记录 where id = [1] "
    Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlLisWork", lngKey)
    If rs.EOF = True Then Exit Function

    If rs("病人来源") <> 2 Then
        strReplace = "门诊费用记录"
    End If

    ShowCharge = -1
    strSQL = _
        "select NVL(A.记录状态,-1) As 记录状态 " & _
              "from 住院费用记录 A, " & _
              "( " & _
                   "select No,记录性质 from 病人医嘱发送 where 医嘱id IN (Select ID From 病人医嘱记录 A,(Select 医嘱id From 检验标本记录 Where ID= [1] Union Select 医嘱id From 检验项目分布 Where 标本id= [1]) B where B.医嘱id =A.相关id and A.诊疗类别 = 'C'  ) " & _
                   "Union " & _
                   "select No,记录性质 from 病人医嘱附费 where 医嘱id IN (Select ID From 病人医嘱记录 A,(Select 医嘱id From 检验标本记录 Where ID= [1] Union Select 医嘱id From 检验项目分布 Where 标本id= [1]) B where B.医嘱id =A.相关id and A.诊疗类别 = 'C'  ) " & _
              ") B " & _
            "Where A.NO = B.NO and mod(a.记录性质,10) = b.记录性质 Order By NVL(A.记录状态,-1)"



    If strReplace <> "" Then
        strSQL = Replace$(strSQL, "住院费用记录", strReplace)
    End If
    Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlLisWork", lngKey)

    If rs.BOF Then Exit Function

    ShowCharge = rs("记录状态").Value
End Function

Private Function CreatePatient() As Boolean
    '功能建立病人信息
     '保存病人信息
    Dim strDate As String
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim strCostType As String
    Dim i As Long
    
    Dim strAge As String
    Dim strInfo As String
    Dim lngTmp As Long
    On Error GoTo errH

    strDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    If mlng病人ID <> 0 Then
        strSQL = "select 病人ID from 病人信息 where 病人id = [1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID)
        If rsTmp.EOF = True Then
            mlng病人ID = 0: PatientType = 1
        End If
    End If
    If PatientType = 1 And mlng病人ID <= 0 Then '门诊病人
        If mlng病人ID > 0 Then '已有的病人
'            strsql = _
                "zl_挂号病人病案_INSERT(3," & mlng病人ID & ",Null," & _
                "'',''," & _
                "'" & txt姓名.Text & "','" & NeedName(cbo性别.Text) & "','" & txt年龄.Text & Replace(Replace(Me.cboAge.Text, "成人", "岁"), "婴儿", "天") & txt年龄1.Text & "'," & _
                "'自费','自费'," & _
                "'','',''," & _
                "'','','',0,'','','','',''," & strDate & ",NULL)"
'            strsql = "Zl_检验病人病案_Insert(3," & mlng病人ID & ",'" & txt姓名.Text & "','" & NeedName(cbo性别.Text) & "','" & _
                        txt年龄.Text & Replace(Replace(Me.cboAge.Text, "成人", "岁"), "婴儿", "天") & txt年龄1.Text & "')"
        Else '新病人
            If txt年龄.Locked = False Then
                strAge = txt年龄.Text
                If IsNumeric(strAge) Then strAge = strAge & cboAge.Text & txt年龄1.Text
                strInfo = CheckAge(strAge)
                If InStr(1, strInfo, "|") > 0 Then
                    lngTmp = Val(Split(strInfo, "|")(0)) '1禁止,0提示
                    strInfo = Split(strInfo, "|")(1)
                    If lngTmp = 1 Then
                        MsgBox strInfo, vbInformation, gstrSysName
                        If txt年龄.Enabled And txt年龄.Visible Then txt年龄.SetFocus: Exit Function
                    End If
                End If
            End If
            mlng病人ID = zlDatabase.GetNextNo(1)
'            strsql = _
                "zl_挂号病人病案_INSERT(1," & mlng病人ID & ",Null," & _
                "'',''," & _
                "'" & txt姓名.Text & "','" & NeedName(cbo性别.Text) & "','" & txt年龄.Text & Replace(Replace(Me.cboAge.Text, "成人", "岁"), "婴儿", "天") & txt年龄1.Text & "'," & _
                "'自费','自费'," & _
                "'','',''," & _
                "'','','',0,'','','','',''," & strDate & ",NULL)"
            strSQL = "select 名称,缺省标志 from 费别 order by 编码"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlLisWork")
            Do While Not rsTmp.EOF
                i = i + 1
                If i = 1 Then
                    strCostType = rsTmp("名称")
                End If
                If rsTmp("缺省标志") = 1 Then
                    strCostType = rsTmp("名称")
                    Exit Do
                End If
                rsTmp.MoveNext
            Loop
            If strCostType = "" Then strCostType = "自费"
            strSQL = "Zl_检验病人病案_Insert(1," & mlng病人ID & ",'" & txt姓名.Text & "','" & NeedName(cbo性别.Text) & "','" & _
                        txt年龄.Text & Replace(Replace(Me.cboAge.Text, "成人", "岁"), "婴儿", "天") & txt年龄1.Text & "','" & strCostType & "')"
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
        End If

        CreatePatient = True
    End If
    Exit Function
errH:
    mlng病人ID = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub chkLastRual(lngKey As Long)
    '功能   检查上次结果是否超标
    '参数   lnkey = 标本id
    Dim blnChk  As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim rsChk As New ADODB.Recordset

    blnChk = zlDatabase.GetPara("核收时提示上次超标结果", 100, 1208, False)

    If blnChk = False Then Exit Sub

    On Error GoTo errH

    gstrSql = "select b.检验项目id from 检验标本记录 a , 检验普通结果 b where a.id = b.检验标本id and a.id = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngKey)

    Do Until rsTmp.EOF
        If Nvl(rsTmp("检验项目ID")) > 0 Then
            '逐个指标检查
            gstrSql = "Select ID, 检验项目id, 警戒下限, 警戒上限, 检验结果,缩写" & vbNewLine & _
                        "From (Select ID, 检验项目id, 警戒下限, 警戒上限, 检验结果,缩写" & vbNewLine & _
                        "       From (Select A.Id, B.检验项目id,f.警示下限 as 警戒下限,f.警示上限 as  警戒上限, B.检验结果,缩写" & vbNewLine & _
                        "              From 检验标本记录 A, 检验普通结果 B," & vbNewLine & _
                        "                   (Select A.病人id, B.检验项目id" & vbNewLine & _
                        ",Zl_To_Number(Zl_Get_Reference(1, b.检验项目id, a.标本类型, Decode(a.性别, '男', 1, '女', 2, 0), a.出生日期,a.仪器id, a.年龄)) as 参考ID " & vbNewLine & _
                        "                     From 检验标本记录 A, 检验普通结果 B" & vbNewLine & _
                        "                     Where A.Id = B.检验标本id And A.Id = [1]) C, 检验项目 D,检验项目参考 F " & vbNewLine & _
                        "              Where A.Id = B.检验标本id And A.病人id = C.病人id And B.检验项目id = C.检验项目id And A.核收时间 Between Sysdate - 1 And Sysdate And" & vbNewLine & _
                        "                    A.Id < [1] And B.检验项目id = [2] And B.检验项目id = D.诊治项目id And c.参考id=F.ID(+)" & vbNewLine & _
                        "              Order By ID Desc)" & vbNewLine & _
                        "       Where Rownum = 1)" & vbNewLine & _
                        "Where 检验结果 <= 警戒下限 Or 检验结果 >= 警戒上限"

            Set rsChk = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngKey, CLng(Val(Nvl(rsTmp("检验项目ID"), 0))))
            If rsChk.EOF = False Then
                MsgBox "当前病人的上一个标本有结果<" & rsChk("缩写") & ">超标！请注意！", vbInformation, Me.Caption
                Exit Sub
            End If
        End If
        rsTmp.MoveNext
    Loop

    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function CheckIsInclude(strSource As String, strTarge As String) As Boolean
    '检查strSource中的每一个字符是否在strTarge中
    Dim i As Long
    CheckIsInclude = False

    Select Case strTarge
    Case "日期"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "时间"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"";|\=+-_)(*&^%$#@!`~"
    Case "日期时间"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"";|\=+_)(*&^%$#@!`~"
    Case "整数"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "小数"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "正整数"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+-_)(*&^%$#@!`~"
    Case "正小数"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/'"":;|\=+-_)(*&^%$#@!`~"
    Case "可打印字符"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/."":;|\=+-_)(*&^%$#@!`~0123456789"
    End Select
    For i = 1 To Len(strSource)
        If InStr(strTarge, Mid(strSource, i, 1)) <= 0 Then Exit Function
    Next
    CheckIsInclude = True
End Function
Public Sub SetPara()
    '主界面变化了参数
    mMakeNoRule = zlDatabase.GetPara("标本序号生成规则", 100, 1208, "今  天")
    mblnLoadLastAdvice = zlDatabase.GetPara("登记时保留上一次申请项目", 100, 1208, False)
    mblnCheckIn = Val(zlDatabase.GetPara("登记时不需要输入项目", 100, 1208, 0))
    mintItemRule = Val(zlDatabase.GetPara("手工项目按项目累加标本号", 100, 1208, 0))
    mSendReport = zlDatabase.GetPara("使用二级报告审核", 100, 1208, 0)
    mblnEmerge = Val(zlDatabase.GetPara("急诊标本", 100, 1208, 0))
    mbln划价单模式 = InStr(GetSysParVal(80, ""), "C") > 0
End Sub
Private Function GetSampleData(lngKey As Long) As String
    Dim rsTmp As New ADODB.Recordset
    '取标本的数据组成字串
    gstrSql = "Select Distinct 检验项目id, Decode(分布医嘱, Null, 标本医嘱, 分布医嘱) 医嘱id, 检验结果, 结果标志, 结果参考, 诊疗项目id, 排列序号" & vbNewLine & _
                "From (Select B.检验项目id, Null 标本医嘱, C.医嘱id 分布医嘱, B.检验结果, B.结果标志, B.结果参考, B.诊疗项目id, B.排列序号" & vbNewLine & _
                "       From 检验标本记录 A, 检验普通结果 B, 检验项目分布 C" & vbNewLine & _
                "       Where A.Id = B.检验标本id And A.Id = C.标本id And B.检验项目id = C.项目id And A.Id = [1]" & vbNewLine & _
                "       Minus" & vbNewLine & _
                "       Select B.检验项目id, A.医嘱id 标本医嘱, Null 分布医嘱, B.检验结果, B.结果标志, B.结果参考, B.诊疗项目id, B.排列序号" & vbNewLine & _
                "       From 检验标本记录 A, 检验普通结果 B" & vbNewLine & _
                "       Where A.Id = B.检验标本id And A.Id = [1]) order by 排列序号 "


    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngKey)
    Do Until rsTmp.EOF
        GetSampleData = GetSampleData & "|" & rsTmp("检验项目id") & Chr(1) & rsTmp("医嘱id") & Chr(1) & _
                        rsTmp("检验结果") & Chr(1) & rsTmp("结果标志") & Chr(1) & rsTmp("结果参考") & Chr(1) & _
                        rsTmp("诊疗项目id") & Chr(1) & rsTmp("排列序号")

        rsTmp.MoveNext
    Loop
    If GetSampleData <> "" Then
        GetSampleData = Mid$(GetSampleData, 2)
    End If
End Function

Private Function GetApplicationFormShowType() As Boolean
    If Not mobjLisInsideComm Is Nothing Then
        GetApplicationFormShowType = mobjLisInsideComm.GetApplicationFormShowType()
    Else
        GetApplicationFormShowType = False
    End If
End Function
