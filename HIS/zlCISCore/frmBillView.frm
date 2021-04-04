VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBillView 
   BorderStyle     =   0  'None
   ClientHeight    =   9120
   ClientLeft      =   0
   ClientTop       =   -285
   ClientWidth     =   9075
   Icon            =   "frmBillView.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   WindowState     =   1  'Minimized
   Begin VB.PictureBox picDoc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7455
      Left            =   0
      ScaleHeight     =   7455
      ScaleWidth      =   8415
      TabIndex        =   14
      Top             =   1080
      Width           =   8415
      Begin VB.PictureBox picFile 
         BorderStyle     =   0  'None
         Height          =   6495
         Left            =   840
         ScaleHeight     =   6495
         ScaleWidth      =   6735
         TabIndex        =   27
         Top             =   2040
         Width           =   6735
         Begin zl9CISCore.ctrlPatientFile ProFile1 
            Height          =   5175
            Index           =   0
            Left            =   480
            TabIndex        =   13
            Top             =   120
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   9128
            Border_Width    =   0
         End
      End
      Begin VB.PictureBox picAdvice 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   1815
         Left            =   0
         ScaleHeight     =   1815
         ScaleWidth      =   9255
         TabIndex        =   15
         Top             =   0
         Width           =   9255
         Begin VB.TextBox txt附加 
            Height          =   300
            Left            =   6440
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   0
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CheckBox chk开始时间 
            BackColor       =   &H80000005&
            Caption         =   "要求时间"
            Height          =   225
            Left            =   315
            TabIndex        =   4
            ToolTipText     =   "是否安排时间"
            Top             =   420
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.TextBox txt单量 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   7050
            MaxLength       =   3
            TabIndex        =   10
            Top             =   1080
            Width           =   1380
         End
         Begin VB.TextBox txt频率 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1350
            TabIndex        =   8
            Top             =   1080
            Width           =   2500
         End
         Begin VB.TextBox txt总量 
            Alignment       =   1  'Right Justify
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   4725
            MaxLength       =   3
            TabIndex        =   9
            Top             =   1080
            Width           =   1380
         End
         Begin VB.CheckBox chk紧急 
            BackColor       =   &H80000005&
            Caption         =   "紧急(&J)"
            Height          =   225
            Left            =   4200
            TabIndex        =   6
            Top             =   405
            Width           =   945
         End
         Begin VB.CommandButton cmdExt 
            Height          =   285
            Left            =   8040
            Picture         =   "frmBillView.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   3
            TabStop         =   0   'False
            ToolTipText     =   "选择检验标本"
            Top             =   0
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.CommandButton cmdSel 
            Caption         =   "…"
            Height          =   285
            Left            =   5280
            TabIndex        =   1
            TabStop         =   0   'False
            ToolTipText     =   "选择项目(*)"
            Top             =   0
            Width           =   285
         End
         Begin VB.ComboBox cbo执行科室 
            Enabled         =   0   'False
            Height          =   300
            ItemData        =   "frmBillView.frx":0102
            Left            =   1350
            List            =   "frmBillView.frx":0104
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1440
            Width           =   2500
         End
         Begin VB.TextBox txt医嘱内容 
            Height          =   300
            Left            =   1350
            MaxLength       =   1000
            MultiLine       =   -1  'True
            TabIndex        =   0
            Top             =   0
            Width           =   3945
         End
         Begin VB.ComboBox cbo医生 
            Height          =   300
            Left            =   5940
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   1425
            Width           =   1590
         End
         Begin VB.TextBox txt医生嘱托 
            Height          =   300
            Left            =   1350
            MaxLength       =   100
            TabIndex        =   7
            Top             =   720
            Width           =   4335
         End
         Begin VB.CommandButton cmd频率 
            Enabled         =   0   'False
            Height          =   240
            Left            =   3575
            Picture         =   "frmBillView.frx":0106
            Style           =   1  'Graphical
            TabIndex        =   16
            TabStop         =   0   'False
            ToolTipText     =   "选择项目(F4)"
            Top             =   1110
            Width           =   270
         End
         Begin MSComCtl2.DTPicker txt开始时间 
            Height          =   300
            Left            =   1350
            TabIndex        =   5
            Top             =   360
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   62193667
            CurrentDate     =   38022
         End
         Begin VB.Line lineTitleSplit 
            BorderColor     =   &H80000000&
            X1              =   400
            X2              =   1440
            Y1              =   320
            Y2              =   320
         End
         Begin VB.Label lbl附加 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "检查部位"
            Height          =   180
            Left            =   5640
            TabIndex        =   28
            Top             =   45
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Label lbl单量 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "每次"
            Height          =   180
            Left            =   6660
            TabIndex        =   26
            Top             =   1140
            Width           =   360
         End
         Begin VB.Label lbl单量单位 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0FF&
            BackStyle       =   0  'Transparent
            Height          =   180
            Left            =   8460
            TabIndex        =   25
            Top             =   1140
            Width           =   15
         End
         Begin VB.Label lbl频率 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "频率"
            Height          =   180
            Left            =   960
            TabIndex        =   24
            Top             =   1140
            Width           =   360
         End
         Begin VB.Label lbl总量单位 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   180
            Left            =   6150
            TabIndex        =   23
            Top             =   1140
            Width           =   15
         End
         Begin VB.Label lbl总量 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "共"
            Height          =   180
            Left            =   4335
            TabIndex        =   22
            Top             =   1140
            Width           =   180
         End
         Begin VB.Label lbl执行科室 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "执行科室"
            Height          =   180
            Left            =   600
            TabIndex        =   21
            Top             =   1500
            Width           =   720
         End
         Begin VB.Label lbl医嘱内容 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "申请项目"
            Height          =   180
            Left            =   600
            TabIndex        =   20
            Top             =   45
            Width           =   720
         End
         Begin VB.Label lbl开始时间 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "要求时间"
            Height          =   180
            Left            =   600
            TabIndex        =   19
            Top             =   435
            Width           =   720
         End
         Begin VB.Label lbl开嘱医生 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "申请医生"
            Height          =   180
            Left            =   5175
            TabIndex        =   18
            Top             =   1485
            Width           =   720
         End
         Begin VB.Label lbl医生嘱托 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "医生嘱托"
            Height          =   180
            Left            =   585
            TabIndex        =   17
            Top             =   795
            Width           =   720
         End
         Begin VB.Line lineSplit 
            X1              =   0
            X2              =   1080
            Y1              =   1800
            Y2              =   1800
         End
      End
   End
End
Attribute VB_Name = "frmBillView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private PatientID As String '病人ID
Private CheckID As String '病案ID或挂号单ID
Private PatientType As Integer '0=门诊病人 1=住院病人
Private FileTypeID As String '病历模板文件ID
Private bSample As Boolean '是否示范
Private blnMoved As Boolean
Private prbRefresh As Object
Attribute prbRefresh.VB_VarHelpID = -1

Private AdviceID As Long '医嘱ID
Private sCheckNo As String '发送单据号
Private iRecordType As Integer '记录性质
Private alngFileID(1) As Long '申请和报告ID
Private intType As Integer '诊疗类别:-1=其他、0=检查组合、1=手术、2=中药、3=检验
Private iTabIndex As Integer

'医嘱编辑
Private rsRelativeAdvice As ADODB.Recordset '相关医嘱
Private strExtData As String '附加项目

Private iCurrElementIndex As Integer '当前元素顺序号

Public Sub ShowMe(ByVal lng医嘱ID As Long, Optional objPrbRefresh As Object, Optional DataMoved As Boolean = False)
    Dim rsTmp As New ADODB.Recordset, i As Integer
    Dim strDiagName As String '诊疗项目名称
    Dim strDrAdvice As String '医生嘱托
    Dim bAllowEdit As Boolean
    Dim strSQL As String
    
    If AdviceID = lng医嘱ID Then Exit Sub
    
    AdviceID = lng医嘱ID
    blnMoved = DataMoved
    On Error Resume Next
    '初始化
    Set prbRefresh = objPrbRefresh
    ClearForm
    
    strSQL = "Select a.病人ID,a.主页ID,a.挂号单,Decode(a.主页ID,Null,0,1),b.ID,b.名称,a.医生嘱托," + _
        "医嘱内容,开始执行时间,紧急标志,执行频次,总给予量,单次用量,c.编码 As 科室编码,c.名称 As 科室名称,开嘱医生,nvl(b.计算单位,' ') As 计算单位,b.类别,nvl(a.标本部位,' ') As 标本部位,Nvl(a.申请ID,0) As 申请ID,d.病历文件ID " + _
        "From 病人医嘱记录 a,诊疗项目目录 b,部门表 c,诊疗单据应用 d Where (a.ID=[1] Or a.相关ID=[1]) And a.诊疗项目ID=b.ID And a.执行科室ID=c.ID " + _
        "And b.ID=d.诊疗项目ID And d.应用场合=Decode(a.主页ID,Null,1,2) Order By nvl(a.相关ID,0)"
    If blnMoved Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
    End If
    Set rsTmp = OpenSQLRecord(strSQL, Me.Name, lng医嘱ID)
    If rsTmp.EOF Then Exit Sub
    prbRefresh.Value = 5
    
    strDiagName = rsTmp(5): strDrAdvice = rsTmp(6)
    
    '构造附加项目串
    rsTmp.MoveNext: strExtData = ""
    Do While Not rsTmp.EOF
        strExtData = strExtData & "," & rsTmp(4)
    
        rsTmp.MoveNext
    Loop
    If Len(strExtData) > 0 Then strExtData = Mid(strExtData, 2)
    rsTmp.MoveFirst
    
    intType = -1
    Me.txt医嘱内容 = strDiagName
    If rsTmp!类别 = "D" And zlCommFun.NVL(GetItemField(rsTmp(4), "组合项目"), 0) = 1 Then
        '检查组合项目
        intType = 0
        Call AdviceSet检查手术(1, strExtData)
        txt医嘱内容.Text = Get检查手术名称(1, strDiagName)
        Me.txt附加 = Get部位名称
    ElseIf rsTmp!类别 = "F" Then
        '手术：需要输入麻醉项目，及可选择附加手术
        intType = 1
        Call AdviceSet检查手术(2, strExtData)
        txt医嘱内容.Text = Get检查手术名称(2, strDiagName)
        Me.txt附加 = Get麻醉名称
    ElseIf InStr(",7,8,", rsTmp!类别) > 0 Then
        '中药配方(单味草药当配方处理)
        intType = 2
    ElseIf rsTmp!类别 = "C" Then
        '检验项目选择检验标本
        intType = 3
        Me.txt附加 = rsTmp("标本部位")
    End If
    
    alngFileID(0) = rsTmp("申请ID"): PatientID = rsTmp(0): CheckID = IIf(rsTmp(3) = 0, rsTmp(2), rsTmp(1))
    PatientType = rsTmp(3): FileTypeID = rsTmp("病历文件ID"): bSample = False
    
    '显示医嘱内容
    If IsNull(rsTmp("开始执行时间")) Then
        Me.chk开始时间.Visible = True: Me.lbl开始时间.Visible = False: Me.chk开始时间.Value = 0
        Me.txt开始时间 = CDate(Date & " " & Time): Me.txt开始时间.Enabled = False
    Else
        Me.txt开始时间 = rsTmp("开始执行时间"): Me.txt开始时间.Enabled = True
    End If
    Me.chk紧急.Value = rsTmp("紧急标志")
    If Not IsNull(rsTmp("医生嘱托")) Then Me.txt医生嘱托 = rsTmp("医生嘱托")
    Me.txt频率 = rsTmp("执行频次"): Me.txt频率.Enabled = True: Me.cmd频率.Enabled = True
    Me.lbl总量单位.Caption = Trim(rsTmp("计算单位"))
    If Not IsNull(rsTmp("总给予量")) Then Me.txt总量 = rsTmp("总给予量"): Me.txt总量.Enabled = True
    If Not IsNull(rsTmp("单次用量")) Then Me.txt单量 = rsTmp("单次用量"): Me.txt单量.Enabled = True: Me.txt单量.BackColor = Me.txt医嘱内容.BackColor: Me.lbl单量单位.Caption = Trim(rsTmp("计算单位"))
    Me.cbo执行科室.Clear: Me.cbo执行科室.AddItem rsTmp("科室编码") & "-" & rsTmp("科室名称")
    Me.cbo执行科室.Text = rsTmp("科室编码") & "-" & rsTmp("科室名称"): Me.cbo执行科室.Enabled = True
    Me.cbo医生.Clear: Me.cbo医生.AddItem rsTmp("开嘱医生")
    Me.cbo医生.Text = rsTmp("开嘱医生"): Me.cbo医生.Enabled = True
    Me.picAdvice.Enabled = False
    
    SetItemFormat
    prbRefresh.Value = 15
    '初始化结束
    
    '判断能否编辑申请
    bAllowEdit = False
    iCurrElementIndex = 1
    
    Me.MousePointer = vbHourglass
'    ProFile1(0).ShowFile IIf(alngFileID(0) = 0, "", CStr(alngFileID(0))), PatientID, CheckID, PatientType, FileTypeID, bSample, 1, prbRefresh, , , , blnMoved
'    ProFile1(0).SetActiveElement 1
    Me.MousePointer = vbDefault
End Sub

Public Sub ShowMe_Report(ByVal strNO As String, ByVal int记录性质 As Integer, Optional objPrbRefresh As Object, Optional DataMoved As Boolean = False)
    Dim rsTmp As New ADODB.Recordset, i As Integer
    Dim strDiagName As String '诊疗项目名称
    Dim strDrAdvice As String '医生嘱托
    Dim bAllowEdit As Boolean
    Dim strSQL As String

    If sCheckNo = strNO And iRecordType = int记录性质 Then Exit Sub
    
    sCheckNo = strNO: iRecordType = int记录性质
    blnMoved = DataMoved
    On Error Resume Next
    '初始化
    Set prbRefresh = objPrbRefresh
    ClearForm

    strSQL = "Select a.病人ID,a.主页ID,a.挂号单,Decode(a.主页ID,Null,0,1),b.ID,b.名称,a.医生嘱托,a.ID,a.申请ID," + _
        "医嘱内容,开始执行时间,紧急标志,执行频次,总给予量,单次用量,d.编码 As 科室编码,d.名称 As 科室名称,开嘱医生,b.类别,nvl(a.标本部位,' ') As 标本部位,nvl(c.报告ID,0) As 报告ID,e.病历文件ID  " + _
        "From 病人医嘱记录 a,诊疗项目目录 b,病人医嘱发送 c,部门表 d,诊疗单据应用 e Where" & _
        " c.NO=[1] and c.记录性质=[2] And a.诊疗项目ID=b.ID(+) And a.ID=c.医嘱ID And a.执行科室ID=d.ID(+) " + _
        "And b.ID=e.诊疗项目ID And e.应用场合=Decode(a.主页ID,Null,1,2) Order By nvl(a.相关ID,0)"
    If blnMoved Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
        strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
    End If
    Set rsTmp = OpenSQLRecord(strSQL, Me.Name, strNO, int记录性质)
    If rsTmp.EOF Then Exit Sub
    prbRefresh.Value = 5
    
    strDiagName = rsTmp(5): strDrAdvice = rsTmp(6)
    
     '构造附加项目串
    rsTmp.MoveNext: strExtData = ""
    Do While Not rsTmp.EOF
        strExtData = strExtData & "," & rsTmp(4)

        rsTmp.MoveNext
    Loop
    If Len(strExtData) > 0 Then strExtData = Mid(strExtData, 2)
    rsTmp.MoveFirst

    intType = -1
    Me.txt医嘱内容 = strDiagName
    If rsTmp!类别 = "D" And zlCommFun.NVL(GetItemField(rsTmp(4), "组合项目"), 0) = 1 Then
        '检查组合项目
        intType = 0
        Call AdviceSet检查手术(1, strExtData)
        txt医嘱内容.Text = Get检查手术名称(1, strDiagName)
        Me.txt附加 = Get部位名称
    ElseIf rsTmp!类别 = "F" Then
        '手术：需要输入麻醉项目，及可选择附加手术
        intType = 1
        Call AdviceSet检查手术(2, strExtData)
        txt医嘱内容.Text = Get检查手术名称(2, strDiagName)
        Me.txt附加 = Get麻醉名称
    ElseIf InStr(",7,8,", rsTmp!类别) > 0 Then
        '中药配方(单味草药当配方处理)
        intType = 2
    ElseIf rsTmp!类别 = "C" Then
        '检验项目选择检验标本
        intType = 3
        Me.txt附加 = rsTmp("标本部位")
    End If

    alngFileID(0) = rsTmp("报告ID"): PatientID = rsTmp(0): CheckID = IIf(rsTmp(3) = 0, rsTmp(2), rsTmp(1))
    PatientType = rsTmp(3): FileTypeID = 0: bSample = False: AdviceID = rsTmp(7)

    '显示医嘱内容
    If IsNull(rsTmp("开始执行时间")) Then
        Me.chk开始时间.Visible = True: Me.lbl开始时间.Visible = False: Me.chk开始时间.Value = 0
        Me.txt开始时间 = CDate(Date & " " & Time): Me.txt开始时间.Enabled = False
    Else
        Me.txt开始时间 = rsTmp("开始执行时间"): Me.txt开始时间.Enabled = True
    End If
    Me.chk紧急.Value = rsTmp("紧急标志")
    If Not IsNull(rsTmp("医生嘱托")) Then Me.txt医生嘱托 = rsTmp("医生嘱托")
    Me.txt频率 = rsTmp("执行频次"): Me.txt频率.Enabled = True: Me.cmd频率.Enabled = True
    Me.lbl总量单位.Caption = Trim(rsTmp("计算单位"))
    If Not IsNull(rsTmp("总给予量")) Then Me.txt总量 = rsTmp("总给予量"): Me.txt总量.Enabled = True
    If Not IsNull(rsTmp("单次用量")) Then Me.txt单量 = rsTmp("单次用量"): Me.txt单量.Enabled = True: Me.txt单量.BackColor = Me.txt医嘱内容.BackColor: Me.lbl单量单位.Caption = Trim(rsTmp("计算单位"))
    Me.cbo执行科室.Clear: Me.cbo执行科室.AddItem rsTmp("科室编码") & "-" & rsTmp("科室名称")
    Me.cbo执行科室.Text = rsTmp("科室编码") & "-" & rsTmp("科室名称"): Me.cbo执行科室.Enabled = True
    Me.cbo医生.Clear: Me.cbo医生.AddItem rsTmp("开嘱医生")
    Me.cbo医生.Text = rsTmp("开嘱医生"): Me.cbo医生.Enabled = True
    Me.picAdvice.Enabled = False

    SetItemFormat
    prbRefresh.Value = 15
    '初始化结束

    '判断能否编辑申请
    bAllowEdit = False
    iCurrElementIndex = 1
    
    Me.MousePointer = vbHourglass
    ProFile1(0).ShowFile IIf(alngFileID(0) = 0, "", CStr(alngFileID(0))), PatientID, CheckID, PatientType, FileTypeID, bSample, 2, prbRefresh, , , , blnMoved
    ProFile1(0).SetActiveElement 1
    Me.MousePointer = vbDefault
End Sub

Private Sub ClearForm()
    On Error Resume Next
    Me.txt单量 = "": Me.txt附加 = "": Me.txt开始时间 = "": Me.txt频率 = "": Me.txt医生嘱托 = ""
    Me.txt总量 = "": Me.txt医嘱内容 = "": Me.chk紧急 = 0: Me.chk开始时间 = 0: Me.cbo医生.ListIndex = -1: Me.cbo执行科室.ListIndex = -1
    
    Me.MousePointer = vbHourglass
    ProFile1(0).ShowFile "", "", "", 10, "0", False
    Me.MousePointer = vbDefault
End Sub

Private Sub SetItemFormat()   '根据申请项目决定显示方式
    Select Case intType
        Case 0
            Me.lbl医嘱内容.Caption = "检查项目": Me.lbl附加.Caption = "检查部位": Me.cmdExt.ToolTipText = "选择检查部位"
            Me.lbl附加.Visible = True: Me.txt附加.Visible = True: Me.cmdExt.Visible = True
        Case 1
            Me.lbl医嘱内容.Caption = "手术项目": Me.lbl附加.Caption = "麻醉方式": Me.cmdExt.ToolTipText = "选择麻醉方式"
            Me.lbl附加.Visible = True: Me.txt附加.Visible = True: Me.cmdExt.Visible = True
        Case 3
            Me.lbl医嘱内容.Caption = "检验项目": Me.lbl附加.Caption = "检验标本": Me.cmdExt.ToolTipText = "选择检验标本"
            Me.lbl附加.Visible = True: Me.txt附加.Visible = True: Me.cmdExt.Visible = True
        Case Else
            Me.lbl附加.Visible = False: Me.txt附加.Visible = False: Me.cmdExt.Visible = False
    End Select
End Sub

Private Sub Form_Load()
    ProFile1(0).ifShowDiagItem = False
End Sub

Private Sub Form_Resize()
    Dim lngTools As Single, lngStatus As Single
    Dim lngTxtWidth As Single
    Dim lngDistance As Single
    
    If WindowState = 1 Then Exit Sub
    lngDistance = 300
    
    On Error Resume Next
    With picDoc
        .Left = 0: .Top = 0
        .Width = Me.ScaleWidth: .Height = Me.ScaleHeight - .Top
    End With
    With picAdvice
        .Left = 0: .Top = 0
        .Width = picDoc.ScaleWidth
    End With
    With lineSplit
        .X2 = picAdvice.Width + .X1
    End With
    With Me.chk紧急
        .Left = picAdvice.Width - Me.lbl开始时间.Left - .Width
        If .Left < Me.txt开始时间.Left + Me.txt开始时间.Width + lngDistance Then .Left = Me.txt开始时间.Left + Me.txt开始时间.Width + lngDistance
    End With
    
    lngTxtWidth = (picAdvice.ScaleWidth - Me.lbl开始时间.Left - Me.cmdSel.Width - Me.txt医嘱内容.Left - lngDistance - _
        Me.lbl附加.Width - Me.cmdExt.Width - 60) / 2
    With Me.txt医嘱内容
        .Width = lngTxtWidth
        Me.cmdSel.Left = .Left + .Width
        Me.lbl附加.Left = Me.cmdSel.Left + Me.cmdSel.Width + lngDistance
    End With
    With Me.txt附加
        .Left = Me.lbl附加.Left + Me.lbl附加.Width + 30
        .Width = lngTxtWidth
        Me.cmdExt.Left = .Left + .Width
    End With
    Me.lineTitleSplit.X2 = Me.cmdExt.Left + Me.cmdExt.Width + 200

    With Me.txt医生嘱托
        .Width = picAdvice.Width - Me.lbl开始时间.Left - .Left
    End With
    
    lngTxtWidth = (picAdvice.Width - Me.lbl开始时间.Left - Me.txt频率.Left - Me.txt频率.Width - _
        (Me.lbl总量单位.Width + Me.lbl总量.Width + lngDistance + 2 * 30) - _
        (Me.lbl单量单位.Width + Me.lbl单量.Width + lngDistance + 2 * 30)) / 2
    If lngTxtWidth < 1000 Then lngTxtWidth = 1000
    Me.lbl总量.Left = Me.txt频率.Left + Me.txt频率.Width + lngDistance
    With Me.txt总量
        .Left = Me.lbl总量.Left + Me.lbl总量.Width + 30
        .Width = lngTxtWidth
    End With
    Me.lbl总量单位.Left = Me.txt总量.Left + Me.txt总量.Width + 30
    Me.lbl单量.Left = Me.lbl总量单位.Left + Me.lbl总量单位.Width + lngDistance
    With Me.txt单量
        .Left = Me.lbl单量.Left + Me.lbl单量.Width + 30
        .Width = lngTxtWidth
    End With
    Me.lbl单量单位.Left = Me.txt单量.Left + Me.txt单量.Width + 30
    
    With Me.cbo医生
        .Left = Me.txt单量.Left
        .Width = picAdvice.Width - Me.lbl开始时间.Left - .Left
    End With
    Me.lbl开嘱医生.Left = Me.cbo医生.Left - Me.lbl开嘱医生.Width
    
    With picFile
        .Left = 0: .Top = picAdvice.Top + picAdvice.Height
        .Width = picDoc.ScaleWidth
        .Height = picDoc.ScaleHeight - .Top
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    zlCommFun.OpenIme False
End Sub

Private Sub picFile_Resize()
    On Error Resume Next
    With ProFile1(iTabIndex)
        .Left = 0: .Top = 0
        .Width = picFile.ScaleWidth
        .Height = picFile.ScaleHeight
        
        If .Width > picFile.ScaleWidth Then Me.Width = .Width
        If .Height > picFile.ScaleHeight Then Me.Height = .Height + picFile.Top
    End With
End Sub

Private Sub AdviceSet检查手术(ByVal int类型 As Integer, ByVal strDataIDs As String)
'功能：1.重新设置指定检查组合项目的部位行,用于新输入检查组合项目或修改部位
'      2.重新设置指定手术项目的附加手术及麻醉项目行,用于新输入手术项目或手术项目的附加手术及麻醉项目
'参数：int类型=1=处理检查部位项目,2=处理附加手术及麻醉项目
'      strDataIDs=检查:包含检查部位信息,手术:包含附加手术及麻醉项目信息,其中可能没有附加手术和麻醉
    Dim strSQL As String, i As Long
    Dim arrIDs As Variant
    
    On Error GoTo errH
            
    '重新加入部位行或附加手术行及麻醉项目行
    If int类型 = 2 Then
        strDataIDs = Trim(Replace(strDataIDs, ";", ","))
        If Left(strDataIDs, 1) = "," Then strDataIDs = Mid(strDataIDs, 2)
        If Right(strDataIDs, 1) = "," Then strDataIDs = Mid(strDataIDs, 1, Len(strDataIDs) - 1)
    End If
    
    If strDataIDs <> "" Then
        If Not rsRelativeAdvice Is Nothing Then
            rsRelativeAdvice.Close
        Else
            Set rsRelativeAdvice = New ADODB.Recordset
        End If
        strSQL = "Select ID,编码,名称,nvl(标本部位,' ') As 标本部位," + _
        "类别,nvl(计价性质,0) As 计价性质,nvl(执行科室,0) As 执行科室 From 诊疗项目目录 Where ID IN(" & strDataIDs & ")"
        OpenRecord rsRelativeAdvice, strSQL, Me.Caption
    Else
        If Not rsRelativeAdvice Is Nothing Then rsRelativeAdvice.Close: Set rsRelativeAdvice = Nothing
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
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
            If rsRelativeAdvice("类别") <> "G" Then
                strTmp = strTmp & "," & rsRelativeAdvice("名称")
            End If
        End If
        
        rsRelativeAdvice.MoveNext
    Loop
    
    If strTmp <> "" Then
        Get检查手术名称 = txtMainAdvice & " 及 " & Mid(strTmp, 2)
    Else
        Get检查手术名称 = txtMainAdvice
    End If
End Function

Private Function Get麻醉名称() As String
    If rsRelativeAdvice Is Nothing Then Get麻醉名称 = "": Exit Function
    rsRelativeAdvice.MoveFirst
    Do While Not rsRelativeAdvice.EOF
        If Len(Trim(rsRelativeAdvice("名称"))) > 0 Then
            If rsRelativeAdvice("类别") = "G" Then
                Get麻醉名称 = rsRelativeAdvice("名称")
            End If
        End If
        
        rsRelativeAdvice.MoveNext
    Loop
End Function

Private Function Get部位名称() As String
    If rsRelativeAdvice Is Nothing Then Get部位名称 = "": Exit Function
        
    rsRelativeAdvice.MoveFirst
    Do While Not rsRelativeAdvice.EOF
        If Len(Trim(rsRelativeAdvice("标本部位"))) > 0 Then
            Get部位名称 = Get部位名称 & "," & rsRelativeAdvice("标本部位")
        End If
        
        rsRelativeAdvice.MoveNext
    Loop
    If Len(Get部位名称) > 0 Then Get部位名称 = Mid(Get部位名称, 2)
End Function

Private Function GetItemField(ByVal lng项目ID As Long, ByVal strField As String) As Variant
'功能：获取指定诊疗项目的指定字段信息
'说明：未处理NULL值
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select " & strField & " From 诊疗项目目录 Where ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng项目ID)
    If Not rsTmp.EOF Then GetItemField = rsTmp.Fields(strField).Value
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
