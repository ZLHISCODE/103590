VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLabFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "过滤"
   ClientHeight    =   6105
   ClientLeft      =   2475
   ClientTop       =   2355
   ClientWidth     =   6780
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLabFilter.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5340
      TabIndex        =   20
      Top             =   5610
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3750
      TabIndex        =   19
      Top             =   5610
      Width           =   1100
   End
   Begin TabDlg.SSTab sTab 
      Height          =   5415
      Left            =   60
      TabIndex        =   31
      Top             =   60
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   9551
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "常用"
      TabPicture(0)   =   "frmLabFilter.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraPatient"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "高级"
      TabPicture(1)   =   "frmLabFilter.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TxtVariable2"
      Tab(1).Control(1)=   "cboIF2"
      Tab(1).Control(2)=   "chkAdvanced"
      Tab(1).Control(3)=   "cmdUpdate"
      Tab(1).Control(4)=   "cmdAdd"
      Tab(1).Control(5)=   "cmdDel"
      Tab(1).Control(6)=   "TxtVariable1"
      Tab(1).Control(7)=   "cboIF1"
      Tab(1).Control(8)=   "txtVerifyItem"
      Tab(1).Control(9)=   "lvwComPages"
      Tab(1).Control(10)=   "lblPatient(14)"
      Tab(1).ControlCount=   11
      Begin VB.TextBox TxtVariable2 
         Height          =   285
         Left            =   -70470
         TabIndex        =   27
         Top             =   4140
         Width           =   765
      End
      Begin VB.ComboBox cboIF2 
         Height          =   315
         Left            =   -71100
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   4140
         Width           =   615
      End
      Begin VB.CheckBox chkAdvanced 
         Caption         =   "使用组合过滤"
         Height          =   195
         Left            =   -74880
         TabIndex        =   21
         Top             =   390
         Width           =   2205
      End
      Begin VB.CommandButton cmdUpdate 
         Height          =   285
         Left            =   -68820
         Picture         =   "frmLabFilter.frx":0044
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   4140
         Width           =   375
      End
      Begin VB.CommandButton cmdAdd 
         Height          =   285
         Left            =   -69660
         Picture         =   "frmLabFilter.frx":018E
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   4140
         Width           =   375
      End
      Begin VB.CommandButton cmdDel 
         Height          =   285
         Left            =   -69240
         Picture         =   "frmLabFilter.frx":02D8
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   4140
         Width           =   375
      End
      Begin VB.TextBox TxtVariable1 
         Height          =   285
         Left            =   -71790
         TabIndex        =   25
         Top             =   4140
         Width           =   675
      End
      Begin VB.ComboBox cboIF1 
         Height          =   315
         Left            =   -72420
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   4140
         Width           =   615
      End
      Begin VB.TextBox txtVerifyItem 
         Height          =   285
         Left            =   -74070
         TabIndex        =   23
         Top             =   4140
         Width           =   1635
      End
      Begin VB.Frame fraPatient 
         Height          =   4935
         Left            =   150
         TabIndex        =   32
         Top             =   330
         Width           =   6345
         Begin VB.ComboBox cboAntiResult 
            Height          =   315
            Left            =   4080
            Style           =   2  'Dropdown List
            TabIndex        =   56
            Top             =   3210
            Width           =   1995
         End
         Begin VB.ComboBox cboAnti 
            Height          =   315
            Left            =   1110
            TabIndex        =   54
            Text            =   "cboSelectItem"
            Top             =   3210
            Width           =   1995
         End
         Begin VB.ComboBox cboMicrobe 
            Height          =   315
            Left            =   1110
            TabIndex        =   52
            Text            =   "cboSelectItem"
            Top             =   2790
            Width           =   4965
         End
         Begin VB.ComboBox cboAgeUnit 
            Height          =   315
            ItemData        =   "frmLabFilter.frx":0422
            Left            =   5310
            List            =   "frmLabFilter.frx":0424
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   240
            Width           =   765
         End
         Begin VB.TextBox TxtNO 
            Height          =   285
            Left            =   1110
            TabIndex        =   7
            Top             =   1020
            Width           =   4965
         End
         Begin VB.ComboBox cboSelectItem 
            Height          =   315
            Left            =   1110
            TabIndex        =   11
            Text            =   "cboSelectItem"
            Top             =   2370
            Width           =   3525
         End
         Begin VB.ComboBox cboMachine 
            Height          =   315
            Left            =   1110
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1530
            Width           =   4965
         End
         Begin VB.OptionButton optOneItem 
            Caption         =   "单项"
            Height          =   255
            Left            =   5400
            TabIndex        =   13
            Top             =   2415
            Width           =   675
         End
         Begin VB.OptionButton optUnionItem 
            Caption         =   "组合"
            Height          =   225
            Left            =   4680
            TabIndex        =   12
            Top             =   2415
            Value           =   -1  'True
            Width           =   675
         End
         Begin VB.ComboBox cboSex 
            Height          =   315
            ItemData        =   "frmLabFilter.frx":0426
            Left            =   2460
            List            =   "frmLabFilter.frx":0428
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   240
            Width           =   645
         End
         Begin VB.TextBox txtAgeEnd 
            Height          =   285
            Left            =   4800
            TabIndex        =   3
            Top             =   270
            Width           =   465
         End
         Begin VB.TextBox txtAgeBegin 
            Height          =   285
            Left            =   4080
            TabIndex        =   2
            Top             =   270
            Width           =   465
         End
         Begin VB.TextBox txtPatientName 
            Height          =   285
            Left            =   1110
            TabIndex        =   0
            Top             =   255
            Width           =   855
         End
         Begin VB.TextBox txtSample 
            Height          =   285
            Left            =   4080
            TabIndex        =   6
            Top             =   645
            Width           =   1995
         End
         Begin VB.TextBox txtSampleID 
            Height          =   285
            Left            =   1110
            TabIndex        =   5
            Top             =   645
            Width           =   1995
         End
         Begin VB.ComboBox cboApplyDept 
            Height          =   315
            Left            =   1110
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   4470
            Width           =   1995
         End
         Begin VB.ComboBox cboApplyMan 
            Height          =   315
            Left            =   4080
            TabIndex        =   18
            Top             =   4470
            Width           =   1995
         End
         Begin VB.ComboBox cboVerifyman 
            Height          =   315
            Left            =   4080
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1950
            Width           =   1995
         End
         Begin VB.Frame Frame1 
            Height          =   30
            Left            =   60
            TabIndex        =   34
            Top             =   1410
            Width           =   6225
         End
         Begin VB.ComboBox cboVerifyType 
            Height          =   315
            Left            =   1110
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1950
            Width           =   1995
         End
         Begin VB.Frame Frame2 
            Height          =   30
            Left            =   60
            TabIndex        =   33
            Top             =   4320
            Width           =   6225
         End
         Begin VB.CheckBox chkDisableDate 
            Caption         =   "检验时间"
            Height          =   225
            Left            =   240
            TabIndex        =   14
            Top             =   3630
            Value           =   1  'Checked
            Width           =   1035
         End
         Begin MSComCtl2.DTPicker DTPVerifyBegin 
            Height          =   315
            Left            =   1110
            TabIndex        =   15
            Top             =   3900
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   104726531
            CurrentDate     =   39073
         End
         Begin MSComCtl2.DTPicker DTPVerifyEnd 
            Height          =   315
            Left            =   4080
            TabIndex        =   16
            Top             =   3900
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   104726531
            CurrentDate     =   39073
         End
         Begin VB.Label lblPatient 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "药敏结果:"
            Height          =   195
            Index           =   16
            Left            =   3240
            TabIndex        =   55
            Top             =   3270
            Width           =   780
         End
         Begin VB.Label lblPatient 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "抗  生  素:"
            Height          =   195
            Index           =   10
            Left            =   240
            TabIndex        =   53
            Top             =   3270
            Width           =   780
         End
         Begin VB.Label lblPatient 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "检验细菌:"
            Height          =   195
            Index           =   9
            Left            =   240
            TabIndex        =   51
            Top             =   2850
            Width           =   780
         End
         Begin VB.Label lblPatient 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "单  据  号:"
            Height          =   195
            Index           =   7
            Left            =   240
            TabIndex        =   50
            Top             =   1065
            Width           =   780
         End
         Begin VB.Label lblVerifyType 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "检验仪器:"
            Height          =   195
            Index           =   0
            Left            =   270
            TabIndex        =   49
            Top             =   1590
            Width           =   780
         End
         Begin VB.Label lblPatient 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "性别:"
            Height          =   195
            Index           =   11
            Left            =   2010
            TabIndex        =   47
            Top             =   300
            Width           =   420
         End
         Begin VB.Label lblPatient 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "--"
            Height          =   195
            Index           =   13
            Left            =   4620
            TabIndex        =   46
            Top             =   300
            Width           =   120
         End
         Begin VB.Label lblPatient 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "年        龄:"
            Height          =   195
            Index           =   12
            Left            =   3240
            TabIndex        =   45
            Top             =   300
            Width           =   780
         End
         Begin VB.Label lblPatient 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(忽略检验时间后查询时间会相对较长)"
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   15
            Left            =   1320
            TabIndex        =   44
            Top             =   3630
            Visible         =   0   'False
            Width           =   3000
         End
         Begin VB.Label lblPatient 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "标  本  号:"
            Height          =   195
            Index           =   1
            Left            =   3240
            TabIndex        =   43
            Top             =   675
            Width           =   780
         End
         Begin VB.Label lblPatient 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "标  识  号:"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   42
            Top             =   690
            Width           =   780
         End
         Begin VB.Label lblPatient 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "申请科室:"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   41
            Top             =   4530
            Width           =   780
         End
         Begin VB.Label lblPatient 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "申  请  人:"
            Height          =   195
            Index           =   4
            Left            =   3240
            TabIndex        =   40
            Top             =   4530
            Width           =   780
         End
         Begin VB.Label lblPatient 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ">>>>>>>>"
            Height          =   195
            Index           =   6
            Left            =   3120
            TabIndex        =   39
            Top             =   3960
            Width           =   960
         End
         Begin VB.Label lblPatient 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "检验项目:"
            Height          =   195
            Index           =   5
            Left            =   240
            TabIndex        =   38
            Top             =   2415
            Width           =   780
         End
         Begin VB.Label lblVerifyType 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "检验类别:"
            Height          =   195
            Index           =   5
            Left            =   240
            TabIndex        =   37
            Top             =   2010
            Width           =   780
         End
         Begin VB.Label lblPatient 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "检  验  人:"
            Height          =   195
            Index           =   8
            Left            =   3240
            TabIndex        =   36
            Top             =   2010
            Width           =   780
         End
         Begin VB.Label lblPatient 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "病人姓名:"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   35
            Top             =   300
            Width           =   780
         End
      End
      Begin MSComctlLib.ListView lvwComPages 
         Height          =   3465
         Left            =   -74910
         TabIndex        =   22
         Top             =   630
         Width           =   6465
         _ExtentX        =   11404
         _ExtentY        =   6112
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label lblPatient 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "检验项目:"
         Height          =   195
         Index           =   14
         Left            =   -74880
         TabIndex        =   48
         Top             =   4185
         Width           =   780
      End
   End
End
Attribute VB_Name = "frmLabFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngDept As Long                                        '检验科室
Private mlngMachine As Long                                     '检验仪器
Private mstrMachines As String                                  '可以操作的仪器ID
Private mstrCondition As String                                 '返回条件字串

Private Sub InitControl()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim intLoop As Integer                                  '循环变量
    Dim strTmp As String                                    '临时字串变量
    Dim varAFilter As Variant                               '组合条件
    Dim varAItem As Variant                                 '组合条件项目
    Dim Item As ListItem                                    '列表项目对象
    
    On Error GoTo errH
    
    '性别
    With Me.cboSex
        .AddItem ""
        .AddItem "男"
        .AddItem "女"
        .ListIndex = 0
    End With
    
    '年龄单位
    With Me.cboAgeUnit
        .AddItem ""
        .AddItem "岁"
        .AddItem "月"
        .AddItem "天"
        .AddItem "小时"
        .AddItem "成人"
        .AddItem "婴儿"
    End With
    
    '检验仪器
    With Me.cboMachine
        .AddItem ""
        strSQL = "select id , 编码, 名称 from 检验仪器 a where " & _
                 " A.ID In(Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))) " & _
                 " order by 编码"
        Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, mstrMachines)

        Do Until rsTmp.EOF
            .AddItem Nvl(rsTmp("编码")) & "-" & Nvl(rsTmp("名称"))
            .ItemData(.NewIndex) = Nvl(rsTmp("id"))
            If mlngMachine = Nvl(rsTmp("ID")) Then .ListIndex = .NewIndex
            rsTmp.MoveNext
        Loop
        If .ListIndex = -1 Then
            .ListIndex = 0
        End If
'        .ListIndex = 0
    End With
    
    '检验类别
'    With Me.cboVerifyType
'        .AddItem ""
'        strSQL = "select 编码,名称 from 诊疗检验类型 order by 编码"
'        zldatabase.OpenRecordset rsTmp, strSQL, gstrSysName
'        Do Until rsTmp.EOF
'            .AddItem Nvl(rsTmp("编码")) & "-" & Nvl(rsTmp("名称"))
'            .ItemData(.NewIndex) = Nvl(rsTmp("编码"))
'            rsTmp.MoveNext
'        Loop
'        .ListIndex = 0
'    End With
    
    '科室人员
    With Me.cboVerifyman
        .AddItem ""
        strSQL = "select a.id,a.编号,a.姓名 from 人员表 a , 部门人员 b " & _
                 " Where a.ID = b.人员id And b.部门ID = [1]  " & _
                 " And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) " & _
                 " order by a.编号"
        Set rsTmp = zldatabase.OpenSQLRecord(strSQL, gstrSysName, mlngDept)
        Do Until rsTmp.EOF
            .AddItem Nvl(rsTmp("编号")) & "-" & Nvl(rsTmp("姓名"))
            .ItemData(.NewIndex) = Nvl(rsTmp("ID"))
            rsTmp.MoveNext
        Loop
        .ListIndex = 0
    End With
        
    '送检科室
    With Me.cboApplyDept
        .AddItem ""
        strSQL = "select distinct a.id,a.编码,a.名称 from 部门表 a , 部门性质说明 b " & _
                 " where a.id = b.部门id and b.工作性质 in ('检验','护理','临床','体检')  order by a.编码"
        zldatabase.OpenRecordset rsTmp, strSQL, gstrSysName
        Do Until rsTmp.EOF
            .AddItem Nvl(rsTmp("编码")) & "-" & Nvl(rsTmp("名称"))
            .ItemData(.NewIndex) = Nvl(rsTmp("id"))
            rsTmp.MoveNext
        Loop
        .ListIndex = 0
    End With
    
    '初始化时间
    Me.DTPVerifyBegin.Value = Format(zldatabase.Currentdate - 1, "yyyy-MM-dd 00:00:00")
    Me.DTPVerifyEnd.Value = Format(zldatabase.Currentdate, "yyyy-MM-dd 23:59:59")
     
    '组合条件判断
    With Me.cboIF1
        .AddItem "="
        .AddItem ">"
        .AddItem ">="
        .AddItem "<"
        .AddItem "<="
        .AddItem "<>"
        .ListIndex = 0
    End With
    
    With Me.cboIF2
        .AddItem ""
        .AddItem "="
        .AddItem ">"
        .AddItem ">="
        .AddItem "<"
        .AddItem "<="
        .AddItem "<>"
        .ListIndex = 0
    End With
    
    With Me.cboAntiResult
        .AddItem ""
        .AddItem "R-耐药": .ItemData(.NewIndex) = "1"
        .AddItem "I-中介": .ItemData(.NewIndex) = "2"
        .AddItem "S-敏感": .ItemData(.NewIndex) = "3"
        .ListIndex = 0
    End With
    
    '列表头初始化
    With Me.lvwComPages
        .ColumnHeaders.Add , "A", "检验项目", 2000
        .ColumnHeaders.Add , "B", "计算1", 900
        .ColumnHeaders.Add , "C", "值1", 1200
        .ColumnHeaders.Add , "D", "计算2", 900
        .ColumnHeaders.Add , "E", "值2", 1200
        .View = lvwReport
    End With
    
    Me.sTab.Tab = 0
    
    '读入注册表中的设置
    chkAdvanced.Value = zldatabase.GetPara("使用组合查询", 100, 1208, 1)
    chkDisableDate.Value = zldatabase.GetPara("是否使用时间", 100, 1208, 1)
    strTmp = zldatabase.GetPara("组合查询", 100, 1208, "")
    
    If strTmp <> "" And Trim(strTmp) <> "0" Then
        varAFilter = Split(strTmp, ",")
        For intLoop = 0 To UBound(varAFilter)
            varAItem = Split(varAFilter(intLoop), "^")
            With Me.lvwComPages.ListItems
                Set Item = .Add(, "A" & varAItem(0), varAItem(1))
                Item.SubItems(1) = varAItem(2)
                Item.SubItems(2) = varAItem(3)
                Item.SubItems(3) = varAItem(4)
                Item.SubItems(4) = varAItem(5)
            End With
        Next
    End If
    '每次进来的时候都不使用高级过滤功能
    chkAdvanced.Value = 0
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub ShowMe(objfrm As Object, lngVerifyDept As Long, lngVerifyMachine As Long, strMachines As String, ByRef strCondition As String)
    '功能                           '过滤查询
    '参数                           objfrm 主窗体对象
    '                               lngVerifyDept 检验科室  0=所有科室(小组下)
    '                               strMachines 可以查询的仪器ID字串
    '                               lngVerifyMachine 检验仪器 0=所有仪器(小组下) -1=手工
    '返回                           strCondition 返回查询字串
    mlngDept = lngVerifyDept
    mlngMachine = lngVerifyMachine
    mstrMachines = strMachines
    Me.Show vbModal, objfrm
    strCondition = mstrCondition
    mstrCondition = ""
End Sub

Private Sub cboAnti_Click()
    If Me.cboAnti.ListIndex > -1 Then
        Me.cboAnti.Tag = Me.cboAnti.ItemData(Me.cboAnti.ListIndex)
    Else
        Me.cboAnti.Tag = ""
    End If
End Sub

Private Sub cboAnti_KeyPress(KeyAscii As Integer)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim strVerifyType As String                             '检验类别
    Dim mint简码 As Integer                                 '按那种方式的简码查找
    Dim vRect As RECT                                       '选择框定位
    
    On Error GoTo errH
    
    If KeyAscii = 13 Then
        '==抗生素
        If Val(Me.cboMicrobe.Tag) <> 0 Then
            strSQL = "Select distinct A.ID, A.编码, A.中文名, A.英文名" & vbNewLine & _
                " From 检验用抗生素 A, 检验细菌抗生素 B, 检验抗生素用药 C" & vbNewLine & _
                " Where B.抗生素分组id = C.抗生素分组id And C.抗生素id = A.ID And B.细菌id = [1] and " & vbNewLine & _
                " (a.编码 like [2] or a.中文名 like [2] or a.英文名 like [2]) order by a.编码 "
        Else
            strSQL = "select ID,编码,中文名,英文名 from 检验用抗生素 a " & vbNewLine & _
                     " where (a.编码 like [2] or a.中文名 like [2] or a.英文名 like [2])  order by 编码 "
        End If
        
        vRect = GetControlRect(cboAnti.hWnd)
        Set rsTmp = zldatabase.ShowSQLSelect(Me, strSQL, 0, "检验项目", False, "", "", False, False, True, _
                    vRect.Left, vRect.Top, cboAnti.Height, False, False, True, CLng(Me.cboMicrobe.Tag), UCase(Me.cboAnti.Text & "%"))
                    
        If Not rsTmp Is Nothing Then
            If rsTmp.State <> 0 Then
                cboAnti.Tag = Nvl(rsTmp("ID"))
                cboAnti.Text = Nvl(rsTmp("编码")) & "-" & Nvl(rsTmp("中文名")) & _
                                    IIf(Nvl(rsTmp("英文名")) <> "", "(" & Nvl(rsTmp("英文名")) & ")", "")
                zlCommFun.PressKey vbKeyTab
            End If
        Else
            cboAnti.Tag = ""
            cboAnti = ""
        End If
        cboAnti.SelStart = 0
        cboAnti.SelLength = Len(cboAnti.Text)
        
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cboAntiResult_Click()
    If Me.cboAntiResult.ListIndex > -1 Then
        Me.cboAntiResult.Tag = Me.cboAntiResult.ItemData(Me.cboAntiResult.ListIndex)
        Select Case Me.cboAntiResult.Tag
            Case 1          '耐药
                Me.cboAntiResult.Tag = "R"
            Case 2          '中介
                Me.cboAntiResult.Tag = "I"
            Case 3          '敏感
                Me.cboAntiResult.Tag = "S"
            Case Else
                Me.cboAntiResult.Tag = ""
        End Select
    Else
        Me.cboAntiResult.Tag = ""
    End If
End Sub

Private Sub cboApplyDept_Click()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim lngApplyDept As Long                                '送检科室
    
    On Error GoTo errH
    
    lngApplyDept = Me.cboApplyDept.ItemData(Me.cboApplyDept.ListIndex)
    
    '读入对应科室下的人员
    strSQL = "select distinct a.id,a.编号,a.姓名 from 人员表 a , 人员性质说明 b , 部门人员 c " & _
                 " where a.id = b.人员id and a.id = c.人员ID and  b.人员性质 in ('医生','护士') " & _
                 " and (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) "
    If lngApplyDept = 0 Then
        strSQL = strSQL & " order by a.编号"
    Else
        strSQL = strSQL & " and c.部门id = [1] order by a.编号 "
    End If
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, gstrSysName, lngApplyDept)
            
    Me.cboApplyMan.Clear
                
    With Me.cboApplyMan
        .AddItem ""
        Do Until rsTmp.EOF
            .AddItem Nvl(rsTmp("编号")) & "-" & Nvl(rsTmp("姓名"))
            .ItemData(.NewIndex) = Nvl(rsTmp("ID"))
            rsTmp.MoveNext
        Loop
        .ListIndex = 0
    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub



Private Sub cboApplyDept_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cboApplyMan_Click()
    Me.cboApplyMan.Tag = Me.cboApplyMan.Text
End Sub

Private Sub cboApplyMan_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intIdx As Long, i As Long
    Dim strInput As String
    Dim vRect As RECT, blnCancel As Boolean
    Dim intLoop As Integer
    
    If KeyAscii = 13 Then
    
        If cboApplyMan.Text = "" Then '无输入
            Exit Sub
        End If
        
        If cboApplyMan.Text = cboApplyMan.Tag Then
            zlCommFun.PressKey vbKeyTab
        End If
        
        strInput = UCase(NeedName(cboApplyMan.Text))
        '全院医生
        strSQL = "Select Distinct 部门ID From 部门性质说明 Where 服务对象 IN(1,2,3) "
        If cboApplyDept.Text <> "" Then
            strSQL = strSQL & " And 部门ID = [3] "
        End If
        strSQL = "Select Distinct A.ID,A.编号,A.姓名,A.简码" & _
            " From 人员表 A,部门人员 B,人员性质说明 C" & _
            " Where A.ID=B.人员ID And A.ID=C.人员ID And C.人员性质='医生'" & _
            " And B.部门ID IN(" & strSQL & ")" & _
            " And (Upper(A.编号) Like [1] Or Upper(A.姓名) Like [2] Or Upper(A.简码) Like [2])" & _
            " And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) " & _
            " Order by A.简码"
    
        On Error GoTo errH
        vRect = GetControlRect(cboApplyMan.hWnd)
        Set rsTmp = zldatabase.ShowSQLSelect(Me, strSQL, 0, "申请人", False, "", "", False, False, _
            True, vRect.Left, vRect.Top, cboApplyMan.Height, blnCancel, False, True, strInput & "%", strInput & "%", _
            cboApplyDept.ItemData(cboApplyDept.ListIndex))
        If Not rsTmp Is Nothing Then
            For intLoop = 0 To Me.cboApplyMan.ListCount - 1
                If Me.cboApplyMan.ItemData(intLoop) = rsTmp("ID") Then
                    Me.cboApplyMan.ListIndex = intLoop
                    Me.cboApplyMan.Tag = Me.cboApplyMan.Text
                    Exit Sub
                End If
            Next
'            cboApplyMan.Text = rsTmp!姓名
    '        Me.dtp(0).SetFocus
    '        SetFocusNextIndex Me.cbo医生.TabIndex
    
    
        Else
            Me.cboApplyMan.Tag = ""
            If Not blnCancel Then
                MsgBox "未找到对应的申请人。", vbInformation, gstrSysName
            End If
            Exit Sub
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cboIF1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cboIF2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cboMachine_Click()
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim lngMachine As Long
    Dim strVerifyType As String
    
    If Me.cboMachine.Text <> "" Then
        strSQL = "select distinct D.编码 , d.名称 from  检验报告项目 a , 诊疗项目目录 b , 检验仪器项目 c , 诊疗检验类型 D " & _
                 " Where a.诊疗项目id = b.ID And a.报告项目id = c.项目id And  b.操作类型 = d.名称 and c.仪器id = [1] order by D.编码"
    Else
        strSQL = "select distinct D.编码 , d.名称 from   诊疗检验类型 D order by d.编码"
    End If
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, gstrSysName, Me.cboMachine.ItemData(Me.cboMachine.ListIndex))
    
    If Me.cboVerifyType.ListCount > 0 And Me.cboVerifyType.Text <> "" Then
        strVerifyType = Me.cboVerifyType.ItemData(Me.cboVerifyType.ListIndex)
    End If
    
    If Me.cboVerifyType.ListCount > 0 Then
        strVerifyType = Me.cboVerifyType.ItemData(Me.cboVerifyType.ListIndex)
    End If
    Me.cboVerifyType.Clear
    Me.cboVerifyType.AddItem ""
    Do Until rsTmp.EOF
        Me.cboVerifyType.AddItem rsTmp("编码") & "-" & rsTmp("名称")
        Me.cboVerifyType.ItemData(Me.cboVerifyType.NewIndex) = rsTmp("编码")
        If Val(strVerifyType) = Me.cboVerifyType.ItemData(Me.cboVerifyType.NewIndex) Then
            Me.cboVerifyType.ListIndex = Me.cboVerifyType.NewIndex
        End If
        rsTmp.MoveNext
    Loop
    
    If Me.cboVerifyType.ListCount > 0 And Me.cboVerifyType.Text = "" Then
        Me.cboVerifyType.ListIndex = 0
    End If
    
End Sub

Private Sub cboMicrobe_Click()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    If Me.cboMicrobe.ListIndex > -1 Then
        Me.cboMicrobe.Tag = Me.cboMicrobe.ItemData(Me.cboMicrobe.ListIndex)
    Else
        Me.cboMicrobe.Tag = ""
    End If
    
    On Error GoTo errH
    
    '==抗生素
    If Val(Me.cboMicrobe.Tag) <> 0 Then
        strSQL = "Select distinct A.ID, A.编码, A.中文名, A.英文名" & vbNewLine & _
            " From 检验用抗生素 A, 检验细菌抗生素 B, 检验抗生素用药 C" & vbNewLine & _
            " Where B.抗生素分组id = C.抗生素分组id And C.抗生素id = A.ID And B.细菌id = [1]"
    Else
        strSQL = "select ID,编码,中文名,英文名 from 检验用抗生素 order by 编码"
    End If
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, CLng(Me.cboMicrobe.Tag))
    
    With Me.cboAnti
        .Clear: .AddItem ""
        Do Until rsTmp.EOF
            .AddItem Nvl(rsTmp("编码")) & "-" & Nvl(rsTmp("中文名")) & _
            IIf(Nvl(rsTmp("英文名")) <> "", "(" & Nvl(rsTmp("英文名")) & ")", "")
            .ItemData(.NewIndex) = rsTmp("ID")
            rsTmp.MoveNext
        Loop
        If .ListCount > 0 And Trim(.Text) <> "" Then
            .ListIndex = 0
        End If
    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cboMicrobe_KeyPress(KeyAscii As Integer)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim strVerifyType As String                             '检验类别
    Dim mint简码 As Integer                                 '按那种方式的简码查找
    Dim vRect As RECT                                       '选择框定位
    
    On Error GoTo errH
    
    If KeyAscii = 13 Then
        strSQL = " select id,编码,中文名,英文名 from 检验细菌 " & vbNewLine & _
                 " where (编码 like [1] or 中文名 like [1] or 英文名 like [1] ) "
        
        vRect = GetControlRect(cboMicrobe.hWnd)
        Set rsTmp = zldatabase.ShowSQLSelect(Me, strSQL, 0, "检验项目", False, "", "", False, False, True, _
                    vRect.Left, vRect.Top, cboMicrobe.Height, False, False, True, UCase(Me.cboMicrobe.Text & "%"))
                    
        If Not rsTmp Is Nothing Then
            If rsTmp.State <> 0 Then
                cboMicrobe.Tag = Nvl(rsTmp("ID"))
                cboMicrobe.Text = Nvl(rsTmp("编码")) & "-" & Nvl(rsTmp("中文名")) & _
                                    IIf(Nvl(rsTmp("英文名")) <> "", "(" & Nvl(rsTmp("英文名")) & ")", "")
                zlCommFun.PressKey vbKeyTab
            End If
        Else
            cboMicrobe.Tag = ""
            cboMicrobe = ""
        End If
        cboMicrobe.SelStart = 0
        cboMicrobe.SelLength = Len(cboMicrobe.Text)
        
        '==抗生素
        If Me.cboMicrobe.Tag <> "" Then
            strSQL = "Select distinct A.ID, A.编码, A.中文名, A.英文名" & vbNewLine & _
                " From 检验用抗生素 A, 检验细菌抗生素 B, 检验抗生素用药 C" & vbNewLine & _
                " Where B.抗生素分组id = C.抗生素分组id And C.抗生素id = A.ID And B.细菌id = [1]"
        Else
            strSQL = "select ID,编码,中文名,英文名 from 检验用抗生素 order by 编码"
        End If
        Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, CLng(Val(Me.cboMicrobe.Tag)))
        
        With Me.cboAnti
            .Clear: .AddItem ""
            Do Until rsTmp.EOF
                .AddItem Nvl(rsTmp("编码")) & "-" & Nvl(rsTmp("中文名")) & _
                IIf(Nvl(rsTmp("英文名")) <> "", "(" & Nvl(rsTmp("英文名")) & ")", "")
                .ItemData(.NewIndex) = rsTmp("ID")
                rsTmp.MoveNext
            Loop
            If .ListCount > 0 And Trim(.Text) <> "" Then
                .ListIndex = 0
            End If
        End With
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cboSelectItem_Click()
    Me.cboSelectItem.Tag = Me.cboSelectItem.ItemData(Me.cboSelectItem.ListIndex)
End Sub

Private Sub cboSelectItem_GotFocus()
    Me.cboSelectItem.SelStart = 0
    Me.cboSelectItem.SelLength = Len(Me.cboSelectItem.Text)
End Sub

Private Sub cboSelectItem_KeyPress(KeyAscii As Integer)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim strVerifyType As String                             '检验类别
    Dim mint简码 As Integer                                 '按那种方式的简码查找
    Dim vRect As RECT                                       '选择框定位
    
    On Error GoTo errH
    
    mint简码 = zldatabase.GetPara("简码方式") '简码匹配方式：0-拼音,1-五笔
    
    If KeyAscii = 13 Then
        strSQL = " select distinct b.id ,b.编码,b.名称 " & _
                 " from 检验报告项目 a , 诊疗项目目录 b ,  诊治所见项目 c ,诊疗项目别名 d  " & _
                 " where a.诊疗项目id = b.id and  a.报告项目id = c.id and b.ID = d.诊疗项目id " & _
                 "       and d.码类 = [2] and (c.英文名 like [1] or b.名称 like [1] or d.简码 like [1] ) "
        
        '如果选择类别按类别查
        If Me.cboVerifyType.Text <> "" Then
            strSQL = strSQL & " and b.操作类型 = [3] "
            strVerifyType = Mid(Me.cboVerifyType.Text, InStr(1, Me.cboVerifyType.Text, "-") + 1)
        End If
        
        '按查找组合还是单一项目
        If Me.optUnionItem.Value = True Then
            strSQL = strSQL & " and b.组合项目 = 1 "
        Else
            strSQL = strSQL & " and b.组合项目 <> 1 "
        End If
        
        
        If Me.optUnionItem.Value = False Then
            strSQL = Replace(strSQL, "distinct b.id", "distinct c.id")
        End If
        
        
        
        vRect = GetControlRect(cboSelectItem.hWnd)
        Set rsTmp = zldatabase.ShowSQLSelect(Me, strSQL, 0, "检验项目", False, "", "", False, False, True, _
                    vRect.Left, vRect.Top, cboSelectItem.Height, False, False, True, UCase(cboSelectItem) & "%", mint简码 + 1, strVerifyType)
                    
        If Not rsTmp Is Nothing Then
            If rsTmp.State <> 0 Then
                cboSelectItem.Tag = Nvl(rsTmp("ID"))
                cboSelectItem.Text = Nvl(rsTmp("名称"))
                zlCommFun.PressKey vbKeyTab
            End If
        Else
            cboSelectItem.Tag = ""
            cboSelectItem = ""
        End If
        cboSelectItem.SelStart = 0
        cboSelectItem.SelLength = Len(cboSelectItem.Text)
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cboSex_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cboVerifyman_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cboVerifyType_Click()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim strSelectItem As String
    
    '==普通检验项目
    If optUnionItem.Value = True Then
        '组合
        If Me.cboVerifyType.Text <> "" Then
            strSQL = "select id ,编码,名称 from 诊疗项目目录 where 组合项目 = 1 and 操作类型 =[1] order by 编码"
        Else
            strSQL = "select id, 编码,名称 from 诊疗项目目录 where 组合项目 = 1 and 类别 = 'C' order by 编码 "
        End If
    Else
        '细项
        If Me.cboVerifyType.Text <> "" Then
            strSQL = "Select Distinct c.Id, b.编码, b.名称" & _
                    " From 检验报告项目 A, 诊疗项目目录 B, 诊治所见项目 C, 诊疗项目别名 D" & _
                    " Where a.诊疗项目id = b.ID And a.报告项目id = c.ID And b.ID = d.诊疗项目id And b.组合项目 <> 1 And b.操作类型 = [1]" & _
                    " Order By 编码"
        Else
            strSQL = "Select Distinct c.Id, b.编码, b.名称" & _
                    " From 检验报告项目 A, 诊疗项目目录 B, 诊治所见项目 C, 诊疗项目别名 D" & _
                    " Where a.诊疗项目id = b.Id And a.报告项目id = c.Id And b.Id = d.诊疗项目id And b.组合项目 <> 1 And b.类别 = 'C'" & _
                    " Order By 编码"
        End If
    End If
        
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, gstrSysName, Mid(Me.cboVerifyType.Text, InStr(Me.cboVerifyType.Text, "-") + 1))
                    
    If Me.cboSelectItem.ListCount > 0 And Me.cboSelectItem.ListIndex <> -1 Then
        strSelectItem = Me.cboSelectItem.ItemData(Me.cboSelectItem.ListIndex)
    End If
    Me.cboSelectItem.Clear
    Me.cboSelectItem.AddItem ""
    Do Until rsTmp.EOF
        Me.cboSelectItem.AddItem rsTmp("编码") & "-" & rsTmp("名称")
        Me.cboSelectItem.ItemData(Me.cboSelectItem.NewIndex) = rsTmp("id")
        If Val(strSelectItem) = Me.cboSelectItem.ItemData(Me.cboSelectItem.NewIndex) Then
            Me.cboSelectItem.ListIndex = Me.cboSelectItem.NewIndex
        End If
        rsTmp.MoveNext
    Loop
    
    If Me.cboSelectItem.ListCount > 0 And Me.cboSelectItem.Text = "" Then
        Me.cboSelectItem.ListIndex = 0
    End If
    
    '==细菌
    strSQL = "select ID,编码,中文名,英文名 from 检验细菌 order by 编码"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption)
    With Me.cboMicrobe
        .Clear: .AddItem ""
        Do Until rsTmp.EOF
            .AddItem Nvl(rsTmp("编码")) & "-" & Nvl(rsTmp("中文名")) & _
            IIf(Nvl(rsTmp("英文名")) <> "", "(" & Nvl(rsTmp("英文名")) & ")", "")
            .ItemData(.NewIndex) = rsTmp("ID")
            rsTmp.MoveNext
        Loop
        If .ListCount > 0 And Trim(.Text) = "" Then
            .ListIndex = 0
        End If
    End With
    
    '==抗生素
    If Val(Me.cboMicrobe.Tag) <> 0 Then
        strSQL = "Select distinct A.ID, A.编码, A.中文名, A.英文名" & vbNewLine & _
            " From 检验用抗生素 A, 检验细菌抗生素 B, 检验抗生素用药 C" & vbNewLine & _
            " Where B.抗生素分组id = C.抗生素分组id And C.抗生素id = A.ID And B.细菌id = [1]"
    Else
        strSQL = "select ID,编码,中文名,英文名 from 检验用抗生素 order by 编码"
    End If
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, CLng(Me.cboMicrobe.Tag))
    
    With Me.cboAnti
        .Clear: .AddItem ""
        Do Until rsTmp.EOF
            .AddItem Nvl(rsTmp("编码")) & "-" & Nvl(rsTmp("中文名")) & _
            IIf(Nvl(rsTmp("英文名")) <> "", "(" & Nvl(rsTmp("英文名")) & ")", "")
            .ItemData(.NewIndex) = rsTmp("ID")
            rsTmp.MoveNext
        Loop
        If .ListCount > 0 And Trim(.Text) <> "" Then
            .ListIndex = 0
        End If
    End With
End Sub

Private Sub cboVerifyType_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub chkAdvanced_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub chkDisableDate_Click()
    If Me.chkDisableDate.Value = 1 Then
        Me.DTPVerifyBegin.Enabled = True
        Me.DTPVerifyEnd.Enabled = True
    Else
        Me.DTPVerifyBegin.Enabled = False
        Me.DTPVerifyEnd.Enabled = False
    End If
End Sub

Private Sub cmdRefuse_Click()

End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command3_Click()

End Sub

Private Sub cmdAdd_Click()
    Dim Item As ListItem
    Dim intLoop As Integer
    
    '是否满足条件不满足条件时退出
    If txtVerifyItem.Tag = "" Then
        MsgBox "请输入检验项目!", vbInformation, gstrSysName
        txtVerifyItem.SetFocus
        Exit Sub
    End If
    
    If cboIF1.Text = "" Then
        MsgBox "请选择条件!", vbInformation, gstrSysName
        cboIF1.SetFocus
        Exit Sub
    End If
    
    If Trim(Me.TxtVariable1.Text) = "" Then
        MsgBox "请输入条件值!", vbInformation, gstrSysName
        TxtVariable1.SetFocus
        Exit Sub
    End If
    
    '加入列表
    With Me.lvwComPages
        
        For intLoop = 1 To .ListItems.Count
            If .ListItems(intLoop).Key = "A" & txtVerifyItem.Tag Then
                MsgBox "项目<" & txtVerifyItem.Text & ">已在列表中存在,你可以选择修改!", vbQuestion, gstrSysName
                .ListItems(intLoop).Selected = True
                Exit Sub
            End If
        Next
    
        Set Item = .ListItems.Add(, "A" & txtVerifyItem.Tag, txtVerifyItem.Text)
        Item.SubItems(1) = cboIF1.Text
        Item.SubItems(2) = TxtVariable1.Text
        Item.SubItems(3) = cboIF2.Text
        Item.SubItems(4) = TxtVariable2.Text
        
        Me.txtVerifyItem.Text = ""
        Me.txtVerifyItem.Tag = ""
        Me.TxtVariable1.Text = ""
        Me.TxtVariable2.Text = ""
        
        Me.txtVerifyItem.SetFocus
    End With
    
        
End Sub

Private Sub CmdCancel_Click()
    Dim intLoop As Integer                              '循环变量
    Dim strAdvancedWhere As String                      '高级条件
    
                    
    '高级条件写入
    With Me.lvwComPages
        For intLoop = 1 To .ListItems.Count
            strAdvancedWhere = strAdvancedWhere & "," & Mid(.ListItems(intLoop).Key, 2) & "^" & .ListItems(intLoop).Text & "^" & _
                                .ListItems(intLoop).SubItems(1) & "^" & .ListItems(intLoop).SubItems(2) & "^" & _
                                .ListItems(intLoop).SubItems(3) & "^" & .ListItems(intLoop).SubItems(4)
        Next
    End With
    
    '保存一些设置到注册表以方便以后调用
    zldatabase.SetPara "使用组合查询", Me.chkAdvanced.Value, 100, 1208
    zldatabase.SetPara "组合查询", Mid(strAdvancedWhere, 2), 100, 1208
    zldatabase.SetPara "是否使用时间", Me.chkDisableDate.Value, 100, 1208
    Unload Me
End Sub

Private Sub cmdDel_Click()
    With Me.lvwComPages
        If .ListItems.Count > 0 Then
            .ListItems.Remove (.SelectedItem.Index)
        End If
        .SetFocus
    End With
End Sub

Private Sub cmdOK_Click()
    Dim intLoop As Integer                              '循环变量
    Dim strAdvancedWhere As String                      '高级条件
    On Error Resume Next
    
    '常用条件写入
    mstrCondition = txtPatientName & ";" & cboSex & ";" & txtAgeBegin & "," & txtAgeEnd & ";" & cboAgeUnit.Text & ";" & _
                    txtSampleID.Text & ";" & txtSample.Text & ";" & TxtNO.Text & ";" & _
                    Mid(cboVerifyType, InStr(1, cboVerifyType, "-") + 1) & _
                    ";" & Mid(cboVerifyman, InStr(1, cboVerifyman, "-") + 1) & ";" & _
                    IIf(Me.cboSelectItem.Tag = 0, "", Me.cboSelectItem.Tag) & "," & Me.optUnionItem.Value & ";" & _
                    IIf(Me.chkDisableDate, Me.DTPVerifyBegin.Value & "," & Me.DTPVerifyEnd, ",") & ";" & _
                    cboApplyDept.ItemData(cboApplyDept.ListIndex) & ";" & Mid(cboApplyMan.Text, InStr(1, cboApplyMan.Text, "-") + 1) & ";" & _
                    Me.cboMachine.ItemData(cboMachine.ListIndex) & ";" & IIf(Me.cboMicrobe.Tag = 0, "", Me.cboMicrobe.Tag) & ";" & _
                    Me.cboAnti.Tag & ";" & Me.cboAntiResult.Tag
                    
    '高级条件写入
    With Me.lvwComPages
        mstrCondition = mstrCondition & ";" & Me.chkAdvanced.Value
        For intLoop = 1 To .ListItems.Count
            strAdvancedWhere = strAdvancedWhere & "," & Mid(.ListItems(intLoop).Key, 2) & "^" & .ListItems(intLoop).Text & "^" & _
                                .ListItems(intLoop).SubItems(1) & "^" & .ListItems(intLoop).SubItems(2) & "^" & _
                                .ListItems(intLoop).SubItems(3) & "^" & .ListItems(intLoop).SubItems(4)
        Next
    End With
    mstrCondition = mstrCondition & ";" & Mid(strAdvancedWhere, 2)
    
    '保存一些设置到注册表以方便以后调用
    zldatabase.SetPara "使用组合查询", Me.chkAdvanced.Value, 100, 1208
    zldatabase.SetPara "组合查询", Mid(strAdvancedWhere, 2), 100, 1208
    zldatabase.SetPara "是否使用时间", Me.chkDisableDate.Value, 100, 1208
    
    Unload Me
End Sub

Private Sub cmdUpdate_Click()

    With Me.lvwComPages
        If .ListItems.Count <= 0 Then
            MsgBox "列表里没有条件可以更新，请增加一个组合条件！", vbInformation, gstrSysName
            Exit Sub
        Else
            If Me.txtVerifyItem.Tag = "" Then
                MsgBox "请输入检验项目!", vbInformation, gstrSysName
                Me.txtVerifyItem.SetFocus
                Exit Sub
            End If
            .ListItems(.SelectedItem.Index).Key = "A" & Me.txtVerifyItem.Tag
            .ListItems(.SelectedItem.Index).Text = Me.txtVerifyItem.Text
            .ListItems(.SelectedItem.Index).SubItems(1) = Me.cboIF1.Text
            .ListItems(.SelectedItem.Index).SubItems(2) = Me.TxtVariable1.Text
            .ListItems(.SelectedItem.Index).SubItems(3) = Me.cboIF2.Text
            .ListItems(.SelectedItem.Index).SubItems(4) = Me.TxtVariable2.Text
        End If
    End With
    
End Sub

Private Sub DTPVerifyBegin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub DTPVerifyEnd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub Form_Load()
    InitControl
End Sub

Private Sub lvwComPages_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    txtVerifyItem.Tag = Mid(Item.Key, 2)
    txtVerifyItem.Text = Item.Text
    If Item.SubItems(1) <> "" Then
        cboIF1.Text = Item.SubItems(1)
    End If
    TxtVariable1.Text = Item.SubItems(2)
    If Item.SubItems(3) <> "" Then
        cboIF2.Text = Item.SubItems(3)
    End If
    TxtVariable2.Text = Item.SubItems(4)
    
End Sub

Private Sub lvwComPages_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub optOneItem_Click()
    Call cboVerifyType_Click
End Sub

Private Sub OptUnionItem_Click()
    Call cboVerifyType_Click
End Sub

Private Sub sTab_Click(PreviousTab As Integer)
    If PreviousTab = 0 Then
        Me.txtVerifyItem.SetFocus
    End If
End Sub

Private Sub txtAgeBegin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey vbKeyTab
    Else
        KeyAscii = FilterKeyAscii(KeyAscii, 99, "0123456789")
    End If
    
End Sub

Private Sub txtAgeBegin_LostFocus()
    If Me.cboAgeUnit.Text = "" Then
        If Me.txtAgeBegin.Text <> "" Or Me.txtAgeEnd.Text <> "" Then
            Me.cboAgeUnit.ListIndex = 1
        Else
            Me.cboAgeUnit.ListIndex = 0
        End If
    Else
        If Me.txtAgeBegin.Text = "" And Me.txtAgeEnd.Text = "" Then
            Me.cboAgeUnit.ListIndex = 0
        End If
    End If
End Sub

Private Sub txtAgeEnd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey vbKeyTab
    Else
        KeyAscii = FilterKeyAscii(KeyAscii, 99, "0123456789")
    End If
End Sub

Private Sub txtAgeEnd_LostFocus()
    If Me.cboAgeUnit.Text = "" Then
        If Me.txtAgeBegin.Text <> "" Or Me.txtAgeEnd.Text <> "" Then
            Me.cboAgeUnit.ListIndex = 1
        Else
            Me.cboAgeUnit.ListIndex = 0
        End If
    Else
        If Me.txtAgeBegin.Text = "" And Me.txtAgeEnd.Text = "" Then
            Me.cboAgeUnit.ListIndex = 0
        End If
    End If
End Sub

Private Sub txtPatientName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txtSample_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txtSampleID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub





Private Sub TxtVariable_GotFocus()
    TxtVariable1.SelStart = 0
    TxtVariable1.SelLength = Len(TxtVariable1.Text)
End Sub

Private Sub TxtVariable_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(Me.TxtVariable1.Text) <> "" Then
            cmdAdd_Click
        End If
    End If
End Sub

Private Sub TxtVariable1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub TxtVariable2_GotFocus()
    TxtVariable2.SelStart = 0
    TxtVariable2.SelLength = Len(TxtVariable2)
End Sub

Private Sub TxtVariable2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txtVerifyItem_GotFocus()
    txtVerifyItem.SelStart = 0
    txtVerifyItem.SelLength = Len(txtVerifyItem.Text)
End Sub

Private Sub txtVerifyItem_KeyPress(KeyAscii As Integer)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim strVerifyType As String                             '检验类别
    Dim mint简码 As Integer                                 '按那种方式的简码查找
    Dim vRect As RECT                                       '选择框定位
    
    mint简码 = zldatabase.GetPara("简码方式") '简码匹配方式：0-拼音,1-五笔
    
    If KeyAscii = 13 Then
        strSQL = " select distinct C.ID, C.编码, C.中文名 || '(' || C.英文名 || ')' as 名称,b.操作类型　" & _
                 " from 检验报告项目 a , 诊疗项目目录 b ,  诊治所见项目 c ,诊疗项目别名 d  " & _
                 " where a.诊疗项目id = b.id and  a.报告项目id = c.id and b.ID = d.诊疗项目id and B.组合项目 <> 1" & _
                 "       and d.码类 = [2] and (c.英文名 like [1] or b.名称 like [1] or d.简码 like [1] ) order by b.操作类型 "
        
        vRect = GetControlRect(txtVerifyItem.hWnd)
        Set rsTmp = zldatabase.ShowSQLSelect(Me, strSQL, 0, "检验项目", False, "", "", False, False, True, _
                    vRect.Left, vRect.Top, txtVerifyItem.Height, False, False, True, UCase(txtVerifyItem) & "%", mint简码 + 1)
                    
        If Not rsTmp Is Nothing Then
            If rsTmp.State <> 0 Then
                txtVerifyItem.Tag = Nvl(rsTmp("ID"))
                txtVerifyItem.Text = Nvl(rsTmp("名称"))
                zlCommFun.PressKey vbKeyTab
            End If
        Else
            txtVerifyItem.Tag = ""
            txtVerifyItem = ""
        End If
        txtVerifyItem.SelStart = 0
        txtVerifyItem.SelLength = Len(txtVerifyItem.Text)
    End If
End Sub
