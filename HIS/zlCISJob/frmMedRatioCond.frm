VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMedRatioCond 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9615
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   9615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picCond 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4290
      Left            =   3405
      ScaleHeight     =   4290
      ScaleWidth      =   2970
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   90
      Width           =   2970
      Begin VB.CheckBox chkDrug 
         BackColor       =   &H80000005&
         Caption         =   "分类统计时药品区分西药、成药、中药。"
         Height          =   420
         Left            =   75
         TabIndex        =   30
         Top             =   3360
         Width           =   2640
      End
      Begin VB.Frame fraWay 
         BackColor       =   &H80000005&
         Caption         =   "统计方式"
         Height          =   1455
         Left            =   30
         TabIndex        =   20
         Top             =   1725
         Width           =   2670
         Begin VB.PictureBox picPatiType 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   390
            ScaleHeight     =   270
            ScaleWidth      =   2025
            TabIndex        =   31
            Top             =   1125
            Width           =   2025
            Begin VB.OptionButton optPatiType 
               BackColor       =   &H80000005&
               Caption         =   "出院"
               Height          =   255
               Index           =   1
               Left            =   1275
               TabIndex        =   34
               Top             =   -15
               Width           =   660
            End
            Begin VB.OptionButton optPatiType 
               BackColor       =   &H80000005&
               Caption         =   "在院"
               Height          =   255
               Index           =   0
               Left            =   450
               TabIndex        =   33
               Top             =   -15
               Value           =   -1  'True
               Width           =   750
            End
            Begin VB.Label lblPatiType 
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "类型"
               Height          =   180
               Left            =   0
               TabIndex        =   32
               Top             =   0
               Width           =   360
            End
         End
         Begin VB.OptionButton optWay 
            BackColor       =   &H80000005&
            Caption         =   "按病人"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   23
            Top             =   870
            Width           =   1410
         End
         Begin VB.OptionButton optWay 
            BackColor       =   &H80000005&
            Caption         =   "按开单人"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   22
            Top             =   570
            Width           =   1650
         End
         Begin VB.OptionButton optWay 
            BackColor       =   &H80000005&
            Caption         =   "按开单科室"
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   21
            Top             =   255
            Value           =   -1  'True
            Width           =   1275
         End
      End
      Begin VB.Frame fraRange 
         BackColor       =   &H80000005&
         Caption         =   "费用范围"
         Height          =   580
         Left            =   45
         TabIndex        =   16
         Top             =   0
         Width           =   2670
         Begin VB.OptionButton optRan 
            BackColor       =   &H80000005&
            Caption         =   "全院"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   19
            Top             =   270
            Value           =   -1  'True
            Width           =   705
         End
         Begin VB.OptionButton optRan 
            BackColor       =   &H80000005&
            Caption         =   "住院"
            Height          =   195
            Index           =   2
            Left            =   930
            TabIndex        =   18
            Top             =   270
            Width           =   705
         End
         Begin VB.OptionButton optRan 
            BackColor       =   &H80000005&
            Caption         =   "门诊"
            Height          =   195
            Index           =   1
            Left            =   1740
            TabIndex        =   17
            Top             =   270
            Width           =   705
         End
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "查询(&O)"
         Height          =   330
         Left            =   1800
         TabIndex        =   15
         Top             =   3840
         Width           =   870
      End
      Begin VB.ComboBox cboTim 
         Height          =   300
         Left            =   885
         TabIndex        =   14
         Text            =   "Combo2"
         Top             =   660
         Width           =   1800
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   315
         Index           =   1
         Left            =   885
         TabIndex        =   24
         Top             =   1320
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   556
         _Version        =   393216
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   106168323
         CurrentDate     =   41636
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   315
         Index           =   0
         Left            =   885
         TabIndex        =   25
         Top             =   1005
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   556
         _Version        =   393216
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   106168323
         CurrentDate     =   37952
      End
      Begin VB.Label lblEnd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   195
         Index           =   1
         Left            =   645
         TabIndex        =   28
         Top             =   1380
         Width           =   180
      End
      Begin VB.Label lblBegin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "从"
         Height          =   195
         Index           =   0
         Left            =   645
         TabIndex        =   27
         Top             =   1065
         Width           =   180
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "时间范围"
         Height          =   195
         Left            =   60
         TabIndex        =   26
         Top             =   690
         Width           =   720
      End
   End
   Begin VB.PictureBox picDetail 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4065
      Left            =   6510
      ScaleHeight     =   4065
      ScaleWidth      =   2850
      TabIndex        =   0
      Top             =   930
      Width           =   2850
      Begin VB.Frame fraOutTime 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   45
         TabIndex        =   35
         Top             =   735
         Width           =   2760
         Begin VB.ComboBox cboOutTime 
            Height          =   300
            Left            =   795
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   15
            Width           =   1305
         End
         Begin VB.Label lblOutTime 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "出院时间"
            Height          =   180
            Left            =   0
            TabIndex        =   36
            Top             =   60
            Width           =   720
         End
      End
      Begin VB.Frame fraDept 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   340
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   2745
         Begin VB.ComboBox cboDept 
            Height          =   300
            Left            =   435
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   0
            Width           =   1905
         End
         Begin VB.Label lblDept 
            BackColor       =   &H80000005&
            Caption         =   "科室"
            Height          =   285
            Left            =   0
            TabIndex        =   12
            Top             =   45
            Width           =   435
         End
      End
      Begin VB.Frame fraList 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   2880
         Left            =   60
         TabIndex        =   4
         Top             =   1125
         Width           =   2715
         Begin MSComctlLib.ListView lvwPati 
            Height          =   930
            Left            =   45
            TabIndex        =   7
            Top             =   225
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1640
            View            =   3
            Arrange         =   2
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "姓名"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "床号"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "住院号"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "性别"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Text            =   "年龄"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Text            =   "入院时间"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   6
               Text            =   "费别"
               Object.Width           =   2540
            EndProperty
         End
         Begin MSComctlLib.ListView lvwDoc 
            Height          =   630
            Left            =   1425
            TabIndex        =   9
            Top             =   135
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   1111
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "姓名"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "编码"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "性别"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.CommandButton cmdNone 
            Caption         =   "全清"
            Height          =   330
            Left            =   855
            TabIndex        =   6
            TabStop         =   0   'False
            ToolTipText     =   "Ctrl + R"
            Top             =   2535
            Width           =   870
         End
         Begin VB.CommandButton cmdAll 
            Caption         =   "全选"
            Height          =   330
            Left            =   1785
            TabIndex        =   5
            TabStop         =   0   'False
            ToolTipText     =   "Ctrl + A"
            Top             =   2535
            Width           =   870
         End
         Begin MSComctlLib.ListView lvwDept 
            Height          =   870
            Left            =   1380
            TabIndex        =   8
            Top             =   960
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   1535
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "名称"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "编码"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "简码"
               Object.Width           =   2540
            EndProperty
         End
         Begin MSComctlLib.ListView lvwPatiOut 
            Height          =   930
            Left            =   90
            TabIndex        =   38
            Top             =   1260
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1640
            View            =   3
            Arrange         =   2
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "姓名"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "住院号"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "性别"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "年龄"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Text            =   "入院时间"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "出院时间"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "费别"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame fraDoc 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   340
         Left            =   15
         TabIndex        =   1
         Top             =   360
         Width           =   2715
         Begin VB.ComboBox cboDoc 
            Height          =   300
            Left            =   420
            TabIndex        =   2
            Text            =   "cboDoc"
            Top             =   0
            Width           =   1905
         End
         Begin VB.Label lblDoc 
            BackColor       =   &H80000005&
            Caption         =   "医生"
            Height          =   225
            Left            =   0
            TabIndex        =   3
            Top             =   45
            Width           =   390
         End
      End
   End
   Begin XtremeSuiteControls.TaskPanel tkpMain 
      Height          =   4530
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   3390
      _Version        =   589884
      _ExtentX        =   5980
      _ExtentY        =   7990
      _StockProps     =   64
      VisualTheme     =   6
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
End
Attribute VB_Name = "frmMedRatioCond"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event DoQuery(ByVal bytRan As Byte, ByVal bytWay As Byte, ByVal lngIDs As String, ByVal datBegin As Date, ByVal datEnd As Date)
Public Event CountWay(ByVal strWay As String, ByVal blnDrug As Boolean)

Private mstrPrivs As String

Private mstrDepParIDs As String
Private mstrDocParIDs As String
Private mstrPatParIDs As String

Private mstrPreDocDepID As String

Private mstrPrePatDepID As String
Private mstrPrePatDocID As String

Private mdatOutBegin As Date, mdatOutEnd As Date
Private mintOutPreTime As Integer '上一次选择的时间列表的值
Private mdatCurr As Date
Private mblnFirst As Boolean '第一次出现院病人列表

Private Enum SeaRan '查询范围
    ran全院 = 0
    ran住院 = 2
    ran门诊 = 1
End Enum

Private Enum SeaWay '查询方式
    way开单科室 = 0
    way开单人 = 1
    way病人 = 2
End Enum

Private Enum mCtlID
    opt_病人类型_在院 = 0
    opt_病人类型_出院 = 1
End Enum

Private Sub Form_Load()
    Dim objGroup As TaskPanelGroup
    Dim objItem As TaskPanelGroupItem
    Dim objPane As Pane
    Dim strDiffDate, strTmp As String
    Dim i, k As Integer
    Dim intIndex As Integer
    
    mstrPrivs = gstrPrivs
    Me.Width = tkpMain.Width: Me.Height = tkpMain.Height

    '分组控件------------------------------------------
    Call tkpMain.SetMargins(8, 8, 8, 8, 8)

    Set objGroup = tkpMain.Groups.Add(1, "条件列表")
    objGroup.Expandable = False
    Set objItem = objGroup.Items.Add(0, "", xtpTaskItemTypeControl)
    Set objItem.Control = picCond
    picCond.BackColor = objItem.BackColor
    
    Set objGroup = tkpMain.Groups.Add(2, "开单科室列表")
    objGroup.Expandable = False
    Set objItem = objGroup.Items.Add(0, "", xtpTaskItemTypeControl)
    Set objItem.Control = picDetail
    picDetail.BackColor = objItem.BackColor
    
    On Error GoTo errH
    i = Val(zlDatabase.GetPara("药比统计范围", glngSys, 1261, "0"))
    optRan(i).Value = True
    
    If i = ran门诊 Then  '查询范围是门诊不存在按病人查询
        optWay(way病人).Enabled = False
        optWay(way病人).Value = False
    Else
        optWay(way病人).Enabled = True
    End If
    
    chkDrug.Value = Val(zlDatabase.GetPara("药品分别统计", glngSys, 1261, 1))
 
    mdatCurr = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    cboOutTime.Clear '出院病人时间范围
    With cboOutTime
        .AddItem "今天内"
        .ItemData(.NewIndex) = 0
        .AddItem "昨天内"
        .ItemData(.NewIndex) = 1
        .AddItem "前天内"
        .ItemData(.NewIndex) = 2
        .AddItem "一周内"
        .ItemData(.NewIndex) = 7
        .AddItem "30天内"
        .ItemData(.NewIndex) = 30
        .AddItem "60天内"
        .ItemData(.NewIndex) = 60
        .AddItem "[指定...]"
        .ItemData(.NewIndex) = -1
    End With
    If cboOutTime.ListCount > 0 Then cboOutTime.ListIndex = 0
    
    With cboTim
        .AddItem "本月"
        .ItemData(.NewIndex) = 1
        .AddItem "最近两个月"
        .ItemData(.NewIndex) = 2
        .AddItem "最近一季度"
        .ItemData(.NewIndex) = 3
        .AddItem "最近两季度"
        .ItemData(.NewIndex) = 6
        .AddItem "指定[...]"
        .ItemData(.NewIndex) = -1
    End With
    strDiffDate = zlDatabase.GetPara("药比查询间隔", glngSys, 1261, "0")
    If IsNumeric(strDiffDate) Then
        Select Case strDiffDate
            Case "1", "0"
                cboTim.ListIndex = 0
            Case "2"
                cboTim.ListIndex = 1
            Case "3"
                cboTim.ListIndex = 2
            Case "6"
                cboTim.ListIndex = 3
        End Select
    Else
        Call Cbo.SetIndex(cboTim.hwnd, 4)
        dtpDate(0).Value = Format(Split(strDiffDate, "<Tab>")(0), "yyyy-MM-dd HH:mm")
        dtpDate(1).Value = Format(Split(strDiffDate, "<Tab>")(1), "yyyy-MM-dd HH:mm")
        dtpDate(0).Enabled = True
        dtpDate(1).Enabled = True
    End If
    
    mstrDepParIDs = zlDatabase.GetPara("药比开单科室", glngSys, 1261, "")
    
    mstrDocParIDs = zlDatabase.GetPara("药比开单人", glngSys, 1261, "")
    If InStr(mstrDocParIDs, "|") > 0 Then
        mstrPreDocDepID = Split(mstrDocParIDs, "|")(0)
        mstrDocParIDs = Split(mstrDocParIDs, "|")(1)
    End If
    
    mstrPatParIDs = zlDatabase.GetPara("药比病人", glngSys, 1261, "")
    If InStr(mstrPatParIDs, "|") > 0 Then
        strTmp = Split(mstrPatParIDs, "|")(0)
        If InStr(strTmp, ",") > 0 Then
            mstrPrePatDepID = Split(strTmp, ",")(0)
            mstrPrePatDocID = Split(strTmp, ",")(1)
        End If
        mstrPatParIDs = Split(mstrPatParIDs, "|")(1)
    End If
    
    i = Val(zlDatabase.GetPara("药比统计方式", glngSys, 1261, "0"))
    optWay(i).Value = True
    
    intIndex = i
    Call optWay_Click(intIndex)
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cboOutTime_Click()
'设置时间范围
    Dim intDateCount As Integer
    intDateCount = cboOutTime.ItemData(cboOutTime.ListIndex)
    
    If cboOutTime.ListIndex = mintOutPreTime And mintOutPreTime <> 6 Then Exit Sub
    If intDateCount = -1 Then
        If mdatOutBegin = CDate(0) Then
            mdatOutBegin = mdatCurr
            mdatOutEnd = mdatCurr
        End If
        If Not frmSelectTime.ShowMe(Me, mdatOutBegin, mdatOutEnd, cboOutTime) Then
            '取消时恢复原来的选择
            Call Cbo.SetIndex(cboOutTime.hwnd, mintOutPreTime)
            Exit Sub
        End If
    Else
        mdatOutEnd = mdatCurr
        mdatOutBegin = mdatOutEnd - intDateCount
    End If
    
    If mdatOutBegin = CDate(0) Or mdatOutEnd = CDate(0) Then
        cboOutTime.ToolTipText = ""
    Else
        cboOutTime.ToolTipText = "范围：" & Format(mdatOutBegin, "yyyy-MM-dd") & " 至 " & Format(mdatOutEnd, "yyyy-MM-dd")
    End If

    mintOutPreTime = cboOutTime.ListIndex
    
    Call LoadOutPati
End Sub

Private Sub LoadOutPati(Optional ByVal blnFirst As Boolean)
'功能：加载出院病人
'参数：blnFirst 是否是第一次出现出院病人列表，如果是未找到病人时不提示
    Dim objListItem As ListItem
    Dim strSQL, strTmp As String
    Dim rsTmp As ADODB.Recordset
    Dim strDocName, strPatParIDs, strPar As String
    Dim k As Integer
         
    strSQL = "Select a.病人id,a.主页id,a.住院号,a.姓名,a.性别,a.年龄,a.入院日期,a.出院日期,a.费别 From 病案主页 A where a.出院科室id=[1]"
    
    If cboDoc.Text <> "所有医生" Then
        strSQL = strSQL & " And a.住院医师 = [2]"
        strPar = Split(cboDoc.Text, "-")(1)
    End If
    
    strSQL = strSQL & " and a.出院日期 between [3] and [4]"
    
    strSQL = strSQL & "  Order By a.出院日期 desc"
 
    
    On Error GoTo errH
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cboDept.ItemData(cboDept.ListIndex), strPar, mdatOutBegin, CDate(Format(mdatOutEnd, "YYYY-MM-DD 23:59:59")))
    
    lvwPatiOut.ListItems.Clear
    
    If rsTmp.EOF Then
        If Not blnFirst Then MsgBox "当前条件下未找到出院病人！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    k = 0
    Do While Not rsTmp.EOF
        Set objListItem = lvwPatiOut.ListItems.Add(, "_" & rsTmp!病人ID & "_" & rsTmp!主页ID, "" & rsTmp!姓名)
            objListItem.SubItems(1) = "" & rsTmp!住院号
            objListItem.SubItems(2) = "" & rsTmp!性别
            objListItem.SubItems(3) = "" & rsTmp!年龄
            objListItem.SubItems(4) = "" & rsTmp!入院日期
            objListItem.SubItems(5) = "" & rsTmp!出院日期
            objListItem.SubItems(6) = "" & rsTmp!费别
        rsTmp.MoveNext
    Loop
    Screen.MousePointer = 0
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        If Me.ActiveControl Is lvwDept Then
            Call cmdAll_Click
        ElseIf Me.ActiveControl Is lvwDoc Then
            Call cmdAll_Click
        ElseIf Me.ActiveControl Is lvwPati Then
            Call cmdAll_Click
        End If
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        If Me.ActiveControl Is lvwDoc Then
            Call cmdNone_Click
        ElseIf Me.ActiveControl Is lvwPati Then
            Call cmdNone_Click
        ElseIf Me.ActiveControl Is lvwDept Then
            Call cmdNone_Click
        End If
    ElseIf KeyCode = 13 Then
        Call ZLCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub LoadData(ByVal bytRan As Byte, ByVal bytWay As Byte)
'功能：加载数据，列表和下拉列表
'参数：bytRan 范围,0-全院，1-门诊，2-住院
'      bytPro 费用性质 0-结帐，1-实收
'      bytWay 统计方式 0-科室，1-开单人，2-病人
    Dim strSQL, strTmp As String
    Dim rsTmp As ADODB.Recordset
    Dim objListItem As ListItem
    Dim strDepIDs As String
    Dim strDepParIDs As String
    Dim strDocName As String
    Dim strPreDeptID As String
    Dim i, k As Integer
    
    Screen.MousePointer = 11
    strSQL = "Select Distinct a.Id, a.名称, a.编码, a.简码 From 部门表 A, 部门性质说明 B Where a.Id = b.部门id And b.工作性质='临床' And (a.撤档时间 is NULL or Trunc(a.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))"
    If bytRan = 0 Then
        strTmp = " And b.服务对象 <> 0"
    ElseIf bytRan = 1 Then
        strTmp = "  And b.服务对象 = 1"
    ElseIf bytRan = 2 Then
        strTmp = " And (b.服务对象 = 3 Or b.服务对象 = 2)"
    End If
    If InStr(";" & mstrPrivs & ";", ";全院病人;") = 0 Then
        strTmp = strTmp & " And a.Id in (Select t.部门id From 部门人员 T Where t.人员id = [1])"
    End If
    
    strSQL = strSQL & strTmp & " Order By a.名称"
    
    strDepParIDs = mstrDepParIDs
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    
    strDepIDs = GetUser科室IDs(False)
    lvwDept.ListItems.Clear
    cboDept.Clear
    Do While Not rsTmp.EOF
        Set objListItem = lvwDept.ListItems.Add(, "_" & rsTmp!ID, "" & rsTmp!名称)
            objListItem.SubItems(1) = "" & rsTmp!编码
            objListItem.SubItems(2) = "" & rsTmp!简码
            If InStr("," & strDepIDs & ",", "," & rsTmp!ID & ",") <> 0 And strDepParIDs = "" Then
                objListItem.Checked = True
                If k = 0 Then '为了看到有选择的
                    objListItem.EnsureVisible
                    objListItem.Selected = True
                    k = 1
                End If
            End If
            If InStr("," & strDepParIDs & ",", "," & rsTmp!ID & ",") <> 0 Then
                objListItem.Checked = True
                If k = 0 Then '为了看到有选择的
                    objListItem.EnsureVisible
                    objListItem.Selected = True
                    k = 1
                End If
            End If
            If bytWay <> 0 Then
                With cboDept
                    .AddItem rsTmp!编码 & "-" & rsTmp!名称
                    .ItemData(.NewIndex) = rsTmp!ID
                    If InStr("," & strDepIDs & ",", "," & rsTmp!ID & ",") <> 0 Then Call Cbo.SetIndex(cboDept.hwnd, .NewIndex)
                End With
            End If
        rsTmp.MoveNext
    Loop
    If cboDept.ListCount > 0 Then
        If cboDept.ListIndex = -1 Then Call Cbo.SetIndex(cboDept.hwnd, 0)
        If bytWay = 1 Then Call Cbo.Locate(cboDept, mstrPreDocDepID, True)
        If bytWay = 2 Then Call Cbo.Locate(cboDept, mstrPrePatDepID, True)
        Call cboDept_Click
    End If
    Screen.MousePointer = 0
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cboDept_Click()
'加载医生列表
    Dim objListItem As ListItem
    Dim strSQL, strTmp As String
    Dim strDocName, strDocParIDs As String
    Dim rsTmp As ADODB.Recordset
    Dim bytPrivDoc As Byte '医生权限
    Dim i, k As Integer
    
    For i = 0 To 2
        If optWay(i).Value Then Exit For
    Next i
    
    strDocParIDs = mstrDocParIDs
    strSQL = "Select a.Id, a.编号, a.姓名, a.性别 From 人员表 A, 部门人员 B, 人员性质说明 C" & vbNewLine & _
        "Where a.Id = b.人员id And b.人员id = c.人员id And c.人员性质 = '医生' And b.部门id = [1] And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null)" & vbNewLine & _
        "Order By a.姓名"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cboDept.ItemData(cboDept.ListIndex))
    lvwDoc.ListItems.Clear
    Do While Not rsTmp.EOF
        Set objListItem = lvwDoc.ListItems.Add(, "_" & rsTmp!ID, "" & rsTmp!姓名)
            objListItem.SubItems(1) = "" & rsTmp!编号
            objListItem.SubItems(2) = "" & rsTmp!性别
            If UserInfo.ID = rsTmp!ID And strDocParIDs = "" Then
                objListItem.Checked = True
                If k = 0 Then '为了看到有选择的
                    objListItem.EnsureVisible
                    objListItem.Selected = True
                    k = 1
                End If
            End If
            If InStr("," & strDocParIDs & ",", "," & rsTmp!ID & ",") <> 0 Then
                objListItem.Checked = True
                If k = 0 Then '为了看到有选择的
                    objListItem.EnsureVisible
                    objListItem.Selected = True
                    k = 1
                End If
            End If
        rsTmp.MoveNext
    Loop
    If optWay(2).Value Then
        strSQL = "Select a.Id, a.编号, a.姓名, a.性别, c.人员性质" & vbNewLine & _
            "From 人员表 A, 部门人员 B, 人员性质说明 C" & vbNewLine & _
            "Where a.Id = b.人员id And a.Id(+) = c.人员id And b.部门id = [1] And c.人员性质 = '医生' Order By a.姓名"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cboDept.ItemData(cboDept.ListIndex))
        cboDoc.Clear
        If InStr(";" & mstrPrivs & ";", ";全院病人;") <> 0 Then
            With cboDoc
                .AddItem "所有医生"
                .ItemData(.NewIndex) = 0
            End With
            If InStr(";" & mstrPrivs & ";", ";本科病人;") = 0 Then
                bytPrivDoc = 1
            End If
        End If
        Do While Not rsTmp.EOF
            With cboDoc
                .AddItem rsTmp!编号 & "-" & rsTmp!姓名
                .ItemData(.NewIndex) = rsTmp!ID
                If UserInfo.ID = rsTmp!ID Then Call Cbo.SetIndex(cboDoc.hwnd, .NewIndex)
            End With
            rsTmp.MoveNext
        Loop
        If bytPrivDoc = 1 Then
            cboDoc.Clear
            With cboDoc
                .AddItem UserInfo.编号 & "_" & UserInfo.姓名
                .ItemData(.NewIndex) = UserInfo.ID
            End With
        End If
        If cboDoc.ListCount > 0 Then
            If Not Cbo.Locate(cboDoc, mstrPrePatDocID, True) Then cboDoc.ListIndex = 0
        End If
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cboDoc_Click()
'加载病人列表
    Dim objListItem As ListItem
    Dim strSQL, strTmp As String
    Dim rsTmp As ADODB.Recordset
    Dim strDocName, strPatParIDs, strPar As String
    Dim k As Integer
    
    If optPatiType(opt_病人类型_出院).Value And optPatiType(opt_病人类型_出院).Enabled Then
        Call LoadOutPati
        Exit Sub
    End If
    
    strSQL = "Select a.病人id, b.主页id, LPAD(a.当前床号,10,' ') as 当前床号,a.住院号, NVL(B.姓名,A.姓名) 姓名 ,NVL(B.性别,A.性别) 性别,NVL(B.年龄, A.年龄) 年龄, a.入院时间, a.费别" & vbNewLine & _
        "From 病人信息 A, 病案主页 B, 在院病人 C" & vbNewLine & _
        "Where a.病人id = c.病人id And a.病人id = b.病人id And a.主页id = b.主页id And c.科室id = [1]"
    
    If cboDoc.Text <> "所有医生" Then
        strSQL = strSQL & " And b.住院医师 = [2]"
        strPar = Split(cboDoc.Text, "-")(1)
    End If
    strSQL = strSQL & "  Order By 当前床号"
    strPatParIDs = mstrPatParIDs
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cboDept.ItemData(cboDept.ListIndex), strPar)
    lvwPati.ListItems.Clear

    k = 0
    Do While Not rsTmp.EOF
        Set objListItem = lvwPati.ListItems.Add(, "_" & rsTmp!病人ID & "_" & rsTmp!主页ID, "" & rsTmp!姓名)
            objListItem.SubItems(1) = "" & rsTmp!当前床号
            objListItem.SubItems(2) = "" & rsTmp!住院号
            objListItem.SubItems(3) = "" & rsTmp!性别
            objListItem.SubItems(4) = "" & rsTmp!年龄
            objListItem.SubItems(5) = "" & rsTmp!入院时间
            objListItem.SubItems(6) = "" & rsTmp!费别
            If UserInfo.姓名 = strDocName And strPatParIDs = "" Then
                objListItem.Checked = True
                If k = 0 Then '为了看到有选择的
                    objListItem.EnsureVisible
                    objListItem.Selected = True
                    k = 1
                End If
            End If
            If InStr("," & strPatParIDs & ",", "," & rsTmp!病人ID & ",") <> 0 Then
                objListItem.Checked = True
                If k = 0 Then '为了看到有选择的
                    objListItem.EnsureVisible
                    objListItem.Selected = True
                    k = 1
                End If
            End If
        rsTmp.MoveNext
    Loop
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub optPatiType_Click(Index As Integer)
    Call picDetail_Resize
    
    If Not mblnFirst And optPatiType(opt_病人类型_出院).Value Then
        mblnFirst = True
        mdatOutBegin = mdatCurr
        mdatOutEnd = mdatCurr
        cboOutTime.ToolTipText = "范围：" & Format(mdatOutBegin, "yyyy-MM-dd") & " 至 " & Format(mdatOutEnd, "yyyy-MM-dd")
        Call LoadOutPati(True)
    End If
End Sub

Private Sub optRan_Click(Index As Integer)
'功能：确定什么查询方式 按开单科室 按开单人 按病人
    Dim i As Integer
    Dim intTmp As Integer
    
    intTmp = -1
    
    If Index = ran门诊 Then
        If optWay(way病人).Value Then
            optWay(way开单科室).Value = True
            intTmp = way开单科室
        End If
        optWay(way病人).Enabled = False
        optWay(way病人).Value = False
    Else
        optWay(way病人).Enabled = True
        intTmp = way病人
    End If
    
    If intTmp = -1 Then intTmp = IIf(optWay(way开单科室).Value, way开单科室, way开单人)
    
    Call LoadData(Index, intTmp)
End Sub

Private Sub cmdSearch_Click()
    Dim i As Integer
    Dim intCount As Integer
    Dim strIDs As String
    Dim bytRan As Byte '费用范围
    Dim bytPro As Byte '1-结帐金额，2-实收金额
    Dim blnDebt As Boolean '包含划价单
    Dim bytWay As Byte '统计方式
    Dim strNow As String
    Dim strWorkTim As String
    Dim strTmp As String
    Dim objLvw As Object

    strIDs = ""
    
    On Error GoTo errH
    
    strWorkTim = zlDatabase.GetPara("上午上下班时间", glngSys)
    
    If strWorkTim = "" Then strWorkTim = "08:00 AND 12:00"
    
    strNow = Format(zlDatabase.Currentdate, "hh:mm")

    If Split(strWorkTim, " AND ")(0) < strNow And Split(strWorkTim, " AND ")(1) > strNow And Not optWay(2).Value Then
        MsgBox "目前处于上午上班时间，不允许按开单科室查询和按开单人查询！", vbInformation, gstrSysName
        
        '查询范围不是门诊重新设置过滤条件
        If Not optRan(ran门诊).Value Then optWay(way病人).Value = True
        
        Exit Sub
    End If
    
    If optRan(ran全院).Value Then
        bytRan = 0 '全院
    ElseIf optRan(ran住院).Value Then
        bytRan = 2 '住院
    ElseIf optRan(ran门诊).Value Then
        bytRan = 1 '门诊
    End If
    
    If lvwPati.Visible Or lvwPatiOut.Visible Then
        If lvwPati.Visible Then
            Set objLvw = lvwPati
        Else
            Set objLvw = lvwPatiOut
        End If
                
        With objLvw
            For i = 1 To .ListItems.Count
                If .ListItems(i).Checked Then
                    strIDs = strIDs & "," & Split(.ListItems(i).Key, "_")(1) & ":" & Split(.ListItems(i).Key, "_")(2)
                    strTmp = strTmp & "," & Split(.ListItems(i).Key, "_")(1)
                End If
            Next
        End With
        
        bytWay = 2
        If strIDs = "" Then MsgBox "请选择病人", vbInformation, gstrSysName: Exit Sub
        mstrPrePatDepID = cboDept.ItemData(cboDept.ListIndex)
        mstrPrePatDocID = cboDoc.ItemData(cboDoc.ListIndex)
        mstrPatParIDs = Mid(strTmp, 2)
    ElseIf lvwDept.Visible Then
        For i = 1 To lvwDept.ListItems.Count
            If lvwDept.ListItems(i).Checked Then
                strIDs = strIDs & "," & Split(lvwDept.ListItems(i).Key, "_")(1)
            End If
        Next
        bytWay = 0
        If strIDs = "" Then MsgBox "请选择科室", vbInformation, gstrSysName: Exit Sub
        mstrDepParIDs = Mid(strIDs, 2)
    ElseIf lvwDoc.Visible Then
        For i = 1 To lvwDoc.ListItems.Count
            If lvwDoc.ListItems(i).Checked Then
                strIDs = strIDs & "," & Split(lvwDoc.ListItems(i).Key, "_")(1)
            End If
        Next
        bytWay = 1
        If strIDs = "" Then MsgBox "请选择开单人", vbInformation, gstrSysName: Exit Sub
        mstrPreDocDepID = cboDept.ItemData(cboDept.ListIndex)
        mstrDocParIDs = Mid(strIDs, 2)
    End If
    
    strIDs = Mid(strIDs, 2)
    
    Call zlDatabase.SetPara("药比统计范围", bytRan, glngSys, 1261)
    Call zlDatabase.SetPara("药品分别统计", chkDrug.Value, glngSys, 1261)
'    Call zlDatabase.SetPara("药比含划价单", chkNotPay.Value, glngSys, 1261)
    Call zlDatabase.SetPara("药比统计方式", bytWay, glngSys, 1261)
    
    If cboTim.ItemData(cboTim.ListIndex) = -1 Then
        Call zlDatabase.SetPara("药比查询间隔", dtpDate(0).Value & "<Tab>" & dtpDate(1).Value, glngSys, 1261)
    Else
        intCount = cboTim.ItemData(cboTim.ListIndex)
        Call zlDatabase.SetPara("药比查询间隔", intCount, glngSys, 1261)
    End If
    
    RaiseEvent DoQuery(bytRan, bytWay, strIDs, dtpDate(0).Value, dtpDate(1).Value)
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errH
    
    Call zlDatabase.SetPara("药比开单科室", mstrDepParIDs, glngSys, 1261)
    mstrDocParIDs = mstrPreDocDepID & "|" & mstrDocParIDs
    Call zlDatabase.SetPara("药比开单人", mstrDocParIDs, glngSys, 1261)
    mstrPatParIDs = mstrPrePatDepID & "," & mstrPrePatDocID & "|" & mstrPatParIDs
    Call zlDatabase.SetPara("药比病人", mstrPatParIDs, glngSys, 1261)
    
    mstrPrivs = ""
    mstrDepParIDs = ""
    mstrDocParIDs = ""
    mstrPatParIDs = ""
    mstrPreDocDepID = ""
    mstrPrePatDepID = ""
    mstrPrePatDocID = ""
    mblnFirst = False
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub optWay_Click(Index As Integer)
'功能：点选统计方式  按开单科室，按开单人，按病人
    Dim strWay As String
    Dim i As Integer
    
    Select Case Index
        Case 0
            strWay = "开单科室"
            tkpMain.Groups.Find(2).Caption = "开单科室列表"
        Case 1
            strWay = "开单人"
            tkpMain.Groups.Find(2).Caption = "开单人列表"
        Case 2
            strWay = "病人"
            tkpMain.Groups.Find(2).Caption = "病人列表"
    End Select
    
    For i = 0 To 2
        If optRan(i).Value Then Exit For
    Next i
    
    Call LoadData(i, Index)
    
    Call picDetail_Resize
    
    RaiseEvent CountWay(strWay, chkDrug.Value = 1)
    
End Sub

Private Sub picDetail_Resize()
    
    On Error Resume Next
    
    If optWay(way开单科室).Value Then
            
        lvwPati.Visible = False
        lvwDoc.Visible = False
        lvwDept.Visible = True
        
        fraDept.Visible = False
        fraDoc.Visible = False
        fraList.Top = 0
        picPatiType.Enabled = False
        optPatiType(opt_病人类型_在院).Enabled = False
        optPatiType(opt_病人类型_出院).Enabled = False
        
    ElseIf optWay(way开单人).Value Then
            
        lvwPati.Visible = False
        lvwDoc.Visible = True
        lvwDept.Visible = False
        
        fraDoc.Visible = False
        fraDept.Visible = True
        fraDept.Left = 0
        fraDept.Top = 0
        fraList.Top = 350
        picPatiType.Enabled = False
        optPatiType(opt_病人类型_在院).Enabled = False
        optPatiType(opt_病人类型_出院).Enabled = False
    ElseIf optWay(way病人).Value Then
        
        lvwPati.Visible = True
        lvwDoc.Visible = False
        lvwDept.Visible = False
        
        fraDept.Visible = True
        fraDoc.Visible = True
        fraDept.Left = 0
        fraDoc.Left = 0
        fraDept.Top = 0
        fraDoc.Top = fraDept.Height
        fraList.Top = fraDept.Height + fraDoc.Height + 10
        
        picPatiType.Enabled = True
        optPatiType(opt_病人类型_在院).Enabled = True
        optPatiType(opt_病人类型_出院).Enabled = True
    End If
    
    If optWay(way病人).Value Then
        cboTim.Enabled = False
        dtpDate(0).Enabled = False
        dtpDate(1).Enabled = False
    Else
        cboTim.Enabled = True
        With cboTim
            If .ItemData(.ListIndex) <> -1 Then
                dtpDate(0).Enabled = False
                dtpDate(1).Enabled = False
            Else
                dtpDate(0).Enabled = True
                dtpDate(1).Enabled = True
            End If
        End With
    End If
    
    If picPatiType.Enabled Then
        If optPatiType(opt_病人类型_出院).Value Then
            fraOutTime.Visible = True
            lvwPati.Visible = False
            lvwPatiOut.Visible = True
            fraOutTime.Width = picDetail.Width
            fraOutTime.Left = 0
            cboOutTime.Width = 1530
            fraOutTime.Top = fraDoc.Top + fraDoc.Height
            fraOutTime.Height = 350
            fraList.Top = fraOutTime.Top + fraOutTime.Height
        Else
            fraOutTime.Visible = False
            lvwPatiOut.Visible = False
            fraList.Top = fraDoc.Top + fraDoc.Height
        End If
    Else
        fraOutTime.Visible = False
        lvwPatiOut.Visible = False
    End If
    
    fraList.Left = 0
    fraList.Width = Me.ScaleWidth
    fraList.Height = Me.ScaleHeight
 
    lvwPati.Left = 0
    lvwPati.Width = fraList.Width - 800
    lvwPati.Height = 2490
    lvwPati.Top = 0
    
    
    lvwPatiOut.Left = 0
    lvwPatiOut.Width = fraList.Width - 800
    lvwPatiOut.Height = 2490
    lvwPatiOut.Top = 0
    
 
    lvwDept.Left = 0
    lvwDept.Width = fraList.Width - 800
    lvwDept.Height = 2490
    lvwDept.Top = 0
 
    lvwDoc.Left = 0
    lvwDoc.Width = fraList.Width - 800
    lvwDoc.Height = 2490
    lvwDoc.Top = 0
 
    cmdAll.Left = lvwDoc.Width - cmdAll.Width - 100
    cmdNone.Left = cmdAll.Left - 30 - cmdAll.Width
    cmdNone.Top = cmdAll.Top
End Sub

Private Sub picCond_Resize()
    On Error Resume Next
    
    picCond.Height = picCond.Height
    fraRange.Width = picCond.ScaleWidth - fraRange.Left
    fraRange.Height = 580
    fraRange.Top = 0
    
    fraWay.Width = fraRange.Width
    chkDrug.Width = fraWay.Width
    dtpDate(0).Width = picCond.ScaleWidth - dtpDate(0).Left
    dtpDate(1).Width = dtpDate(0).Width
    dtpDate(2).Width = dtpDate(0).Width
    cboTim.Width = dtpDate(0).Width
    cmdSearch.Width = cmdAll.Width
    cmdSearch.Left = fraRange.Width + fraRange.Left - cmdAll.Width
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    tkpMain.Left = 0
    tkpMain.Top = 0
    tkpMain.Width = Me.ScaleWidth
    tkpMain.Height = Me.ScaleHeight
End Sub

Private Sub chkDrug_Click()
    Dim strWay As String
    
    If optWay(0).Value Then
        strWay = "开单科室"
    ElseIf optWay(1).Value Then
        strWay = "开单人"
    ElseIf optWay(2).Value Then
        strWay = "病人"
    End If
    
    RaiseEvent CountWay(strWay, chkDrug.Value = 1)
End Sub

Private Sub cboTim_Click()
    Dim curDate As Date
    
    On Error GoTo errH
    curDate = zlDatabase.Currentdate
    With cboTim
        If .ItemData(.ListIndex) <> -1 Then
            dtpDate(0).Enabled = False
            dtpDate(1).Enabled = False
            dtpDate(0).Value = Format(DateAdd("m", -1 * .ItemData(.ListIndex) + 1, curDate), "yyyy-MM-1 00:00")
            dtpDate(1).Value = Format(curDate, "yyyy-MM-dd HH:mm")
        Else
            dtpDate(0).Enabled = True
            dtpDate(1).Enabled = True
            dtpDate(0).SetFocus
            dtpDate(0).Value = Format(curDate, "yyyy-MM-1 00:00")
            dtpDate(1).Value = Format(curDate, "yyyy-MM-dd HH:mm")
        End If
    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdNone_Click()
    If lvwPati.Visible Then
        Call SelectLVW(lvwPati, False)
    ElseIf lvwPatiOut.Visible Then
        Call SelectLVW(lvwPatiOut, False)
    ElseIf lvwDept.Visible Then
        Call SelectLVW(lvwDept, False)
    ElseIf lvwDoc.Visible Then
        Call SelectLVW(lvwDoc, False)
    End If
End Sub

Private Sub cmdAll_Click()
    If lvwPati.Visible Then
        Call SelectLVW(lvwPati, True)
    ElseIf lvwPatiOut.Visible Then
        Call SelectLVW(lvwPatiOut, True)
    ElseIf lvwDept.Visible Then
        Call SelectLVW(lvwDept, True)
    ElseIf lvwDoc.Visible Then
        Call SelectLVW(lvwDoc, True)
    End If
End Sub

Private Sub SelectLVW(objLvw As Object, ByVal blnCheck As Boolean)
    Dim i As Long
    
    For i = 1 To objLvw.ListItems.Count
        objLvw.ListItems(i).Checked = blnCheck
    Next
End Sub

Private Sub lvwDept_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call zlControl.LvwSortColumn(lvwDept, ColumnHeader.Index)
End Sub

Private Sub lvwDoc_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call zlControl.LvwSortColumn(lvwDoc, ColumnHeader.Index)
End Sub

Private Sub lvwPati_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call zlControl.LvwSortColumn(lvwPati, ColumnHeader.Index)
End Sub

Private Sub lvwDept_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Item.Checked = IIf(Item.Checked, False, True)
End Sub

Private Sub lvwDoc_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Item.Checked = IIf(Item.Checked, False, True)
End Sub

Private Sub lvwPati_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Item.Checked = IIf(Item.Checked, False, True)
End Sub
