VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSchQuery 
   Caption         =   "检查预约查询"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11055
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSchQuery.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   11055
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdQuit 
      Caption         =   "退出"
      Height          =   375
      Left            =   6840
      TabIndex        =   3
      Top             =   7560
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfSchedule 
      Height          =   5295
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   10815
      _cx             =   19076
      _cy             =   9340
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   10815
      Begin VB.ComboBox cboDatePeriod 
         Height          =   330
         ItemData        =   "frmSchQuery.frx":0442
         Left            =   1200
         List            =   "frmSchQuery.frx":0444
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   1372
         Width           =   1575
      End
      Begin VB.TextBox txtName 
         Height          =   375
         Left            =   1200
         TabIndex        =   21
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox cboSchDevice 
         Height          =   330
         ItemData        =   "frmSchQuery.frx":0446
         Left            =   1200
         List            =   "frmSchQuery.frx":0448
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   817
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker dpDateStart 
         Height          =   375
         Left            =   4440
         TabIndex        =   19
         Top             =   1350
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   57933825
         CurrentDate     =   43286
      End
      Begin VB.TextBox txtOutPatientNo 
         Height          =   375
         Left            =   4440
         TabIndex        =   18
         Top             =   795
         Width           =   1575
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "清空"
         Height          =   375
         Left            =   9480
         TabIndex        =   17
         Top             =   1080
         Width           =   1100
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "查询"
         Default         =   -1  'True
         Height          =   375
         Left            =   9480
         TabIndex        =   16
         Top             =   480
         Width           =   1100
      End
      Begin MSComCtl2.DTPicker dpDateEnd 
         Height          =   375
         Left            =   7320
         TabIndex        =   15
         Top             =   1350
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   57933825
         CurrentDate     =   43286
      End
      Begin VB.CheckBox chkDatePeriod 
         Caption         =   "从"
         Height          =   375
         Left            =   3480
         TabIndex        =   14
         Top             =   1350
         Width           =   615
      End
      Begin VB.TextBox txtInPatientNo 
         Height          =   375
         Left            =   7320
         TabIndex        =   12
         Top             =   795
         Width           =   1575
      End
      Begin VB.TextBox txtCheckNo 
         Height          =   375
         Left            =   7320
         TabIndex        =   8
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtSchNumber 
         Height          =   375
         Left            =   4440
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "到"
         Height          =   195
         Left            =   6600
         TabIndex        =   23
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label7 
         Caption         =   "预约日期"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "住院号"
         Height          =   195
         Left            =   6600
         TabIndex        =   11
         Top             =   885
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "门诊号"
         Height          =   195
         Left            =   3480
         TabIndex        =   10
         Top             =   885
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "预约设备"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   885
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "检查号"
         Height          =   195
         Left            =   6600
         TabIndex        =   7
         Top             =   330
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "预约序号"
         Height          =   195
         Left            =   3480
         TabIndex        =   5
         Top             =   330
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "姓名"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   330
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "打开"
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   7560
      Width           =   1100
   End
   Begin VB.Menu menu_MouseR 
      Caption         =   "右键菜单"
      Visible         =   0   'False
      Begin VB.Menu menu_OpenSchedule 
         Caption         =   "打开检查预约"
      End
      Begin VB.Menu menu_PrintSchdule 
         Caption         =   "打印预约单"
      End
   End
End
Attribute VB_Name = "frmSchQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrDeptIDs As String   '科室ID串
Private mlngScheduleID As Long  '预约ID
Private mlngOrderID As Long     '医嘱ID

Public Sub ZlShowMe(strDeptIDs As String, frmParent As Object)
'------------------------------------------------
'功能：打开窗口
'参数： strDeptIDs -- 科室ID串
'       frmParent -- 父窗体
'返回：无
'------------------------------------------------
    On Error GoTo err
    
    mstrDeptIDs = strDeptIDs
    
    Call LoadData
    Call cmdQuery_Click
    
    Me.Show 1, frmParent
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub chkDatePeriod_Click()
    If chkDatePeriod.value = 1 Then
        dpDateStart.Enabled = True
        dpDateEnd.Enabled = True
        cboDatePeriod.Enabled = False
    Else
        dpDateStart.Enabled = False
        dpDateEnd.Enabled = False
        cboDatePeriod.Enabled = True
    End If
End Sub

Private Sub cmdClear_Click()
    txtName.Text = ""
    txtSchNumber.Text = ""
    txtCheckNo.Text = ""
    cboSchDevice.ListIndex = 0
    txtOutPatientNo.Text = ""
    txtInPatientNo.Text = ""
    cboDatePeriod.Enabled = True
    cboDatePeriod.ListIndex = 0
    chkDatePeriod.value = 0
    dpDateStart.Enabled = False
    dpDateEnd.Enabled = False
End Sub

Private Sub cmdOpen_Click()
    frmSchSchedule.ZlShowMe mlngOrderID, mstrDeptIDs, Me
    Call QuerySchInfo
End Sub

Private Sub cmdQuery_Click()
    Call QuerySchInfo
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Public Sub QuerySchInfo()
'------------------------------------------------
'功能：查询预约清空
'参数：
'返回：无
'------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim dtStart As Date
    Dim dtEnd As Date
    Dim i As Integer
    
    On Error GoTo err
    
    If chkDatePeriod.value = 1 Then
        dtStart = Format(dpDateStart.value, "YYYY-MM-DD") & " 00:00:00"
        dtEnd = Format(dpDateEnd.value, "YYYY-MM-DD") & " 23:59:59"
    Else
        Select Case cboDatePeriod.ItemData(cboDatePeriod.ListIndex)
            Case 1  '今天
                dtStart = Format(Now, "YYYY-MM-DD") & " 00:00:00"
                dtEnd = Format(Now, "YYYY-MM-DD") & " 23:59:59"
            Case 2  '明天
                dtStart = Format(Now + 1, "YYYY-MM-DD") & " 00:00:00"
                dtEnd = Format(Now + 1, "YYYY-MM-DD") & " 23:59:59"
            Case 3  '今天和明天
                dtStart = Format(Now, "YYYY-MM-DD") & " 00:00:00"
                dtEnd = Format(Now + 1, "YYYY-MM-DD") & " 23:59:59"
            Case 4  '最近三天
                dtStart = Format(Now, "YYYY-MM-DD") & " 00:00:00"
                dtEnd = Format(Now + 2, "YYYY-MM-DD") & " 23:59:59"
            Case 5  '最近一周
                dtStart = Format(Now, "YYYY-MM-DD") & " 00:00:00"
                dtEnd = Format(Now + 7, "YYYY-MM-DD") & " 23:59:59"
            Case 6  '最近两周
                dtStart = Format(Now, "YYYY-MM-DD") & " 00:00:00"
                dtEnd = Format(Now + 14, "YYYY-MM-DD") & " 23:59:59"
            Case 7  '最近一月
                dtStart = Format(Now, "YYYY-MM-DD") & " 00:00:00"
                dtEnd = Format(Now + 30, "YYYY-MM-DD") & " 23:59:59"
            Case 8  '最近两月
                dtStart = Format(Now, "YYYY-MM-DD") & " 00:00:00"
                dtEnd = Format(Now + 61, "YYYY-MM-DD") & " 23:59:59"
            Case 9  '最近三月
                dtStart = Format(Now, "YYYY-MM-DD") & " 00:00:00"
                dtEnd = Format(Now + 92, "YYYY-MM-DD") & " 23:59:59"
        End Select
    End If
    
    strSQL = " select distinct  a.ID, a.序号,a.预约日期,b.姓名,b.性别,b.年龄,b.医嘱内容, c.检查号,c.影像类别, " _
        & " e.设备名称,d.门诊号,d.住院号,a.医嘱ID,a.预约开始时间,a.预约结束时间 from 影像预约记录 a , " _
        & " 病人医嘱记录 b,影像检查记录 c ,病人信息 d ,影像预约设备 e Where a.医嘱ID = b.ID " _
        & " And c.医嘱ID = a.医嘱ID And d.病人ID = b.病人ID And a.预约设备id = e.id and " _
        & " a.预约日期 between [1] and [2] and c.执行科室id in (" & mstrDeptIDs & ") "
    If txtName.Text <> "" Then
        strSQL = strSQL & " and b.姓名=[3]"
    End If
    If txtSchNumber.Text <> "" Then
        strSQL = strSQL & " and a.序号=[4]"
    End If
    If txtCheckNo.Text <> "" Then
        strSQL = strSQL & " and c.检查号=[5]"
    End If
    If cboSchDevice.ListIndex <> 0 Then
        strSQL = strSQL & " and e.设备名称=[6]"
    End If
    If Val(txtOutPatientNo.Text) <> 0 Then
        strSQL = strSQL & " and d.门诊号=[7]"
    End If
    If Val(txtInPatientNo.Text) <> 0 Then
        strSQL = strSQL & " and d.住院号=[8]"
    End If
    
    strSQL = strSQL & " order by 预约日期, 序号"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询预约情况", dtStart, _
        dtEnd, txtName.Text, txtSchNumber.Text, txtCheckNo.Text, _
        cboSchDevice.Text, Val(txtOutPatientNo.Text), Val(txtInPatientNo.Text))
    
    '填写查询结果
    With vsfSchedule
        .Clear
        .Cols = 14
        .Rows = rsTemp.RecordCount + 1
        .FixedRows = 1
        .FixedCols = 0
        .AllowUserResizing = flexResizeColumns
        .SelectionMode = flexSelectionByRow
        .Editable = flexEDNone
        .ScrollBars = flexScrollBarBoth
        .ExplorerBar = flexExSort
        .CellAlignment = flexAlignLeftCenter
'        .Cell(flexcpAlignment, 0, 0, 0, 2) = flexAlignCenterCenter
        .ExtendLastCol = True
        .RowHeight(0) = 350
        
        .Sort = flexSortStringAscending
        
        .ColWidthMin = 1200
        .ColWidth(7) = 1800
        
        .TextMatrix(0, 0) = "ID"
        .TextMatrix(0, 1) = "姓名"
        .TextMatrix(0, 2) = "预约序号"
        .TextMatrix(0, 3) = "检查号"
        .TextMatrix(0, 4) = "预约日期"
        .TextMatrix(0, 5) = "开始时间"
        .TextMatrix(0, 6) = "结束时间"
        .TextMatrix(0, 7) = "医嘱内容"
        .TextMatrix(0, 8) = "性别"
        .TextMatrix(0, 9) = "年龄"
        .TextMatrix(0, 10) = "设备名称"
        .TextMatrix(0, 11) = "医嘱ID"
        .TextMatrix(0, 12) = "门诊号"
        .TextMatrix(0, 13) = "住院号"
        
        
        '从数据库加载数据
        For i = 1 To rsTemp.RecordCount
            .TextMatrix(i, 0) = rsTemp!ID
            .TextMatrix(i, 1) = rsTemp!姓名
            .TextMatrix(i, 2) = rsTemp!序号
            .TextMatrix(i, 3) = rsTemp!检查号
            .TextMatrix(i, 4) = rsTemp!预约日期
            .TextMatrix(i, 5) = Format(rsTemp!预约开始时间, "HH:MM")
            .TextMatrix(i, 6) = Format(rsTemp!预约结束时间, "HH:MM")
            .TextMatrix(i, 7) = rsTemp!医嘱内容
            .TextMatrix(i, 8) = rsTemp!性别
            .TextMatrix(i, 9) = rsTemp!年龄
            .TextMatrix(i, 10) = rsTemp!设备名称
            .TextMatrix(i, 11) = rsTemp!医嘱ID
            .TextMatrix(i, 12) = nvl(rsTemp!门诊号)
            .TextMatrix(i, 13) = nvl(rsTemp!住院号)
            rsTemp.MoveNext
        Next i
    
        '隐藏后台数据
        .ColHidden(0) = True    '预约ID
        .ColHidden(11) = True   '医嘱ID
        
        '选择第一行
        If .Rows > 1 Then
            Call .Select(1, 1)
            mlngScheduleID = Val(.TextMatrix(1, 0))
            mlngOrderID = .TextMatrix(1, 11)
        Else
            mlngScheduleID = 0
            mlngOrderID = 0
        End If
    End With
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub LoadData()
'------------------------------------------------
'功能：初始化数据
'参数：
'返回：无
'------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo err
    
    '加载本科室的预约设备
    strSQL = "select ID,设备名称,影像设备号,影像类别,设备说明 from 影像预约设备 where 科室ID in (" & mstrDeptIDs & ") and 是否启用=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询预约设备")
    
    cboSchDevice.Clear
    cboSchDevice.AddItem "全部"
    cboSchDevice.ItemData(cboSchDevice.NewIndex) = 0
    Do Until rsTemp.EOF
        cboSchDevice.AddItem rsTemp!设备名称
        cboSchDevice.ItemData(cboSchDevice.NewIndex) = rsTemp!ID
        rsTemp.MoveNext
    Loop
    If cboSchDevice.ListCount > 0 Then
        cboSchDevice.ListIndex = 0
    End If
    
    '设置预约时间
    cboDatePeriod.Clear
    cboDatePeriod.AddItem "今天"
    cboDatePeriod.ItemData(cboDatePeriod.NewIndex) = 1
    
    cboDatePeriod.AddItem "明天"
    cboDatePeriod.ItemData(cboDatePeriod.NewIndex) = 2
    
    cboDatePeriod.AddItem "今天和明天"
    cboDatePeriod.ItemData(cboDatePeriod.NewIndex) = 3
    
    cboDatePeriod.AddItem "最近三天"
    cboDatePeriod.ItemData(cboDatePeriod.NewIndex) = 4
    
    cboDatePeriod.AddItem "最近一周"
    cboDatePeriod.ItemData(cboDatePeriod.NewIndex) = 5
    
    cboDatePeriod.AddItem "最近两周"
    cboDatePeriod.ItemData(cboDatePeriod.NewIndex) = 6
    
    cboDatePeriod.AddItem "最近一月"
    cboDatePeriod.ItemData(cboDatePeriod.NewIndex) = 7
    
    cboDatePeriod.AddItem "最近两月"
    cboDatePeriod.ItemData(cboDatePeriod.NewIndex) = 8
    
    cboDatePeriod.AddItem "最近三月"
    cboDatePeriod.ItemData(cboDatePeriod.NewIndex) = 9
    
    cboDatePeriod.ListIndex = 0
    
    dpDateStart = Now
    dpDateEnd = Now
    
    Call cmdClear_Click
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub menu_OpenSchedule_Click()
    Call cmdOpen_Click
End Sub

Private Sub menu_PrintSchdule_Click()
    Call PrintSchedule
End Sub

Private Sub vsfSchedule_Click()
    If vsfSchedule.Rows > 1 Then
        mlngOrderID = vsfSchedule.TextMatrix(vsfSchedule.RowSel, 11)
    End If
End Sub

Private Sub vsfSchedule_DblClick()
    Call cmdOpen_Click
End Sub

Private Sub vsfSchedule_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call PopupMenu(menu_MouseR)
    End If
End Sub

Private Sub PrintSchedule()
'------------------------------------------------
'功能：打印当前预约单
'参数：
'返回：无
'------------------------------------------------
    On Error GoTo err
    
    If mlngOrderID <> 0 Then
        If ReportOpen(gcnOracle, 100, "ZL1_Inside_1290_01", Me, "医嘱ID=" & mlngOrderID) = False Then
            Call MsgBox("报表“ZL1_Inside_1290_01”打开不成功，请联系管理员修正此报表。", vbInformation, "检查预约提示")
        End If
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub
