VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~2.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~3.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "CODEJO~1.OCX"
Begin VB.Form frmQCLJAverage 
   Caption         =   "均值LJ质控查询"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11490
   Icon            =   "frmQCLJAverage.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   11490
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picRecord 
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   0  'None
      Height          =   6750
      Left            =   75
      ScaleHeight     =   6750
      ScaleWidth      =   2445
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   540
      Width           =   2445
      Begin VB.CommandButton cmd刷新 
         Height          =   600
         Left            =   2085
         Picture         =   "frmQCLJAverage.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   90
         Width           =   330
      End
      Begin MSComCtl2.DTPicker dtp日期 
         Height          =   300
         Index           =   0
         Left            =   435
         TabIndex        =   8
         Top             =   75
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   206766083
         CurrentDate     =   39110
      End
      Begin MSComCtl2.DTPicker dtp日期 
         Height          =   300
         Index           =   1
         Left            =   435
         TabIndex        =   9
         Top             =   390
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   206766083
         CurrentDate     =   39110
      End
      Begin VSFlex8Ctl.VSFlexGrid vfgItem 
         Height          =   4830
         Left            =   45
         TabIndex        =   10
         Top             =   735
         Width           =   2445
         _cx             =   4313
         _cy             =   8520
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   14737632
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16635590
         ForeColorSel    =   -2147483640
         BackColorBkg    =   14737632
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   -1  'True
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
      Begin VB.Label lbl日期 
         BackStyle       =   0  'Transparent
         Caption         =   "日期"
         Height          =   180
         Index           =   0
         Left            =   45
         TabIndex        =   12
         Top             =   135
         Width           =   405
      End
      Begin VB.Label lbl日期 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   180
         Index           =   1
         Left            =   210
         TabIndex        =   11
         Top             =   420
         Width           =   180
      End
   End
   Begin VB.ComboBox cbo科室 
      Height          =   300
      Left            =   2070
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   90
      Width           =   1845
   End
   Begin VB.ComboBox cbo仪器 
      Height          =   300
      Left            =   4905
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   75
      Width           =   2115
   End
   Begin VB.PictureBox picCharts 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4395
      Left            =   4890
      ScaleHeight     =   4395
      ScaleWidth      =   6510
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   495
      Width           =   6510
      Begin XtremeSuiteControls.TabControl tbcCharts 
         Height          =   3975
         Left            =   150
         TabIndex        =   2
         Top             =   165
         Width           =   6105
         _Version        =   589884
         _ExtentX        =   10769
         _ExtentY        =   7011
         _StockProps     =   64
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7605
      Width           =   11490
      _ExtentX        =   20267
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmQCLJAverage.frx":6BDC
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15187
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin C1Chart2D8.Chart2D chtCopy 
      Height          =   435
      Left            =   1260
      TabIndex        =   5
      Top             =   60
      Visible         =   0   'False
      Width           =   765
      _Version        =   524288
      _Revision       =   7
      _ExtentX        =   1349
      _ExtentY        =   767
      _StockProps     =   0
      ControlProperties=   "frmQCLJAverage.frx":746E
   End
   Begin VB.PictureBox picData 
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   3840
      ScaleHeight     =   1815
      ScaleWidth      =   5280
      TabIndex        =   13
      Top             =   5040
      Visible         =   0   'False
      Width           =   5280
      Begin VSFlex8Ctl.VSFlexGrid vfgRecord 
         Height          =   1860
         Left            =   600
         TabIndex        =   14
         Top             =   240
         Width           =   4305
         _cx             =   7594
         _cy             =   3281
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   14737632
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16635590
         ForeColorSel    =   -2147483640
         BackColorBkg    =   14737632
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   3
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   -1  'True
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
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   60
      Top             =   15
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmQCLJAverage.frx":7ACD
      Left            =   615
      Top             =   210
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmQCLJAverage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mColL  '均值数据表列
    序号 = 0: 日期: 结果: 实际日期
End Enum

Const conPane_Record = 201
Const conPane_Charts = 202
Const conPane_Report = 203
Const conPane_Data = 204
'-----------------------------------------------------
'窗体变量
'-----------------------------------------------------
Private mstrPrivs As String     '当前使用者权限串
Private mlngListWidth As Long   '列表窗体的设计宽度

Private mfrmChartLJAverage As frmQCChartLJAverage     'LJ控制图窗格

'-----------------------------------------------------
'临时变量
'-----------------------------------------------------
Dim cbrControl As CommandBarControl
Dim cbrCustom As CommandBarControlCustom
Dim cbrMenuBar As CommandBarPopup
Dim cbrToolBar As CommandBar

Dim lngCount As Long
Private mstr项目 As String                  '存储用户当前选中的项目
Private mrsAverage As New ADODB.Recordset      '缓存均值,SD数据
Private mrsData As New ADODB.Recordset      '缓存检验结果均值数据

Private mLastStartDate As Date, mLastEndDate As Date
Private mLastCell As String '焦点离开前的单元格，用于弃用与采信功能

Private Const ID_MENU_MOUSE = 90                                    '右键菜单
Private mlngDeptID As Long
Private mlngMachineID As Long
Private mlngItemID As Long                                          '当前选中的项目ID
Private mLastItemID As Long                                         '上次显示的项目ID，避免重复刷新
'-----------------------------------------------------
'以下为内部公共程序
'-----------------------------------------------------
Private Function zlRefRecord() As Long
    '功能：刷新质控结果记录
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim lngCol As Long, lngRow As Long
    Dim dtStart As Date, dtEnd As Date

    Err = 0: On Error GoTo ErrHand
    If mlngItemID = 0 Then Exit Function
    
    dtStart = Format(Me.dtp日期(0).Value, "yyyy-MM-dd")
    dtEnd = Format(Me.dtp日期(1).Value, "yyyy-MM-dd 23:59:59")
    
    '获取指定时间范围内通过审核的标本 检验项目结果的平均值
    gstrSql = " Select Trunc(a.检验时间) As 日期,Avg(Translate(Zl_To_Number(b.检验结果,0),'>=<+-','00000')) As 结果 " & _
                "From 检验标本记录 A, 检验普通结果 b " & _
                "Where a.审核人 Is Not Null And a.id=b.检验标本ID And b.检验项目id + 0 = [1] And a.检验时间 Between [2] And [3] " & _
                "Group By Trunc(a.检验时间)  order by Trunc(a.检验时间) "
            'Nvl(弃用结果, 0) * -1 +
    Set mrsData = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngItemID, dtStart, dtEnd)
 
    '数据缓存，画质控图时会用到(frmQCChartLJAverage)
     
    With Me.vfgRecord
        .Redraw = flexRDNone
        .Clear
        .FixedCols = 3
        .Cols = .FixedCols
        .ExtendLastCol = False '不自动扩展最后一列的宽度
        .Rows = 4
        
        .ColWidth(0) = 1200
        .TextMatrix(mColL.序号, 0) = ""
        .TextMatrix(mColL.序号, 1) = "靶值": .ColWidth(1) = 500
        .TextMatrix(mColL.序号, 2) = "SD": .ColWidth(2) = 500
        .TextMatrix(mColL.日期, 0) = "日期"
        .TextMatrix(mColL.结果, 0) = "结果"
        .TextMatrix(mColL.实际日期, 0) = "实际日期": .RowHidden(mColL.实际日期) = True
        .ColAlignment(0) = flexAlignLeftCenter
        
        '将检验结果均值填充到结果列表中
        Do Until mrsData.EOF
            .Cols = .Cols + 1
            .TextMatrix(mColL.序号, mrsData.AbsolutePosition + 2) = mrsData.AbsolutePosition
            .TextMatrix(mColL.日期, mrsData.AbsolutePosition + 2) = Format(Nvl(mrsData!日期), "yy-MM-dd")
            .TextMatrix(mColL.结果, mrsData.AbsolutePosition + 2) = Round(Nvl(mrsData!结果, 0), 2)
            .TextMatrix(mColL.实际日期, mrsData.AbsolutePage + 2) = Nvl(mrsData!日期)
            mrsData.MoveNext
        Loop
        
        '填写均值、SD、CV
        'translate(b.检验结果,'>=<-+','00000')
        gstrSql = "Select Round(Avg(结果), 2) As 均值, Round(Stddev(结果), 3) As Sd " & _
                  "From (Select Trunc(a.检验时间) As 日期,Avg(Translate(Zl_To_Number(b.检验结果,0),'>=<+-','00000')) As 结果 " & _
                        "From 检验标本记录 A, 检验普通结果 b Where a.审核人 Is Not Null And a.id=b.检验标本ID " & _
                        "And b.检验项目id + 0 = [1] And a.检验时间 Between [2] And [3] " & _
                  "Group By Trunc(a.检验时间))"
        Set mrsAverage = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngItemID, dtStart, dtEnd)
        
        '数据缓存，画质控图时会用到(frmQCChartLJAverage)
        
        If Not mrsAverage.EOF Then
            .TextMatrix(mColL.结果, 1) = Val("" & mrsAverage!均值)
            .TextMatrix(mColL.结果, 2) = Val("" & mrsAverage!SD)
        End If
 
        If .Cols > .FixedCols Then
            .Cell(flexcpAlignment, mColL.序号, .FixedCols, mColL.日期, .Cols - 1) = flexAlignCenterCenter
            .AutoSize 0, .Cols - 1
        End If
        .Redraw = flexRDDirect
        If .Cols > .FixedCols Then .COL = .FixedCols
    End With
    
    zlRefRecord = Me.vfgRecord.Cols - Me.vfgRecord.FixedCols
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    zlRefRecord = 0
End Function

Private Sub zlRefOthers()
    '功能：根据显示属性，刷新除质控记录外图形和报告
    Dim strLists As String, intLists As Integer
    Dim lngItemID As Long, strFromDate As String, strToDate As String
    Dim str仪器名称 As String

    If mlngItemID = 0 Then Exit Sub

    If mlngItemID = mLastItemID Then Exit Sub
    If mlngItemID = -1 Then mlngItemID = mLastItemID

    mLastItemID = mlngItemID
    lngItemID = mlngItemID
    strFromDate = Format(Me.dtp日期(0).Value, "yyyy-MM-dd")
    strToDate = Format(Me.dtp日期(1).Value, "yyyy-MM-dd")
    str仪器名称 = Trim(Left$(Me.cbo仪器.Text, 30))
    
    '获得当前选择的质控图
    Dim intSelTab As Integer
    For lngCount = 0 To Me.tbcCharts.ItemCount - 1
        If Me.tbcCharts.Item(lngCount).Selected Then intSelTab = lngCount: Exit For
    Next

    If Me.tbcCharts.Item(intSelTab).Visible = False Then Me.tbcCharts.Item(0).Selected = True
    If Me.tbcCharts.Item(0).Selected Then Call mfrmChartLJAverage.zlRefresh(lngItemID, str仪器名称, strFromDate, strToDate, mrsData, mrsAverage)
End Sub

Private Sub cbo科室_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim lngMachineID As Long                '仪器ID
    
    On Error GoTo errH
    
    lngMachineID = mlngMachineID
    
    If Me.cbo科室.ListCount <= 0 Then Exit Sub
    
    If InStr(1, mstrPrivs, "所有科室") > 0 Then
        gstrSql = " Select Distinct D.ID, D.编码, D.名称, D.质控水平数 From 检验仪器 D " & _
                    " Where  Nvl(D.微生物, 0) <> 1 and d.使用小组id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(Me.cbo科室.ItemData(Me.cbo科室.ListIndex)))
    Else
        gstrSql = "Select Distinct D.ID, D.编码, D.名称, D.质控水平数" & vbNewLine & _
                    " From 检验仪器 D " & vbNewLine & _
                    " Where Nvl(D.微生物, 0) <> 1 And D.使用小组id = [2] And" & vbNewLine & _
                    "      D.ID In (Select Distinct D.ID" & vbNewLine & _
                    "               From 检验小组成员 A, 检验小组 B, 检验小组仪器 C, 检验仪器 D" & vbNewLine & _
                    "               Where A.小组id = B.ID And B.ID = C.小组id　and 人员id = [1] And C.仪器id = D.ID)"

        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(UserInfo.ID), CLng(Me.cbo科室.ItemData(Me.cbo科室.ListIndex)))
    End If
    
    With rsTemp
        Me.cbo仪器.Clear
        
        Do While Not .EOF
            Me.cbo仪器.AddItem !名称 & Space(200) & !质控水平数
            Me.cbo仪器.ItemData(Me.cbo仪器.NewIndex) = !ID
            If !ID = lngMachineID Then
                Me.cbo仪器.ListIndex = Me.cbo仪器.NewIndex
            End If
            .MoveNext
        Loop
        If Me.cbo仪器.ListCount > 0 And cbo仪器.ListIndex = -1 Then
            Me.cbo仪器.ListIndex = 0
        End If
    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'-----------------------------------------------------
'以下为控件事件处理
'-----------------------------------------------------
Private Sub cbo仪器_Click()
    Dim lngItemID As Long   '项目ID

    Dim rsTemp As New ADODB.Recordset
    
    mlngItemID = Val(zlDatabase.GetPara("项目", glngSys, 1209, 0))
    

    If Me.cbo仪器.ListIndex = -1 Then Exit Sub
    Me.cbo仪器.Tag = Right(Me.cbo仪器.Text, 1)
    
    Err = 0: On Error GoTo ErrHand
    '获取仪器相关的所有  定量项目
    gstrSql = "Select Distinct b.ID, b.编码, b.英文名, b.中文名" & vbNewLine & _
                " From 检验仪器项目 a, 诊治所见项目 b, 检验项目 c" & vbNewLine & _
                " Where a.项目id = b.ID And a.项目id = c.诊治项目id And c.结果类型 = 1 And 仪器ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(Me.cbo仪器.ItemData(Me.cbo仪器.ListIndex)))
    
    If rsTemp.RecordCount <= 0 Then MsgBox "尚未完成仪器检验项目设置！", vbInformation, gstrSysName: vfgItem.Clear:  Exit Sub
    
    With Me.vfgItem
        .FixedRows = 1
        .SelectionMode = flexSelectionByRow
        
        Set .DataSource = rsTemp
        .ColWidth(0) = 0
        .ColWidth(1) = 500
        .ColWidth(2) = 600
        .ColWidth(3) = 600
        .ColHidden(0) = True
        .AutoSize 1, 2
        .ColWidth(1) = 20
    End With
    Call vfgItem_RowColChange
        
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim rsTmp As New ADODB.Recordset
    
    '------------------------------------
    Select Case Control.ID
    
    Case conMenu_File_PrintSet
        Select Case Me.tbcCharts.Selected.Index
        Case 0: ReportPrintSet gcnOracle, glngSys, "ZL1_INSIDE_1208_8", Me
        End Select
    Case conMenu_File_Preview
        If Not Me.vfgRecord.Cols > Me.vfgRecord.FixedCols Then Exit Sub
        Select Case Me.tbcCharts.Selected.Index
        Case 0: Call mfrmChartLJAverage.ChartPrint: Call PrintQC(False)
        End Select
    Case conMenu_File_Print
        If Not Me.vfgRecord.Cols > Me.vfgRecord.FixedCols Then Exit Sub
        Select Case Me.tbcCharts.Selected.Index
        Case 0: Call mfrmChartLJAverage.ChartPrint: Call PrintQC(True)
        End Select
    Case conMenu_Edit_Save
        If Not Me.vfgRecord.Cols > Me.vfgRecord.FixedCols Then Exit Sub
        Select Case Me.tbcCharts.Selected.Index
        Case 0: Call mfrmChartLJAverage.ChartSaveAs
        End Select
    Case conMenu_Edit_MarkMap
        Select Case Me.tbcCharts.Selected.Index
        Case 0: Call mfrmChartLJAverage.ChartCopy
        End Select
    Case conMenu_File_Exit: Unload Me
    
    Case conMenu_View_ToolBar_Button
        Me.cbsThis(2).Visible = Not Me.cbsThis(2).Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text
        For Each cbrControl In Me.cbsThis(2).Controls
            cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Size
        Me.cbsThis.Options.LargeIcons = Not Me.cbsThis.Options.LargeIcons
        Me.cbsThis.RecalcLayout
    Case conMenu_View_StatusBar
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_Refresh
        mLastItemID = 0
        Call RefreshData
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case conMenu_Tool_Reference_1
        '上
        Call ItemMoveUpDown(1)
    Case conMenu_Tool_Reference_2
        '下
        Call ItemMoveUpDown(2)
    Case Else
        If Control.ID < conMenu_ReportPopup * 100# + 1 Or Control.ID > conMenu_ReportPopup * 100# + 99 Then Exit Sub
        Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me)
    End Select
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_BatPrint, conMenu_Edit_Save, conMenu_Edit_MarkMap
        Control.Enabled = ((Me.vfgRecord.Cols > Me.vfgRecord.FixedCols) And mrsAverage.RecordCount <> 0)
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
    End Select
End Sub

Private Sub cmd刷新_Click()
    mLastItemID = 0
    Call RefreshData
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_Record
        Item.Handle = Me.picRecord.hWnd
    Case conPane_Charts
        Item.Handle = Me.picCharts.hWnd
    Case conPane_Data
        Item.Handle = Me.picData.hWnd
    End Select
End Sub

Private Sub dkpMan_RClick(ByVal Pane As XtremeDockingPane.IPane)
    If Pane.ID = conPane_Data Then
        Me.picData.Visible = True
    End If
End Sub

Private Sub RefreshData()
    Dim objControl As CommandBarControl
    Dim intRow As Integer
    Dim rsTemp As New ADODB.Recordset

    If mlngItemID = 0 Then Exit Sub

    Err = 0: On Error GoTo ErrHand
    
    If Me.dtp日期(1).Value < Me.dtp日期(0).Value Then
        MsgBox "结束日期不能大于开始日期！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    mLastStartDate = Format(dtp日期(0).Value, "yyyy-MM-dd")
    mLastEndDate = Format(dtp日期(1).Value, "yyyy-MM-dd")

    '刷新结果数据
    Call zlRefRecord
    Call zlRefOthers
    Call picRecord_Resize
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim lngDeptID As Long  '科室ID
    '-----------------------------------------------------
    mlngListWidth = Me.picRecord.Width
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, False)
    
    lngDeptID = mlngDeptID
    
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Me.cbsThis.Icons = zlCommFun.GetPubIcons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsThis.EnableCustomization False
    
    '-----------------------------------------------------
    '菜单定义
    Me.cbsThis.ActiveMenuBar.Title = "菜单"
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览控制图")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印控制图(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "另存控制图(&S)...")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_MarkMap, "复制控制图(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): cbrControl.BeginGroup = True
    End With
    
    
    Set cbrControl = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlLabel, 0, "科室")
    cbrControl.Flags = xtpFlagRightAlign
    Set cbrCustom = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlCustom, 0, "科室")
    cbrCustom.Handle = Me.cbo科室.hWnd: cbrCustom.Flags = xtpFlagRightAlign
    Set cbrControl = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlLabel, 0, "仪器")
    cbrControl.Flags = xtpFlagRightAlign
    Set cbrCustom = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlCustom, 0, "仪器")
    cbrCustom.Handle = Me.cbo仪器.hWnd: cbrCustom.Flags = xtpFlagRightAlign
    cbrMenuBar.Visible = False
    
    '快键绑定
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("C"), conMenu_Edit_MarkMap
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    
        .Add 0, VK_UP, conMenu_Tool_Reference_1
        .Add 0, VK_DOWN, conMenu_Tool_Reference_2
    
    End With
    
    '设置不常用菜单
    With Me.cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_Edit_MarkMap
      '  .AddHiddenCommand conMenu_View_Refresh
    End With
    
    '-----------------------------------------------------
    '工具栏定义
    Set cbrToolBar = Me.cbsThis.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "另存为")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    '-----------------------------------------------------
    '设置停靠窗格
    Dim panThis As Pane, panChild As Pane, panSub As Pane
    
    With Me.dkpMan
        Set panThis = .CreatePane(conPane_Record, 200, 400, DockLeftOf, Nothing)
        panThis.Title = "均值结果表"
        panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        Set panThis = .CreatePane(conPane_Charts, 400, 500, DockRightOf, Nothing)
        panThis.Title = "均值LJ质控图"
        panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        
        Set panChild = .CreatePane(conPane_Data, 400, 100, DockBottomOf, panThis)
        panChild.Title = "检验结果"
        panChild.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable

        panChild.Select
        
        .SetCommandBars Me.cbsThis
        .Options.ThemedFloatingFrames = True
        .Options.HideClient = True
    End With

    '-----------------------------------------------------
    '设置表格附加窗格
    Dim tbiThis As TabControlItem
    Set mfrmChartLJAverage = New frmQCChartLJAverage

    With Me.tbcCharts
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        Set tbiThis = .InsertItem(0, mfrmChartLJAverage.Caption, mfrmChartLJAverage.hWnd, 0)
        
    End With
    
    '-----------------------------------------------------
    '界面恢复
    Call RestoreWinState(Me, App.ProductName)
    
    '-----------------------------------------------------
    '装入基本数据
    Dim rsTemp As New ADODB.Recordset
    
    Me.dtp日期(1).Value = zlDatabase.Currentdate: Me.dtp日期(0).Value = CDate(Format(Me.dtp日期(1).Value, "yyyy-MM") & "-01")
    Err = 0: On Error GoTo ErrHand
    
    If InStr(1, mstrPrivs, "所有科室") > 0 Then
        gstrSql = " Select Distinct b.Id, b.编码 , b.名称 As 科室 From 检验仪器 a ,部门表 b Where a.使用小组ID = b.ID "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, gstrSysName)
        
    Else

        gstrSql = "Select Distinct B.ID, B.编码, B.名称 As 科室" & vbNewLine & _
                " From 检验仪器 A, 部门表 B " & vbNewLine & _
                " Where A.使用小组id = B.ID And" & vbNewLine & _
                "      A.使用小组id In (Select Distinct D.使用小组id" & vbNewLine & _
                "                   From 检验小组成员 A, 检验小组 B, 检验小组仪器 C, 检验仪器 D" & vbNewLine & _
                "                   Where A.小组id = B.ID And B.ID = C.小组id　and 人员id = [1] And C.仪器id = D.ID)"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, gstrSysName, UserInfo.ID)
    End If
    
    Me.cbo科室.Clear
    Do Until rsTemp.EOF
        Me.cbo科室.AddItem rsTemp("编码") & "-" & rsTemp("科室")
        Me.cbo科室.ItemData(Me.cbo科室.NewIndex) = rsTemp("Id")
        If rsTemp("ID") = lngDeptID Then
            Me.cbo科室.ListIndex = Me.cbo科室.NewIndex
        End If
        rsTemp.MoveNext
    Loop
    If Me.cbo科室.ListCount = 0 Then MsgBox "尚未完成仪器使用小组的设置！", vbInformation, gstrSysName: Unload Me: Exit Sub
    If cbo科室.ListIndex = -1 Then
        Me.cbo科室.ListIndex = 0
    End If
    If Me.cbo科室.ListCount = 1 Then Me.cbo科室.Enabled = False
    
    mLastStartDate = CDate(0)
    mLastEndDate = CDate(0)
    
    
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    Dim panThis As Pane
    If Me.WindowState = vbMinimized Then Exit Sub
    Set panThis = Me.dkpMan.FindPane(conPane_Record)
    panThis.MinTrackSize.SetSize mlngListWidth / Screen.TwipsPerPixelX, panThis.MinTrackSize.Height
    panThis.MaxTrackSize.SetSize mlngListWidth / Screen.TwipsPerPixelX, panThis.MaxTrackSize.Height
    Me.dkpMan.RecalcLayout
    Me.dkpMan.NormalizeSplitters
    panThis.MinTrackSize.SetSize mlngListWidth / Screen.TwipsPerPixelX, panThis.MinTrackSize.Height
    panThis.MaxTrackSize.SetSize Screen.Width / Screen.TwipsPerPixelX, panThis.MaxTrackSize.Height
    Me.dkpMan.RecalcLayout
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload mfrmChartLJAverage
    Set mfrmChartLJAverage = Nothing
    
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub picCharts_Resize()
    Err = 0: On Error Resume Next
    With Me.tbcCharts
        .Left = Me.picCharts.ScaleLeft: .Width = Me.picCharts.ScaleWidth - .Left
        .Top = Me.picCharts.ScaleTop: .Height = Me.picCharts.ScaleHeight - .Top
    End With
End Sub

Private Sub picData_Resize()
    Err = 0: On Error Resume Next
    '数据列表
    With Me.vfgRecord
        .Left = Me.picData.ScaleLeft: .Width = Me.picData.ScaleWidth - .Left
        .Top = Me.picData.ScaleTop
        .Height = Me.picData.ScaleHeight - .Top
    End With
End Sub

Private Sub picRecord_Resize()
    Err = 0: On Error Resume Next
    
    Me.cmd刷新.Left = Me.picRecord.ScaleWidth - Me.cmd刷新.Width - 15
    Me.dtp日期(1).Width = Me.picRecord.ScaleWidth - Me.cmd刷新.Width - 15 - Me.dtp日期(1).Left - 15
    Me.dtp日期(0).Width = Me.dtp日期(1).Width

    '项目列表
    With Me.vfgItem
        .Left = Me.picRecord.ScaleLeft: .Width = Me.picRecord.ScaleWidth - .Left
        .Height = Me.picRecord.ScaleHeight - .Top
    End With
    
End Sub

Private Sub tbcCharts_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    mlngItemID = -1 '强制刷新
    If Me.Visible Then Call zlRefOthers
End Sub

Private Sub vfgItem_RowColChange()

    
    If mLastStartDate <> CDate(0) And mLastEndDate <> CDate(0) Then
        Me.dtp日期(0) = mLastStartDate
        Me.dtp日期(1) = mLastEndDate
    
    Else
        Me.dtp日期(0) = CDate(Format(Now, "yyyy-MM-01"))
        Me.dtp日期(1) = CDate(Format(Now, "yyyy-MM-dd"))
    End If
    With Me.vfgItem
        If .Row >= .FixedRows Then
            mstr项目 = Trim(.TextMatrix(.Row, 3)) & "/" & Trim(.TextMatrix(.Row, 2))
            If mlngItemID <> Val(.TextMatrix(.Row, 0)) Then
                mlngItemID = Val(.TextMatrix(.Row, 0))
                
                Call RefreshData
            End If
        End If
    End With
    
End Sub

Private Sub vfgRecord_EnterCell()
    With vfgRecord
        mLastCell = .Row & "," & .COL
    End With
End Sub

Private Sub vfgRecord_LeaveCell()
    With vfgRecord
        mLastCell = .Row & "," & .COL
    End With
End Sub

Private Sub vfgRecord_RowColChange()
    With vfgRecord
        mLastCell = .Row & "," & .COL
    End With
End Sub

Private Sub PrintQC(blnPrintMode As Boolean)
    '打印或预览质控图
    '参数           intPrintMode =1 打印 =2 预览
    '               intPrintType 0=LJ 1=FQ 2=ZS 3=YD 4=CS 5=MN
    
    Dim rsTmp As New ADODB.Recordset
    Dim strPrintType As String                  '对应的单据
    Dim strQCID As String                       '质控品ID可能会是以","分隔的多个ID
    Dim lngQCID As Long                         '单个质控品ID
    Dim lngItemID As String                     '项目ID
    Dim lngMachine As Long                      '仪器ID
    Dim intLoop As Integer
    
    
    On Error GoTo errH
    
    strPrintType = "ZL1_INSIDE_1208_8"
    gstrSql = "Select b.w, b.h " & vbNewLine & _
                " From Zlreports a, Zlrptitems b" & vbNewLine & _
                " Where a.Id = b.报表id And a.编号 = [1] And b.名称 = '均值LJ图'"
                
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strPrintType)
    '没有找到时退出
    If rsTmp.EOF Then
        MsgBox "在单据定义中没有定义<均值LJ图>,请在单据中定义一个名为<均值LJ图>的图像框!", vbQuestion, Me.Caption
        Exit Sub
    End If
    
    If Dir(App.path & "\QCLJAverage_Tmp") <> "" Then
        With Me.chtCopy
            .Load App.path & "\QCLJAverage_Tmp"
            Kill App.path & "\QCLJAverage_Tmp"
            .Width = Nvl(rsTmp("w"), 1280 * Screen.TwipsPerPixelX)
            .Height = Nvl(rsTmp("h"), 500 * Screen.TwipsPerPixelY)
            .Header.Text = ""
            .ChartLabels.RemoveAll
            .ChartArea.Location.Top = -5
            .ChartArea.Location.Height = .ChartArea.Location.Height + 15
    
            .SaveImageAsJpeg App.path & "\QC_LJAverage" & ".jpg", 1000, False, False, False
        End With
    End If
    
    '得到项目ID
    If mlngItemID = 0 Then Exit Sub
    lngItemID = mlngItemID
    lngMachine = CLng(Me.cbo仪器.ItemData(Me.cbo仪器.ListIndex))
    
    If Dir(App.path & "\QC_LJAverage.jpg") <> "" Then
        Call ReportOpen(gcnOracle, glngSys, strPrintType, Me, _
        "单位=" & gstrUnitName, "开始时间=" & dtp日期(0).Value, "结束时间=" & dtp日期(1).Value, "仪器=" & Left(Trim(cbo仪器.Text), 30), "项目=" & mstr项目, _
        "计算均值=" & Val("" & mrsAverage!均值), "SD=" & Val("" & mrsAverage!SD), "均值LJ图=" & App.path & "\QC_LJAverage.jpg", _
        IIf(blnPrintMode, 2, 1))
    End If
    
    If Dir(App.path & "\QC*.jpg") <> "" Then Kill App.path & "\QC*.jpg"
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ItemMoveUpDown(ByVal intUpDown As Integer)
    '上下键处理
    On Error Resume Next
    With Me.vfgItem
        If intUpDown = 1 Then
            If .Row - 1 > .FixedRows Then .Select .Row - 1, .COL
        Else
            If .Row + 1 < .Rows Then .Select .Row + 1, .COL
        End If
    End With
End Sub

Public Sub ShowMe(frmParent As Object, ByVal strPrivs As String, ByVal lngDeptID As Long, ByVal lngMachineID As Long)
    mstrPrivs = strPrivs
    mlngDeptID = lngDeptID
    mlngMachineID = lngMachineID
    
    Me.Show vbModal, frmParent
End Sub


