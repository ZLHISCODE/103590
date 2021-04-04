VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmClinicPlanChangeHistory 
   Caption         =   "临床出诊安排变动信息"
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11760
   Icon            =   "frmClinicPlanChangeHistory.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   11760
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picButton 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   11760
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   7695
      Width           =   11760
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "退出(&E)"
         Height          =   350
         Left            =   10230
         TabIndex        =   9
         Top             =   180
         Width           =   1100
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "帮助(&H)"
         Height          =   350
         Left            =   450
         TabIndex        =   8
         Top             =   180
         Width           =   1100
      End
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "查询(&F)"
      Height          =   350
      Left            =   5760
      TabIndex        =   4
      Top             =   60
      Width           =   1100
   End
   Begin MSComCtl2.DTPicker dtpStartDate 
      Height          =   300
      Left            =   900
      TabIndex        =   1
      Top             =   90
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   529
      _Version        =   393216
      CalendarTitleBackColor=   -2147483630
      CalendarTitleForeColor=   -2147483634
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   172556291
      CurrentDate     =   42453
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfChangeInfo 
      Height          =   6735
      Left            =   30
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   450
      Width           =   10245
      _cx             =   18071
      _cy             =   11880
      Appearance      =   1
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
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483638
      GridColorFixed  =   -2147483638
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmClinicPlanChangeHistory.frx":6852
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
      Begin VB.PictureBox picImgPlan 
         BorderStyle     =   0  'None
         Height          =   135
         Left            =   30
         Picture         =   "frmClinicPlanChangeHistory.frx":68C7
         ScaleHeight     =   135
         ScaleWidth      =   150
         TabIndex        =   6
         Top             =   60
         Width           =   150
      End
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   300
      Left            =   3330
      TabIndex        =   3
      Top             =   90
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   529
      _Version        =   393216
      CalendarTitleBackColor=   -2147483630
      CalendarTitleForeColor=   -2147483634
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   172556291
      CurrentDate     =   42453
   End
   Begin VB.Label lblTimeRange 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "～"
      Height          =   180
      Left            =   3090
      TabIndex        =   2
      Top             =   150
      Width           =   180
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "查询时间"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   720
   End
End
Attribute VB_Name = "frmClinicPlanChangeHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModule As Long
Private mblnFirst As Boolean

Public Function ShowMe(frmParent As Object, ByVal lngModule As Long) As Boolean
    '程序入口
    mlngModule = lngModule
    Err = 0: On Error Resume Next
    
    Me.Show 1, frmParent
End Function

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    If DateDiff("s", dtpStartDate.Value, dtpEndDate.Value) <= 0 Then
        MsgBox "查询终止时间必须大于开始时间！", vbInformation, gstrSysName
        If dtpEndDate.Visible And dtpEndDate.Enabled Then dtpEndDate.SetFocus
        Exit Sub
    End If
    Call RefreshData
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub dtpEndDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab: Exit Sub
End Sub

Private Sub dtpStartDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab: Exit Sub
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
    If cmdFind.Visible And cmdFind.Enabled Then cmdFind.SetFocus
End Sub

Private Sub Form_Load()
    mblnFirst = True
    '初始化日期，默认显示一个星期
    dtpStartDate.Value = Format(Now - 7, "yyyy-mm-dd hh:mm:ss")
    dtpEndDate.Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
    
    '初始化表格
    Call InitGrid
    Call zl_vsGrid_Para_Restore(mlngModule, vsfChangeInfo, Me.Name, "变动信息")
    
    Call RefreshData
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With vsfChangeInfo
        .Left = 20
        .Top = 450
        .Width = Me.ScaleWidth - .Left * 2
        .Height = Me.ScaleHeight - picButton.Height - .Top - 20
    End With
End Sub

Private Sub InitGrid()
    '初始化表格
    Dim strHead As String, varData As Variant
    Dim strHeadSub As String, varDataSub As Variant
    Dim i As Long, lngCol As Long
    Dim arrDate As Variant
    Dim dtCurDate As Date, intDays As Integer
    
    Err = 0: On Error GoTo errHandler
    With vsfChangeInfo
        .Redraw = False
        .Rows = 2
        
        strHead = ",4,220|号类,4,500|号码,4,500|科室,1,1000|项目,1,0|医生,1,700|变动日期,4,1100|" & _
                "变动原因,1,1200|变动前内容,1,2500|变动后内容,1,2500|" & _
                "登记人,1,700|登记时间,4,1900|审批人,1,700|审批时间,4,1900|取消人,1,700|取消时间,4,1900"
        varData = Split(strHead, "|")
        .Cols = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColAlignment(i) = Split(varData(i), ",")(1)
            .ColWidth(i) = Split(varData(i), ",")(2)
            .ColKey(i) = Split(varData(i), ",")(0)
        Next
        .FixedCols = 1: .FixedRows = 1
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .ExplorerBar = flexExSort
        .HighLight = flexHighlightAlways
        .FocusRect = flexFocusNone
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .GridLines = flexGridFlat
        '.WordWrap = True '允许自动换行
        .RowHeightMin = 350
        
        '列属性设置,用于用户选择显示列
        For i = 0 To .Cols - 1
            'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)|列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
            Select Case i
            Case .ColIndex("号码"), .ColIndex("科室"), .ColIndex("医生")
                .ColData(i) = "1|0"
            End Select
        Next
        .Redraw = True
    End With
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function RefreshData() As Boolean
    Dim strSQL As String, lngRow As Long
    Dim rsData As ADODB.Recordset
    Dim dtStart As Date, dtEnd As Date
    
    Err = 0: On Error GoTo errHandler
    dtStart = dtpStartDate.Value
    dtEnd = dtpEndDate.Value
    
    vsfChangeInfo.Clear 1
    vsfChangeInfo.Rows = 2
    zlCommFun.ShowFlash "正在加载数据，请稍等...", Me
    '限号、限约、诊室调整
    strSQL = "Select Max(记录id) As 记录id, Max(变动原因) As 变动原因, Max(变动前 || Decode(变动性质, 1, 诊室)) As 变动前," & vbNewLine & _
            "        Max(变动后 || Decode(变动性质, 2, 诊室)) As 变动后, Max(登记人) As 登记人, Max(登记时间) As 登记时间, Max(登记人) As 审批人, Max(登记时间) As 审批时间," & vbNewLine & _
            "        Null As 取消人, Null As 取消时间" & vbNewLine & _
            " From (Select m.ID As 变动id, n.变动性质, Max(m.记录id) As 记录id, Max(Decode(m.变动类型, 1, '限号调整', 2, '限约调整', 3, '诊室变动')) As 变动原因," & vbNewLine & _
            "               Max(Decode(m.变动类型, 1, '限号:' || 原数量, 2, '限约:' || 原数量, 3," & vbNewLine & _
            "                           Decode(原分诊方式, 0, '不分诊', 1, '指定诊室:', 2, '动态分诊:', 3, '平均分诊:'))) As 变动前," & vbNewLine & _
            "               Max(Decode(m.变动类型, 1, '限号:' || 现数量, 2, '限约:' || 现数量, 3," & vbNewLine & _
            "                           Decode(现分诊方式, 0, '不分诊', 1, '指定诊室:', 2, '动态分诊:', 3, '平均分诊:'))) As 变动后," & vbNewLine & _
            "               f_List2str(Cast(Collect(n.门诊诊室) As t_Strlist)) As 诊室, Max(m.操作员姓名) As 登记人, Max(m.登记时间) As 登记时间" & vbNewLine & _
            "        From 临床出诊变动记录 M, 临床出诊变动明细 N" & vbNewLine & _
            "        Where m.Id = n.变动id(+) And m.变动类型 In (1, 2, 3) And m.登记时间 Between [1] And [2]" & vbNewLine & _
            "        Group By m.id, n.变动性质)" & vbNewLine & _
            " Group By 变动id"
    '停诊、替诊
    strSQL = strSQL & vbNewLine & _
            " Union All" & vbNewLine & _
            " Select m.记录id, Decode(m.替诊医生姓名, Null, '停诊', '替诊') As 变动原因, '' As 变动前," & vbNewLine & _
            "       Decode(m.替诊医生姓名, Null," & vbNewLine & _
            "               '停诊时间:' || To_Char(n.出诊日期, 'yyyy-mm-dd') || ' ' || To_Char(m.开始时间, 'hh24:mi') || '至' ||" & vbNewLine & _
            "                To_Char(m.终止时间, 'hh24:mi'), n.上班时段 || ',替诊医生:' || m.替诊医生姓名) As 变动后, m.申请人 As 登记人, m.申请时间 As 登记时间," & vbNewLine & _
            "       m.审批人, m.审批时间, m.取消人, m.取消时间" & vbNewLine & _
            " From 临床出诊停诊记录 M, 临床出诊记录 N" & vbNewLine & _
            " Where m.记录id Is Not Null And m.记录id = n.Id And m.审批时间 Is Not Null And m.申请时间 Between [1] And [2]"
    
    strSQL = "Select c.号类, c.号码, d.名称 As 项目, e.名称 As 科室, b.医生姓名 As 医生, To_Char(b.出诊日期, 'yyyy-mm-dd') As 变动日期," & _
            "        a.变动原因, a.变动前, a.变动后, a.登记人, To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间, " & vbNewLine & _
            "        a.登记人 As 审批人, To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 审批时间, " & vbNewLine & _
            "        a.取消人, Decode(a.取消时间, Null, '', To_Char(a.取消时间, 'yyyy-mm-dd hh24:mi:ss')) As 取消时间" & vbNewLine & _
            " From (" & strSQL & ") A, 临床出诊记录 B," & vbNewLine & _
            "      临床出诊号源 C, 收费项目目录 D, 部门表 E" & vbNewLine & _
            " Where a.记录id = b.Id And b.号源id = c.Id And b.项目id = d.Id And b.科室id = e.Id(+)"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, dtStart, dtEnd)
    If rsData Is Nothing Then Exit Function
    If rsData.RecordCount = 0 Then Exit Function
    
    '加载数据
    With vsfChangeInfo
        .Redraw = False
        .Rows = rsData.RecordCount + 1
        lngRow = 1
        Do While Not rsData.EOF
            .TextMatrix(lngRow, .ColIndex("号类")) = Nvl(rsData!号类)
            .TextMatrix(lngRow, .ColIndex("号码")) = Nvl(rsData!号码)
            .TextMatrix(lngRow, .ColIndex("科室")) = Nvl(rsData!科室)
            .TextMatrix(lngRow, .ColIndex("项目")) = Nvl(rsData!项目)
            .TextMatrix(lngRow, .ColIndex("医生")) = Nvl(rsData!医生)
            .TextMatrix(lngRow, .ColIndex("变动日期")) = Nvl(rsData!变动日期)
            .TextMatrix(lngRow, .ColIndex("变动原因")) = Nvl(rsData!变动原因)
            .TextMatrix(lngRow, .ColIndex("变动前内容")) = Nvl(rsData!变动前)
            .TextMatrix(lngRow, .ColIndex("变动后内容")) = Nvl(rsData!变动后)
            .TextMatrix(lngRow, .ColIndex("登记人")) = Nvl(rsData!登记人)
            .TextMatrix(lngRow, .ColIndex("登记时间")) = Nvl(rsData!登记时间)
            .TextMatrix(lngRow, .ColIndex("审批人")) = Nvl(rsData!审批人)
            .TextMatrix(lngRow, .ColIndex("审批时间")) = Nvl(rsData!审批时间)
            .TextMatrix(lngRow, .ColIndex("取消人")) = Nvl(rsData!取消人)
            .TextMatrix(lngRow, .ColIndex("取消时间")) = Nvl(rsData!取消时间)
            lngRow = lngRow + 1
            rsData.MoveNext
        Loop
        .Redraw = True
    End With
    zlCommFun.StopFlash
    RefreshData = True
    Exit Function
errHandler:
    zlCommFun.StopFlash
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
     Call zl_vsGrid_Para_Save(mlngModule, vsfChangeInfo, Me.Name, "变动信息")
End Sub

Private Sub picButton_Resize()
    On Error Resume Next
    cmdExit.Left = picButton.ScaleWidth - cmdExit.Width - 500
End Sub

Private Sub vsfChangeInfo_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call zl_vsGrid_Para_Save(mlngModule, vsfChangeInfo, Me.Name, "变动信息")
End Sub

Private Sub vsfChangeInfo_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
End Sub

Private Sub picImgPlan_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    
    vRect = zlControl.GetControlRect(picImgPlan.Hwnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgPlan.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsfChangeInfo, lngLeft, lngTop, picImgPlan.Height)
    Call zl_vsGrid_Para_Save(mlngModule, vsfChangeInfo, Me.Name, "变动信息")
End Sub
