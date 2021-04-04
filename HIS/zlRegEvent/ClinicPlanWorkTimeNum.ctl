VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ClinicPlanWorkTimeNum 
   ClientHeight    =   6210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9270
   ScaleHeight     =   6210
   ScaleWidth      =   9270
   Begin VB.PictureBox picFunBack 
      BorderStyle     =   0  'None
      Height          =   350
      Left            =   75
      ScaleHeight     =   345
      ScaleWidth      =   8685
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   8685
      Begin VB.CommandButton cmdFun 
         Caption         =   "清除预约(&R)"
         Height          =   350
         Index           =   4
         Left            =   7410
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   0
         Width           =   1155
      End
      Begin VB.CommandButton cmdFun 
         Caption         =   "全部预约(&A)"
         Height          =   350
         Index           =   3
         Left            =   6165
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   0
         Width           =   1155
      End
      Begin VB.CommandButton cmdFun 
         Caption         =   "其他辅助计算(&E)"
         Height          =   350
         Index           =   2
         Left            =   4574
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
         Width           =   1515
      End
      Begin VB.CommandButton cmdFun 
         Caption         =   "按限号分段(&N)"
         Height          =   350
         Index           =   1
         Left            =   3135
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cmdFun 
         Caption         =   "按频次分段(&C)"
         Height          =   350
         Index           =   0
         Left            =   1740
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   0
         Width           =   1335
      End
      Begin VB.TextBox txtUpd 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   2
         Text            =   "10"
         Top             =   18
         Width           =   345
      End
      Begin MSComCtl2.UpDown updSkip 
         Height          =   315
         Left            =   1410
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   18
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtUpd"
         BuddyDispid     =   196611
         OrigLeft        =   2580
         OrigTop         =   585
         OrigRight       =   2835
         OrigBottom      =   1200
         Max             =   60
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "出诊频次(分)"
         Height          =   180
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   85
         Width           =   1080
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsTimeWork 
      Height          =   4830
      Left            =   30
      TabIndex        =   9
      Top             =   540
      Width           =   8880
      _cx             =   15663
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
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   16772055
      GridColor       =   12632256
      GridColorFixed  =   12632256
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   2
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   0
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"ClinicPlanWorkTimeNum.ctx":0000
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
      Editable        =   2
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
      Begin VB.CommandButton cmd预约 
         Caption         =   "预"
         Height          =   255
         Left            =   4860
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   210
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.CommandButton cmd删除 
         Caption         =   "删"
         Height          =   255
         Left            =   5520
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   210
         Visible         =   0   'False
         Width           =   345
      End
   End
End
Attribute VB_Name = "ClinicPlanWorkTimeNum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'缺省属性值
Private Const m_def_CanReCalic = True

'属性变量:
Dim m_诊疗频次 As Integer
Dim m_IsDataChanged As Boolean
Dim m_EditMode As gRegistPlanEditMode
Private m_CanReCalic As Boolean

Private mobj号序信息集 As 号序信息集
Private mobj上班时段 As 上班时段
Private mcllFixedSN As Collection
Private mcurDate As Date, mcurNextDate As Date
Private mintPreSelFun As Integer  '上次选择的功能
Private mblnClickedFunBtn As Boolean '是否点击过按钮
'*****************************************************************************************
'VsGrid的Cell单列格说明
'1.启用序号且分时段
'  第0列：时间段,yyyy-mm-dd HH:MM:SS
'  第>0列：序号列，共两行
'     a.第一行：序号，存储是否开放预约
'     b.第二行:存储时间段，用分号分隔，格式:开始时间;终止时间 ,时间用yyyy-mm-dd HH:MM:SS表示
'2.不启用序号且分时段
'  列Mod 2:0-表示时间段列，格式为开始时间;终止时间 ,时间用yyyy-mm-dd HH:MM:SS表示
'          1-表示预约数列
'*****************************************************************************************
'事件声明:
Event DataIsChanged()
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "当用户在拥有焦点的对象上释放鼠标发生。"
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "当用户移动鼠标时发生。"
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "当用户在拥有焦点的对象上按下鼠标按钮时发生。"

Event TimeIntervalsChanged(ByVal obj号序信息集 As 号序信息集, ByVal blnClearUnit As Boolean)
'缺省属性值:
Const m_def_诊疗频次 = 5
Const m_def_IsDataChanged = False
Const m_def_EditMode = 0
Private mblnNotClick As Boolean
Private mblnValiedCanSave As Boolean

Public Function LoadData(ByVal obj号序信息集 As 号序信息集, ByVal obj上班时段 As 上班时段, _
    Optional ByVal cllFixedSN As Collection, Optional ByVal blnChanged As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载出诊安排
    '入参:obj号序信息集-出诊安排对象
    '返回:加载成功, 返回true,否则返回False
    '编制:刘兴洪
    '日期:2016-01-12 12:46:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    Set mobj号序信息集 = obj号序信息集
    If mobj号序信息集 Is Nothing Then Set mobj号序信息集 = New 号序信息集
    Set mobj上班时段 = obj上班时段
    Set mcllFixedSN = cllFixedSN
    If mcllFixedSN Is Nothing Then Set mcllFixedSN = New Collection
    m_IsDataChanged = blnChanged
    mblnClickedFunBtn = False
    
    mcurDate = Date: mcurNextDate = Date + 1
    Call InitFace
    LoadData = LoadDatatoGrid(obj号序信息集)
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub ReCalicWordTime()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新计算工作时间段
    '编制:刘兴洪
    '日期:2016-01-13 15:54:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    Call cmdFun_Click(mintPreSelFun)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub InitFace()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化数据
    '编制:刘兴洪
    '日期:2016-01-13 09:52:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    mblnNotClick = True
    If mobj号序信息集 Is Nothing Then txtUpd.Text = m_def_诊疗频次: Exit Sub
    txtUpd.Text = IIf(mobj号序信息集.出诊频次 = 0, m_def_诊疗频次, mobj号序信息集.出诊频次)
    
    picFunBack.Visible = m_CanReCalic And (EditMode = ED_RegistPlan_Edit Or EditMode = ED_RegistPlan_NumLimitModify)
    SetFunVisible mobj号序信息集.是否序号控制
    Call UserControl_Resize
    mblnNotClick = False
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function LoadDatatoGrid(ByVal obj序号集 As 号序信息集) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载数据到网格控件
    '入参:obj序号集-号序信息集
    '返回:加载成功，返回true,否则返回Flase
    '编制:刘兴洪
    '日期:2016-01-12 17:49:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln序号控制 As Boolean
    Dim objColAll As New Collection, obj号序信息 As 号序信息
    
    Err = 0: On Error GoTo Errhand:
    bln序号控制 = obj序号集.是否序号控制
    If bln序号控制 And mobj号序信息集.是否分时段 = False Then Exit Function
    
    SetFunVisible bln序号控制
    'obj序号集必须根据时间先后进行排序，不然要乱
    For Each obj号序信息 In mobj号序信息集
        If mobj号序信息集.是否序号控制 Then
            objColAll.Add Array(obj号序信息.序号, obj号序信息.开始时间, obj号序信息.终止时间, _
                IIf(obj号序信息.是否预约, 1, 0), IIf(obj号序信息.是否停诊, 1, 0))
        Else
            objColAll.Add Array(obj号序信息.序号, obj号序信息.开始时间, obj号序信息.终止时间, obj号序信息.数量)
        End If
    Next
    ShowTimeIntervals bln序号控制, objColAll
    LoadDatatoGrid = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub SetFunVisible(ByVal blnVisible As Boolean)
    cmdFun(1).Visible = blnVisible
    cmdFun(2).Visible = blnVisible
    cmdFun(3).Visible = blnVisible
    cmdFun(4).Visible = blnVisible
    
    cmdFun(3).Enabled = m_EditMode = ED_RegistPlan_Edit Or m_EditMode = ED_RegistPlan_NumLimitModify
    cmdFun(4).Enabled = m_EditMode = ED_RegistPlan_Edit Or m_EditMode = ED_RegistPlan_NumLimitModify
    If mobj号序信息集 Is Nothing Then Exit Sub
    cmdFun(3).Enabled = cmdFun(3).Enabled And mobj号序信息集.预约控制 <> 1
    cmdFun(4).Enabled = cmdFun(4).Enabled And mobj号序信息集.预约控制 <> 1
End Sub

Private Sub cmdFun_Click(index As Integer)
    Dim strTittle As String
    
    On Error GoTo Errhand
    m_IsDataChanged = True: RaiseEvent DataIsChanged
    If mobj号序信息集.是否分时段 Then
        Select Case index
        Case 0 '按频次分段
            If AutoSplitNum(0) = False Then Exit Sub
            mblnClickedFunBtn = True
            mintPreSelFun = index
        Case 1 '按限号数分段
            If AutoSplitNum(1) = False Then Exit Sub
            mblnClickedFunBtn = True
            mintPreSelFun = index
        Case 2 '辅助计算
            If AutoSplitNum(2) = False Then Exit Sub
            mblnClickedFunBtn = True
            mintPreSelFun = index
        Case 3   '全部预约
            Call Set预约标志(False)
        Case 4   '取消预约
            Call Set预约标志(True)
        Case Else
        End Select
    End If
    RaiseEvent TimeIntervalsChanged(Get号序信息集, False)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function AutoSplitNum(ByVal bytType As Byte) As Boolean
    '时间段分段
    '入参：
    '   bytType 0-按频次分段,1-按限号数分段,2-辅助计算
    Dim objFrmClinicWorkTimeOther As New frmClinicWorkTimeOther
    Dim colNumAll As Collection, i As Integer, k As Integer
    Dim objCol As Collection, varTimes As Variant, varTemp As Variant
    Dim varTime As Variant, intInterval As Integer
    Dim dtStartDate As Date, dtEndDate As Date
    Dim bln分时段 As Boolean, bln序号控制 As Boolean, lng限号数 As Long, lng限约数 As Long
    Dim str开始时间 As String, str终止时间 As String, str休息时段 As String
    Dim lng预留时间 As Long, dtStart As Date, dtEnd As Date
    Dim dtCurStart As Date, dtCurEnd As Date, dtCur As Date
    Dim lngCount As Long, lngOverplus As Long, lngCurSN As Long
    Dim colTemp As Collection
    
    Err = 0: On Error GoTo Errhand:
    bln分时段 = True: bln序号控制 = True
    If Not mobj号序信息集 Is Nothing Then
        bln分时段 = mobj号序信息集.是否分时段
        bln序号控制 = mobj号序信息集.是否序号控制
        If Not mobj上班时段 Is Nothing Then
            str开始时间 = mobj上班时段.开始时间
            str终止时间 = mobj上班时段.结束时间
            str休息时段 = mobj上班时段.休息时段
            lng预留时间 = mobj上班时段.出诊预留时间
        End If
        lng限号数 = mobj号序信息集.限号数
        lng限约数 = IIf(mobj号序信息集.预约控制 = 1, 0, _
            IIf(mobj号序信息集.限约数 = 0 And mobj号序信息集.限号数 <> 0, mobj号序信息集.限号数, mobj号序信息集.限约数))
    End If
    
    If bln序号控制 And lng限号数 <= 0 Then
        MsgBox "限号数未设置，请先设置限号数！", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If str开始时间 = "" Then
        dtStartDate = Format(mcurDate, "yyyy-mm-dd") & " 01:00:00"
        dtEndDate = Format(mcurDate, "yyyy-mm-dd") & " 23:59:59"
    Else
        dtStartDate = Format(mcurDate, "yyyy-mm-dd") & " " & Format(str开始时间, "HH:MM")
        dtEndDate = GetWorkTrueDate(dtStartDate, str终止时间)
    End If
    
    '减去预留时间
    Call 减去预留时间(dtStartDate, dtEndDate, lng预留时间, str休息时段)
    
    If Val(txtUpd.Text) = 0 And bytType <> 2 Then
        MsgBox "频次未设置，不能分时段！", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    Set colNumAll = New Collection
    Select Case bytType
    Case 0 '按频次分段
        Set colNumAll = CalculatTimeInterval(0, bln序号控制, Val(txtUpd.Text), lng限号数, dtStartDate, dtEndDate, str休息时段)
    Case 1 '按限号数分段
        Set colNumAll = CalculatTimeInterval(1, bln序号控制, Val(txtUpd.Text), lng限号数, dtStartDate, dtEndDate, str休息时段)
    Case 2 '辅助计算
        If bln序号控制 = False Then AutoSplitNum = True: Exit Function
        If objFrmClinicWorkTimeOther.ShowMe(Me, Val(txtUpd.Text), dtStartDate, dtEndDate, str休息时段, varTimes) = False Then Exit Function
        If varTimes(0) = "时间间隔" Then
            Set colNumAll = CalculatTimeInterval(0, bln序号控制, Val(varTimes(1)), lng限号数, dtStartDate, dtEndDate)
        ElseIf varTimes(0) = "分段间隔" Then
            varTemp = Split(varTimes(1), ";")
            For i = 0 To UBound(varTemp)
                varTime = Split(varTemp(i), ",")(0): intInterval = Val(Split(varTemp(i), ",")(1))
                Set objCol = CalculatTimeInterval(0, bln序号控制, intInterval, lng限号数, Split(varTime, "～")(0), Split(varTime, "～")(1), "", colNumAll.Count + 1)
                Set colNumAll = AddRange(colNumAll, objCol)
            Next
        End If
    End Select
    
    '重新调整预约数量
    If bln分时段 And bln序号控制 = False Then
        Set colTemp = New Collection
        intInterval = lng限约数 \ colNumAll.Count '每个时段的平均限约数
        lngOverplus = lng限约数 - intInterval * colNumAll.Count '剩余未分配完的限约数，将分配到前面的序号上
        For i = 1 To colNumAll.Count
            'Array(序号,开始时间,终止时间,预约数量)
            If intInterval = 0 Then
                '平均限约数等于零时，从前分配，否则多余的放于后面
                colTemp.Add Array(colNumAll(i)(0), colNumAll(i)(1), colNumAll(i)(2), IIf(i <= lngOverplus, 1, 0)), "K_" & i
            Else
                colTemp.Add Array(colNumAll(i)(0), colNumAll(i)(1), colNumAll(i)(2), _
                    intInterval + IIf(i > colNumAll.Count - lngOverplus, 1, 0)), "K_" & i
            End If
        Next
        Set colNumAll = colTemp
    ElseIf lng限约数 > 0 Then
        Set colTemp = New Collection
        For i = 1 To colNumAll.Count
            'Array(序号,开始时间,终止时间,预约数量)
            colTemp.Add Array(colNumAll(i)(0), colNumAll(i)(1), colNumAll(i)(2), 1), "K_" & i
        Next
        Set colNumAll = colTemp
    End If
    
    ShowTimeIntervals bln序号控制, colNumAll
    AutoSplitNum = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub ShowTimeIntervals(ByVal bln序号控制 As Boolean, ByVal objCol As Collection)
    '显示数据
    '入参：
    '   bln序号控制：True-序号控制，False-不序号控制
    '   objCol:Array(序号,开始时间,终止时间,是否允许预约/限制数量,是否停诊)
    Dim varItem As Variant, varTemp As Variant
    Dim i As Integer, j As Integer, blnFind As Boolean
    Dim lngRow As Long, lngCol As Long, strCurTime As String
    
    Err = 0: On Error GoTo Errhand:
    With vsTimeWork
        .Clear
        .Rows = 0: .Cols = 0
        If objCol Is Nothing Then Exit Sub
        If objCol.Count = 0 Then Exit Sub
        .Redraw = flexRDNone
        If bln序号控制 Then
            .Rows = 2: .Cols = 2
            .FixedRows = 0: .FixedCols = 1
            .MergeCellsFixed = flexMergeRestrictColumns
            .HighLight = flexHighlightAlways
            .AllowSelection = True
            .MergeCol(0) = True
            lngRow = -2: lngCol = 1: strCurTime = ""
            For Each varItem In objCol
                If strCurTime <> Format(varItem(1), "hh:00") Then
                    strCurTime = Format(varItem(1), "hh:00")
                    lngRow = lngRow + 2: lngCol = 1
                    If lngRow > .Rows - 2 Then .Rows = .Rows + 2
                    .TextMatrix(lngRow, 0) = Format(varItem(1), "hh:00")
                    .TextMatrix(lngRow + 1, 0) = Format(varItem(1), "hh:00")
                End If
                If lngCol > .Cols - 1 Then .Cols = .Cols + 1
                .TextMatrix(lngRow, lngCol) = varItem(0)
                .TextMatrix(lngRow + 1, lngCol) = Format(varItem(1), "hh:mm") & "-" & Format(varItem(2), "hh:mm")
                .Cell(flexcpData, lngRow + 1, lngCol) = Format(varItem(1), "yyyy-mm-dd hh:mm:ss") & "～" & Format(varItem(2), "yyyy-mm-dd hh:mm:ss") '存储时间范围
                If Val(varItem(3)) = 1 Then '是否预约
                    .Cell(flexcpData, lngRow, lngCol) = 1
                    .Cell(flexcpForeColor, lngRow, lngCol, lngRow + 1, lngCol) = vbBlue
                    .Cell(flexcpFontBold, lngRow, lngCol, lngRow + 1, lngCol) = True
                End If
                If UBound(varItem) = 4 Then
                    If Val(varItem(4)) = 1 Then '是否停诊
                        .Cell(flexcpBackColor, lngRow, lngCol, lngRow + 1, lngCol) = vbRed
                    End If
                End If
                lngCol = lngCol + 1
            Next
            .Cell(flexcpAlignment, 0, 1, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
            .Cell(flexcpFontSize, 0, 0, .Rows - 1, 0) = 12
            .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = True
            .Cell(flexcpAlignment, 0, 0, .Rows - 1, 0) = flexAlignCenterTop
            .ColWidth(-1) = 1300: .ColWidth(0) = 800
        Else
            .Clear
            .Rows = 1: .Cols = 8
            .FixedRows = 1: .FixedCols = 0
            .MergeCellsFixed = flexMergeNever
            .HighLight = flexHighlightNever
            .AllowSelection = False
            
            .Editable = IIf(m_EditMode = ED_RegistPlan_Edit Or m_EditMode = ED_RegistPlan_NumLimitModify, flexEDKbdMouse, flexEDNone)
            If Not mobj号序信息集 Is Nothing Then
                .Editable = IIf(.Editable = flexEDKbdMouse And mobj号序信息集.预约控制 <> 1, flexEDKbdMouse, flexEDNone)
            End If
            For i = 0 To .Cols - 1 Step 2
                .Cell(flexcpText, 0, i, 0, i + 1) = "时间段" & vbTab & "预约人数"
            Next
            lngCol = 0: lngRow = 1
            For Each varItem In objCol
                If lngCol > .Cols - 1 Then lngRow = lngRow + 1: lngCol = 0
                If lngRow > .Rows - 1 Then .Rows = .Rows + 1
                .TextMatrix(lngRow, lngCol) = Format(varItem(1), "hh:mm") & "-" & Format(varItem(2), "hh:mm")
                .Cell(flexcpData, lngRow, lngCol) = varItem(0)
                .Cell(flexcpData, lngRow, lngCol + 1) = Format(varItem(1), "yyyy-mm-dd hh:mm:ss") & "～" & Format(varItem(2), "yyyy-mm-dd hh:mm:ss") '存储时间范围
                .TextMatrix(lngRow, lngCol + 1) = Val(varItem(3))
                lngCol = lngCol + 2
            Next
            .ColWidth(-1) = 1200
            .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        End If
        .Redraw = flexRDBuffered
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Set预约标志(ByVal blnClear As Boolean, Optional lngRow As Long = -1, Optional lngCol As Long = -1, _
    Optional ByVal blnIgnoreErr As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置预约
    '入参:blnClear-清除预约
    '     lngRow=-1或lngCol=-1 针对所有进行设置
    '编制:刘兴洪
    '日期:2016-01-13 14:50:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng限约 As Long, lngCount As Long
    Dim lngSum As Long, lng已预约 As Long
    
    Err = 0: On Error GoTo Errhand:
    If Not mobj号序信息集 Is Nothing Then
        lng限约 = mobj号序信息集.限约数
    End If
    If blnClear = False And lng限约 = 0 Then Exit Sub
    
    With vsTimeWork
        If lngRow < 0 Or lngCol <= 0 Then
            For lngRow = 0 To .Rows - 1 Step 2
                If blnClear = False Then
                    For lngCol = 1 To .Cols - 1
                        If .TextMatrix(lngRow, lngCol) <> "" And Val(.Cell(flexcpData, lngRow, lngCol)) = 0 Then
                            .Cell(flexcpData, lngRow, 1, lngRow, lngCol) = 1
                            .Cell(flexcpForeColor, lngRow, 1, lngRow + 1, lngCol) = IIf(blnClear, &H80000008, vbBlue)
                            .Cell(flexcpFontBold, lngRow, 1, lngRow + 1, lngCol) = IIf(blnClear, False, True)
                        End If
                    Next
                Else
                    For lngCol = 1 To .Cols - 1
                        lng已预约 = 0
                        Call ValiedCanModify(Val(.TextMatrix(lngRow, lngCol)), 0, False, lng已预约)
                        blnClear = lng已预约 = 0
                        .Cell(flexcpData, lngRow, lngCol, lngRow, lngCol) = IIf(blnClear, 0, 1)
                        .Cell(flexcpForeColor, lngRow, lngCol, lngRow + 1, lngCol) = IIf(blnClear, &H80000008, vbBlue)
                        .Cell(flexcpFontBold, lngRow, lngCol, lngRow + 1, lngCol) = IIf(blnClear, False, True)
                    Next
                End If
            Next
        Else
'            lngSum = Get预约总数
            If lngRow Mod 2 = 1 Then lngRow = lngRow - 1
'            If lngSum + 1 > lng限约 And blnClear = False And Val(.Cell(flexcpData, lngRow, lngCol)) <> 1 Then
'                If blnIgnoreErr = False Then MsgBox "超过限约数" & lng限约 & "，不能再设置！", vbInformation + vbOKOnly, gstrSysName
'                Exit Sub
'            End If
            .Cell(flexcpData, lngRow, lngCol) = IIf(blnClear, 0, 1)
            .Cell(flexcpForeColor, lngRow, lngCol, lngRow + 1, lngCol) = IIf(blnClear, &H80000008, vbBlue)
            .Cell(flexcpFontBold, lngRow, lngCol, lngRow + 1, lngCol) = IIf(blnClear, False, True)
        End If
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function Get预约总数() As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取预约总数
    '返回:返回预约总数
    '编制:刘兴洪
    '日期:2016-01-13 15:04:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, lngCol As Long
    Dim lngSum As Long
    With vsTimeWork
        lngSum = 0
        For lngRow = 0 To .Rows - 1 Step 2
            For lngCol = 1 To .Cols - 1
                If Val(.Cell(flexcpData, lngRow, lngCol)) = 1 Then lngSum = lngSum + 1
            Next
        Next
    End With
    Get预约总数 = lngSum
End Function

Private Sub cmdFun_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmd删除_Click()
    Dim lngRow As Long
    
    On Error GoTo Errhand
    m_IsDataChanged = True: RaiseEvent DataIsChanged
    With vsTimeWork
        lngRow = .Row
        If lngRow Mod 2 = 1 Then lngRow = lngRow - 1
        If .TextMatrix(lngRow, .Col) = "" Then cmd删除.Visible = False: Exit Sub
        If ValiedCanModify(Val(.TextMatrix(lngRow, .Col)), 0, True) = False Then
            MsgBox "当前时段或当前时段之后的时段已被使用，不能删除！", vbInformation, gstrSysName
            Exit Sub
        End If
        Call DeleteTime(.Row, .Col)
    End With
    RaiseEvent TimeIntervalsChanged(Get号序信息集, False)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmd删除_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmd预约_Click()
    Dim blnClear As Boolean
    Dim lngRow As Long
    Dim i As Long, j As Long
    Dim intStartRow As Integer, intEndRow As Integer
    Dim intStartCol As Integer, intEndCol As Integer
    
    On Error GoTo Errhand
    m_IsDataChanged = True: RaiseEvent DataIsChanged
    With vsTimeWork
        If .Row < 0 Or .Col < 0 Then Exit Sub
        If .Row = .RowSel And .Col = .ColSel Then
            lngRow = .Row
            If lngRow Mod 2 = 1 Then lngRow = lngRow - 1
            If .TextMatrix(lngRow, .Col) = "" Then cmd预约.Visible = False: Exit Sub
            If ValiedCanModify(Val(.TextMatrix(lngRow, .Col)), 0) = False Then
                MsgBox "当前时段已被使用，不能调整！", vbInformation, gstrSysName
                Exit Sub
            End If
            blnClear = Val(.Cell(flexcpData, lngRow, .Col)) = 1
            Call Set预约标志(blnClear, lngRow, .Col)
        Else
            '82227，批量设置
            intStartRow = IIf(.Row > .RowSel, .RowSel, .Row)
            intEndRow = IIf(.Row > .RowSel, .Row, .RowSel)
            intStartCol = IIf(.Col > .ColSel, .ColSel, .Col)
            intEndCol = IIf(.Col > .ColSel, .Col, .ColSel)
            For i = intStartRow To intEndRow Step 2
                For j = intStartCol To intEndCol
                    If .TextMatrix(i, j) <> "" And ValiedCanModify(Val(.TextMatrix(i - (i Mod 2), j)), 0) Then
                        blnClear = Val(.Cell(flexcpData, i - (i Mod 2), j)) = 1
                        Call Set预约标志(blnClear, i, j, True)
                    End If
                Next
            Next
            .Select intStartRow, intStartCol
        End If
    End With
    RaiseEvent TimeIntervalsChanged(Get号序信息集, False)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmd预约_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtUpd_Change()
    If mblnNotClick Then Exit Sub
    m_IsDataChanged = True: RaiseEvent DataIsChanged
End Sub

Private Sub txtUpd_GotFocus()
    zlControl.TxtSelAll txtUpd
End Sub

Private Sub txtUpd_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub txtUpd_LostFocus()
    If mobj号序信息集 Is Nothing Then Exit Sub
    mobj号序信息集.出诊频次 = Val(txtUpd.Text)
End Sub

Private Sub txtUpd_Validate(Cancel As Boolean)
    If Val(txtUpd.Text) > 60 Or Val(txtUpd.Text) < 1 Then
        MsgBox "出诊频次不能大于60分钟或小于1分钟！", vbInformation, gstrSysName
        zlControl.TxtSelAll txtUpd
        Cancel = True: Exit Sub
    End If
End Sub

Private Sub updSkip_Change()
    If mblnNotClick Then Exit Sub
    m_IsDataChanged = True: RaiseEvent DataIsChanged
    mobj号序信息集.出诊频次 = Val(txtUpd.Text)
End Sub

Private Sub UserControl_Initialize()
    mcurDate = Date: mcurNextDate = Date + 1
End Sub

Private Sub UserControl_LostFocus()
    cmd删除.Visible = False
End Sub

Private Sub UserControl_Resize()
    Err = 0: On Error Resume Next
    With UserControl
        vsTimeWork.Left = .ScaleLeft
        vsTimeWork.Top = IIf(m_CanReCalic And (EditMode = ED_RegistPlan_Edit Or EditMode = ED_RegistPlan_NumLimitModify), picFunBack.Top + picFunBack.Height, 0) + 30
        vsTimeWork.Height = .ScaleHeight - vsTimeWork.Top
        vsTimeWork.Width = .ScaleWidth
    End With
End Sub

Private Function CheckAutoSplitDateIsValied(ByVal dtCurdate As Date) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查自动分配日期是否合法
    '入参:dtCurDate-当前日期
    '返回:合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2016-01-13 11:00:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, varTemp As Variant
    Dim i As Long, str开始时间 As String, str结束时间 As String
    Dim dtStartDate As Date, dtEndDate As Date
    
    Err = 0: On Error GoTo Errhand:
    CheckAutoSplitDateIsValied = True
    If mobj号序信息集 Is Nothing Then Exit Function
    If mobj上班时段.休息时段 = "" Then Exit Function
    
    varData = Split(mobj上班时段.休息时段, ";")
    For i = 0 To UBound(varData)
        If varData(i) <> "" Then
            varTemp = Split(varData(i), "-")
            If UBound(varTemp) <> 0 Then
                str开始时间 = varTemp(0)
                str结束时间 = varTemp(1)
                If CDate(str开始时间) > CDate(str结束时间) Then
                    dtStartDate = CDate(Format(mcurDate, "yyyy-mm-dd") & " " & str开始时间 & ":00")
                    dtEndDate = CDate(Format(mcurNextDate, "yyyy-mm-dd") & " " & str结束时间 & ":59")
                Else
                    dtStartDate = CDate(Format(mcurDate, "yyyy-mm-dd") & " " & str开始时间 & ":00")
                    dtEndDate = CDate(Format(mcurDate, "yyyy-mm-dd") & " " & str结束时间 & ":59")
                End If
                If dtCurdate >= dtStartDate And dtCurdate <= dtEndDate Then
                    CheckAutoSplitDateIsValied = False: Exit Function
                End If
            End If
        End If
    Next
    CheckAutoSplitDateIsValied = True
    Exit Function
Errhand:
    CheckAutoSplitDateIsValied = True
End Function

Private Function Get号序信息集() As 号序信息集
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取号序信息集
    '返回:号序信息集
    '编制:刘兴洪
    '日期:2016-01-13 12:34:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, lngCol As Integer
    Dim objNums As New 号序信息集, objNum As 号序信息, bln预约 As Boolean
    Dim lngSum As Long, lngNO As Long, varTemp As Variant
    Dim i As Long
    Dim blnFind As Boolean '启用时段序号控制时，是否有设置可预约时段，如果一个都没设置的话，默认所有时段都可预约
    
    Err = 0: On Error GoTo Errhand:
    
    '数据未改变，直接返回原集合的副本
    If m_IsDataChanged = False Then
        Set Get号序信息集 = mobj号序信息集.Clone
        Exit Function
    End If
    
    '数据已改变，重新构造集合对象
    Set objNums = mobj号序信息集.Clone
    objNums.RemoveAll
    objNums.是否修改 = True
    If objNums.是否分时段 And Not objNums.是否序号控制 And vsTimeWork.FixedCols = 0 Then
        For lngRow = 1 To vsTimeWork.Rows - 1
            For lngCol = 0 To vsTimeWork.Cols - 1 Step 2
               If vsTimeWork.TextMatrix(lngRow, lngCol) <> "" Then
                    lngNO = Val(vsTimeWork.Cell(flexcpData, lngRow, lngCol))
                    varTemp = Split(vsTimeWork.Cell(flexcpData, lngRow, lngCol + 1), "～")
                    lngSum = Val(vsTimeWork.TextMatrix(lngRow, lngCol + 1))
                    Set objNum = New 号序信息
                    With objNum
                        .序号 = lngNO
                        .开始时间 = varTemp(0)
                        .终止时间 = varTemp(1)
                        .数量 = lngSum
                        .是否预约 = True
                    End With
                    objNums.AddItem objNum
               End If
            Next
        Next
    ElseIf objNums.是否分时段 And objNums.是否序号控制 And vsTimeWork.FixedCols = 1 Then
        For lngRow = 0 To vsTimeWork.Rows - 1 Step 2
            For lngCol = 1 To vsTimeWork.Cols - 1
                If vsTimeWork.TextMatrix(lngRow, lngCol) <> "" _
                    And vsTimeWork.Cell(flexcpFontStrikethru, lngRow, lngCol) = False Then '有删除线的表示本次要删除的
                    
                    lngNO = Val(vsTimeWork.TextMatrix(lngRow, lngCol))
                    varTemp = Split(vsTimeWork.Cell(flexcpData, lngRow + 1, lngCol), "～")
                    bln预约 = Val(vsTimeWork.Cell(flexcpData, lngRow, lngCol)) = 1
                    If bln预约 Then blnFind = True
                    Set objNum = New 号序信息
                    With objNum
                        .序号 = lngNO
                        .开始时间 = varTemp(0)
                        .终止时间 = varTemp(1)
                        .数量 = 1
                        .是否预约 = bln预约
                    End With
                    objNums.AddItem objNum
                End If
            Next
        Next
'        If blnFind = False And mobj号序信息集.预约控制 <> 1 Then
'            '全部允许预约
'            For i = 1 To objNums.Count
'                objNums(i).是否预约 = True
'            Next
'        End If
    ElseIf objNums.是否分时段 = False And objNums.是否序号控制 Then '启用序号不分时段的自动产生序号
        For i = 1 To objNums.限号数
            Set objNum = New 号序信息
            With objNum
                .序号 = i
                .数量 = 1
                .是否预约 = True '都允许预约
                '时间范围填写为时间段的开始时间和终止时间
                If Not mobj上班时段 Is Nothing Then
                    .开始时间 = mobj上班时段.开始时间
                    .终止时间 = mobj上班时段.结束时间
                End If
            End With
            objNums.AddItem objNum
        Next
    End If
    Set Get号序信息集 = objNums
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

'注意！不要删除或修改下列被注释的行！
'MemberInfo=14,0,0,0
Public Property Get Get号序集() As 号序信息集
   Set Get号序集 = Get号序信息集
End Property

Public Property Get 限号数() As Long
    If mobj号序信息集 Is Nothing Then Set mobj号序信息集 = New 号序信息集
    限号数 = mobj号序信息集.限号数
End Property

Public Property Let 限号数(ByVal vNewValue As Long)
    Dim lngOld As Long
    
    On Error GoTo Errhand
    If mobj号序信息集 Is Nothing Then Set mobj号序信息集 = New 号序信息集
    lngOld = mobj号序信息集.限号数
    mobj号序信息集.限号数 = vNewValue
    
    If mobj号序信息集.限号数 = lngOld Then Exit Property
    
    m_IsDataChanged = True: RaiseEvent DataIsChanged
    If mobj号序信息集.限号数 = 0 Then
        ShowTimeIntervals True, Nothing
        RaiseEvent TimeIntervalsChanged(Get号序信息集, False)
        Exit Property
    End If
    
    '重新计算时段
    Call cmdFun_Click(mintPreSelFun)
    Exit Property
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Property

Public Property Get 限约数() As Long
   If mobj号序信息集 Is Nothing Then Set mobj号序信息集 = New 号序信息集
   限约数 = mobj号序信息集.限约数
End Property

Public Property Let 限约数(ByVal vNewValue As Long)
    Dim lngOld As Long, lng已预约 As Long
    Dim lngRow As Long, lngCol As Long, lngSum As Long
    
    If mobj号序信息集 Is Nothing Then Set mobj号序信息集 = New 号序信息集
    lngOld = mobj号序信息集.限约数
    mobj号序信息集.限约数 = vNewValue
    
    If mobj号序信息集.限约数 = lngOld Then Exit Property
    
    On Error GoTo Errhand
    m_IsDataChanged = True: RaiseEvent DataIsChanged
    lngSum = mobj号序信息集.限约数
    If mobj号序信息集.是否分时段 Then
        If mobj号序信息集.是否序号控制 Then '分时段，序号控制
            If mobj号序信息集.限约数 = 0 Then
                '全部取消预约
                Call Set预约标志(True)
            End If
        ElseIf mobj号序信息集.预约控制 = 1 Then '分时段，不序号控制
            For lngRow = 1 To vsTimeWork.Rows - 1
                For lngCol = 0 To vsTimeWork.Cols - 1 Step 2
                    If vsTimeWork.TextMatrix(lngRow, lngCol) <> "" Then
                        lng已预约 = 0
                        Call ValiedCanModify(Val(vsTimeWork.Cell(flexcpData, lngRow, lngCol)), 0, False, lng已预约)
                        vsTimeWork.TextMatrix(lngRow, lngCol + 1) = lng已预约
                    End If
                Next
            Next
        End If
    End If
    RaiseEvent TimeIntervalsChanged(Get号序信息集, False)
    Exit Property
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Property

Public Property Get 启用序号控制() As Boolean
    If mobj号序信息集 Is Nothing Then Set mobj号序信息集 = New 号序信息集
    启用序号控制 = mobj号序信息集.是否序号控制
End Property

Public Property Let 启用序号控制(ByVal vNewValue As Boolean)
    If mobj号序信息集 Is Nothing Then Set mobj号序信息集 = New 号序信息集
    mobj号序信息集.是否序号控制 = vNewValue
    
    m_IsDataChanged = True: RaiseEvent DataIsChanged
    
    SetFunVisible vNewValue
    ShowTimeIntervals mobj号序信息集.是否序号控制, Nothing
    RaiseEvent TimeIntervalsChanged(Get号序信息集, True)
End Property

Public Property Get 启用时段() As Boolean
    If mobj号序信息集 Is Nothing Then Set mobj号序信息集 = New 号序信息集
    启用时段 = mobj号序信息集.是否分时段
End Property

Public Property Let 启用时段(ByVal vNewValue As Boolean)
    If mobj号序信息集 Is Nothing Then Set mobj号序信息集 = New 号序信息集
    mobj号序信息集.是否分时段 = vNewValue
    
    m_IsDataChanged = True: RaiseEvent DataIsChanged
    
    ShowTimeIntervals mobj号序信息集.是否序号控制, Nothing
    RaiseEvent TimeIntervalsChanged(Get号序信息集, True)
End Property

Public Property Get 预约控制() As Integer
    If mobj号序信息集 Is Nothing Then Set mobj号序信息集 = New 号序信息集
    预约控制 = mobj号序信息集.预约控制
End Property

Public Property Let 预约控制(ByVal vNewValue As Integer)
    If mobj号序信息集 Is Nothing Then Set mobj号序信息集 = New 号序信息集
    mobj号序信息集.预约控制 = vNewValue
    
    cmdFun(3).Enabled = (m_EditMode = ED_RegistPlan_Edit Or m_EditMode = ED_RegistPlan_NumLimitModify) And mobj号序信息集.预约控制 <> 1
    cmdFun(4).Enabled = (m_EditMode = ED_RegistPlan_Edit Or m_EditMode = ED_RegistPlan_NumLimitModify) And mobj号序信息集.预约控制 <> 1
    vsTimeWork.Editable = IIf((m_EditMode = ED_RegistPlan_Edit Or m_EditMode = ED_RegistPlan_NumLimitModify) And mobj号序信息集.预约控制 <> 1, flexEDKbdMouse, flexEDNone)
    
    '隐藏按钮
    cmd预约.Visible = False
End Property

Private Sub SetCtrlMove()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:移动控件
    '编制:刘兴洪
    '日期:2016-01-13 14:23:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln启用序号 As Boolean, bln启用时段 As Boolean
    Dim blnDel As Boolean, lng限约 As Long
    Dim lngRow As Long
    Dim lngCol As Long
    Err = 0: On Error GoTo Errhand:
    
    bln启用序号 = True: bln启用时段 = True
    lng限约 = 0
    If Not mobj号序信息集 Is Nothing Then
        bln启用序号 = mobj号序信息集.是否序号控制
        bln启用时段 = mobj号序信息集.是否分时段
        lng限约 = mobj号序信息集.限约数
    End If
    
    cmd预约.Visible = False
    cmd删除.Visible = False
    If Not (bln启用序号 And bln启用时段) Then Exit Sub
    
    With vsTimeWork
        If .Col < 0 And .Cols > 2 Then .Col = 1
        If .Col < 0 Or .Row < 0 Then Exit Sub
        If .TextMatrix(.Row, .Col) = "" Then Exit Sub
        If .Cell(flexcpFontStrikethru, .Row, .Col) Then Exit Sub  '有删除线的表示本次要删除的
        
        lngRow = .Row
        If lngRow Mod 2 = 0 Then lngRow = lngRow + 1
        cmd预约.Left = .CellLeft
        cmd删除.Left = .CellLeft + .CellWidth - cmd删除.Width - 15
        If .Row Mod 2 = 0 Then
            cmd预约.Top = .CellTop
            cmd删除.Top = .CellTop
        Else
            cmd预约.Top = .Cell(flexcpTop, .Row - 1, .Col)
            cmd删除.Top = cmd预约.Top
        End If
        
        cmd预约.Visible = lng限约 <> 0
        cmd预约.Refresh '防止按钮看不见
        
        '删除按钮最后一列才显示
        For lngCol = .Cols - 1 To 1 Step -1
            If Trim(.TextMatrix(.Row, lngCol)) <> "" Then
                If lngCol = .Col Then
                    cmd删除.Visible = True
                    cmd删除.Refresh '防止按钮看不见
                End If
                Exit For
            End If
        Next
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub UserControl_Terminate()
    Set mobj号序信息集 = Nothing
    Set mobj上班时段 = Nothing
    Set mcllFixedSN = Nothing
End Sub

Private Sub vsTimeWork_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    m_IsDataChanged = True: RaiseEvent DataIsChanged
    RaiseEvent TimeIntervalsChanged(Get号序信息集, False)
End Sub

Private Sub vsTimeWork_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If Not (m_EditMode = ED_RegistPlan_Edit Or m_EditMode = ED_RegistPlan_NumLimitModify) Then Exit Sub
    Call SetCtrlMove
End Sub

Private Sub DeleteTime(ByVal lngRow As Long, ByVal lngCol As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:删除时间段
    '编制:刘兴洪
    '日期:2016-01-13 15:13:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, lngSkip As Long
    Err = 0: On Error GoTo Errhand:
    With vsTimeWork
        If lngRow Mod 2 = 1 Then lngRow = lngRow - 1
        If lngCol < 1 Or lngCol > .Cols - 1 Then Exit Sub
        If lngRow < 0 Or lngRow > .Rows - 1 Then Exit Sub
        If lngCol = .Cols - 1 Then
            .TextMatrix(lngRow, lngCol) = ""
            .TextMatrix(lngRow + 1, lngCol) = ""
            .Cell(flexcpData, lngRow, lngCol) = ""
            .Cell(flexcpData, lngRow + 1, lngCol) = ""
            .Cell(flexcpForeColor, lngRow, lngCol, lngRow + 1, lngCol) = &H80000008
        Else
            For i = lngCol To .Cols - 2
                lngSkip = i + 1
                .TextMatrix(lngRow, i) = .TextMatrix(lngRow, lngSkip)
                .TextMatrix(lngRow + 1, i) = .TextMatrix(lngRow + 1, lngSkip)
                .Cell(flexcpData, lngRow, i) = .Cell(flexcpData, lngRow, lngSkip)
                .Cell(flexcpData, lngRow + 1, i) = .Cell(flexcpData, lngRow + 1, lngSkip)
                .Cell(flexcpForeColor, lngRow, i, lngRow + 1, i) = .Cell(flexcpForeColor, lngRow, lngSkip, lngRow + 1, lngSkip)
                
                .TextMatrix(lngRow, lngSkip) = ""
                .TextMatrix(lngRow + 1, lngSkip) = ""
                .Cell(flexcpData, lngRow, lngSkip) = ""
                .Cell(flexcpData, lngRow + 1, lngSkip) = ""
                .Cell(flexcpForeColor, lngRow, lngSkip, lngRow + 1, lngSkip) = &H80000008
            Next
        End If
    End With
    Call ReSetNumNo
    If vsTimeWork.TextMatrix(lngRow, lngCol) = "" Then cmd删除.Visible = False: cmd预约.Visible = False
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub ReSetNumNo()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新调整序号
    '编制:刘兴洪
    '日期:2016-01-13 15:20:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, lngCol As Long, lngNumNo As Long
    
    With vsTimeWork
        lngNumNo = 0
        For lngRow = 0 To .Rows - 1 Step 2
            For lngCol = 1 To .Cols - 1
               If Trim(.TextMatrix(lngRow, lngCol)) <> "" Then
                    lngNumNo = lngNumNo + 1
                    .TextMatrix(lngRow, lngCol) = lngNumNo
               End If
            Next
        Next
    End With
End Sub
 
'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "返回/设置对象中文本和图形的背景色。"
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor = New_BackColor
    PropertyChanged "BackColor"
    SetBackColor Controls, UserControl.BackColor
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "指出 Label 或 Shape 的背景样式是透明的还是不透明的。"
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "返回一个 Font 对象。"
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,0
Public Property Get CanReCalic() As Boolean
    CanReCalic = m_CanReCalic
End Property

Public Property Let CanReCalic(ByVal New_CanReCalic As Boolean)
    m_CanReCalic = New_CanReCalic
    PropertyChanged "CanReCalic"
    picFunBack.Visible = m_CanReCalic
    Call UserControl_Resize
End Property

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "返回/设置当鼠标经过对象某一部分时鼠标的指针类型。"
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'为用户控件初始化属性
Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
    m_CanReCalic = m_def_CanReCalic
    m_EditMode = m_def_EditMode
    m_IsDataChanged = m_def_IsDataChanged
    m_诊疗频次 = m_def_诊疗频次
    txtUpd.Text = IIf(m_诊疗频次 = 0, m_def_诊疗频次, m_诊疗频次)
End Sub

'从存贮器中加载属性值
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_CanReCalic = PropBag.ReadProperty("CanReCalic", m_def_CanReCalic)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    m_EditMode = PropBag.ReadProperty("EditMode", m_def_EditMode)
    m_IsDataChanged = PropBag.ReadProperty("IsDataChanged", m_def_IsDataChanged)
    m_诊疗频次 = PropBag.ReadProperty("诊疗频次", m_def_诊疗频次)
    txtUpd.Text = IIf(m_诊疗频次 = 0, m_def_诊疗频次, m_诊疗频次)
End Sub

'将属性值写到存储器
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("CanReCalic", m_CanReCalic, m_def_CanReCalic)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("EditMode", m_EditMode, m_def_EditMode)
    Call PropBag.WriteProperty("IsDataChanged", m_IsDataChanged, m_def_IsDataChanged)
    Call PropBag.WriteProperty("诊疗频次", m_诊疗频次, m_def_诊疗频次)

End Sub

Private Sub vsTimeWork_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    cmd预约.Visible = False
    cmd删除.Visible = False
End Sub

Private Sub vsTimeWork_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not (m_EditMode = ED_RegistPlan_Edit Or m_EditMode = ED_RegistPlan_NumLimitModify) Then Cancel = True: Exit Sub
    If mobj号序信息集 Is Nothing Then Cancel = True: Exit Sub
    If mobj号序信息集.是否序号控制 Then Cancel = True: Exit Sub
    If Col Mod 2 = 0 Then Cancel = True: Exit Sub
    If vsTimeWork.Cell(flexcpData, Row, Col) = "" Then Cancel = True: Exit Sub
End Sub

Private Sub vsTimeWork_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And vsTimeWork.Editable = flexEDKbdMouse Then
        If vsTimeWork.Row = vsTimeWork.Rows - 1 And vsTimeWork.Col = vsTimeWork.Cols - 1 Then
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            Call ToNextGridPostion(vsTimeWork, 1, 2, 1, 1)
        End If
        KeyCode = 0
    End If
End Sub

Private Sub ToNextGridPostion(vsfGrid As VSFlexGrid, Optional ByVal lngStepRow As Long = 1, Optional ByVal lngStepCol As Long = 1, _
    Optional ByVal lngFirstRow As Long, Optional ByVal lngFirstCol As Long)
    '功能：自动跳到下一个单元格
    '入参：
    '   lngStepRow - 间隔行
    '   lngStepCol - 间隔列
    '   lngFirstRow - 第一行
    '   lngFirstCol - 第一列
    Dim lngCurRow As Long, lngCurCol As Long
    With vsfGrid
        If lngFirstRow < vsfGrid.FixedRows Then lngFirstRow = vsfGrid.FixedRows
        If lngFirstCol < vsfGrid.FixedCols Then lngFirstCol = vsfGrid.FixedCols
        
        lngCurRow = .Row: lngCurCol = .Col
        If lngCurCol < lngFirstCol Then
            lngCurCol = lngFirstCol
            .Col = lngCurCol
            Exit Sub
        End If
        
        If (lngCurCol - lngFirstCol) Mod lngStepCol <> 0 Then
            lngCurCol = lngFirstCol + (lngCurCol - lngFirstCol) \ lngStepCol * lngStepCol
            If lngCurCol < lngFirstCol Then lngCurCol = lngFirstCol
        End If
        '确定下一个列
        If lngCurCol + lngStepCol > .Cols - 1 Then
            lngCurCol = lngFirstCol
            '确定下一个行
            If lngCurRow + lngStepRow > .Rows - 1 Then
                lngCurRow = lngFirstRow
            Else
                lngCurRow = lngCurRow + lngStepRow
            End If
        Else
            lngCurCol = lngCurCol + lngStepCol
        End If
        .Row = lngCurRow: .Col = lngCurCol
    End With
End Sub

Private Sub vsTimeWork_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub vsTimeWork_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyBack Then Exit Sub
    '输入位数限制，整数位长度不能大于9
    If InStr(vsTimeWork.EditText, ".") > 0 Then
        If InStr(vsTimeWork.EditText, ".") > 9 Then KeyAscii = 0
    Else
        If Len(vsTimeWork.EditText) >= 9 Then KeyAscii = 0
    End If
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Public Function IsValied(Optional ByVal blnChanged As Boolean) As Boolean
    '检查数据
    '外面一层是否改变，若改变则本层也要检查
    Dim lngSum As Long, lngCount As Long
    Dim bln启用序号 As Boolean, bln启用时段 As Boolean
    Dim lng限号数 As Long, lng限约数 As Long
    Dim byt预约控制 As Byte, str上班时段 As String
    Dim lngRow As Long, lngCol As Long, bln预约 As Boolean
    Dim strFirstStart As String
    
    Err = 0: On Error GoTo ErrHandler
    '数据未改变不检查
    If m_IsDataChanged = False And blnChanged = False Then IsValied = True: Exit Function
    If mobj号序信息集 Is Nothing Then IsValied = True: Exit Function
    
    If Not mobj上班时段 Is Nothing Then str上班时段 = mobj上班时段.时间段
    bln启用序号 = mobj号序信息集.是否序号控制
    bln启用时段 = mobj号序信息集.是否分时段
    lng限号数 = mobj号序信息集.限号数
    lng限约数 = IIf(mobj号序信息集.预约控制 = 1, 0, _
            IIf(mobj号序信息集.限约数 = 0 And mobj号序信息集.限号数 <> 0, mobj号序信息集.限号数, mobj号序信息集.限约数))
    byt预约控制 = mobj号序信息集.预约控制
    
    If bln启用时段 = False Then IsValied = True: Exit Function
    '----------------------------------------------------------------
    '特殊处理，网格处于正在编辑状态时，检查不到
    mblnValiedCanSave = True
    vsTimeWork.FinishEditing False
    If mblnValiedCanSave = False Then
        Exit Function
    Else
        mblnValiedCanSave = False
    End If
    '----------------------------------------------------------------
    
    With vsTimeWork
        If Not bln启用序号 Then
            For lngRow = 1 To .Rows - 1
                For lngCol = 0 To .Cols - 1 Step 2
                    If .TextMatrix(lngRow, lngCol) <> "" Then
                        lngCount = lngCount + 1
                        lngSum = lngSum + Val(.TextMatrix(lngRow, lngCol + 1))
                        
                        If lngRow = 1 And lngCol = 0 Then
                            strFirstStart = Split(.Cell(flexcpData, lngRow, lngCol + 1), "～")(0)
                        End If
                    End If
                Next
            Next
            If lngSum = 0 And byt预约控制 <> 1 Then
                MsgBox IIf(m_EditMode = ED_RegistPlan_NumLimitModify, "", str上班时段) & "启用了时段则必须要设置限约时段！", vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
            If lngSum > lng限约数 Then
                MsgBox IIf(m_EditMode = ED_RegistPlan_NumLimitModify, "", str上班时段 & "的") & "可预约人数(" & lngSum & ")超过了限约数(" & lng限约数 & ")！", vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
        Else
            For lngRow = 0 To .Rows - 1 Step 2
                For lngCol = 1 To .Cols - 1
                    If .TextMatrix(lngRow, lngCol) <> "" Then
                        lngCount = lngCount + 1
                        bln预约 = Val(.Cell(flexcpData, lngRow, lngCol)) = 1
                        If bln预约 Then lngSum = lngSum + 1
                        
                        If lngRow = 0 And lngCol = 1 Then
                            strFirstStart = Split(.Cell(flexcpData, lngRow + 1, lngCol), "～")(0)
                        End If
                    End If
                Next
            Next
            'If lngSum = 0 Then lngSum = lngCount '等于零表示全部预约
            If lngSum = 0 And byt预约控制 <> 1 Then
                MsgBox IIf(m_EditMode = ED_RegistPlan_NumLimitModify, "", str上班时段) & "启用了时段则必须要设置可预约时段！", vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
            If lngSum < lng限约数 Then
                If MsgBox("注意：" & vbCrLf & _
                    "        " & IIf(m_EditMode = ED_RegistPlan_NumLimitModify, "", str上班时段) & " 可预约时间段的总数(" & lngSum & ")与限约数(" & lng限约数 & ")不等，你确定按当前设置保存吗？", _
                    vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            End If
        End If
    End With
    
    If lngCount = 0 Then
        MsgBox "启用时段时必须要分配时段！", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If Not mobj上班时段 Is Nothing Then
        '检查第一个时段的开始时间是否等于上班时段的开始时间
        If IsDate(strFirstStart) Then
            strFirstStart = Format(mobj上班时段.开始时间, "yyyy-mm-dd") & " " & Format(strFirstStart, "hh:mm:ss")
            If DateDiff("n", mobj上班时段.开始时间, strFirstStart) <> 0 Then
                If MsgBox(mobj上班时段.时间段 & " 第一个序号时段的开始时间(" & Format(strFirstStart, "hh:mm") & _
                    ")与当前上班时段的开始时间(" & Format(mobj上班时段.开始时间, "hh:mm") & _
                    ")不同，你确定按当前设置保存吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            End If
        End If
    End If
    IsValied = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub vsTimeWork_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lngSum As Long
    Dim lngRow As Long, lngCol As Long
    Dim lng已预约 As Long
    
    On Error GoTo Errhand
    If mobj号序信息集 Is Nothing Then Cancel = True: Exit Sub
    If mobj号序信息集.是否序号控制 Then Cancel = True: Exit Sub
    
    With vsTimeWork
        '整数位多余9位的直接截掉,防止溢出
        If InStr(.EditText, ".") > 0 Then
            If InStr(.EditText, ".") > 9 Then
                 .EditText = Left(.EditText, 9)
            End If
        Else
             .EditText = Left(.EditText, 9)
        End If
    
        If ValiedCanModify(Val(.Cell(flexcpData, .Row, .Col - 1)), Val(.EditText), False, lng已预约) = False Then
            MsgBox "当前时段已预约 " & lng已预约 & " ，调整预约数不能小于 " & lng已预约 & " ！", vbInformation, gstrSysName
            Cancel = True: mblnValiedCanSave = False: Exit Sub
        End If
        
        For lngRow = 1 To .Rows - 1
            For lngCol = 0 To .Cols - 1 Step 2
                If .TextMatrix(lngRow, lngCol) <> "" Then
                    If lngRow = Row And lngCol = Col - 1 Then
                        lngSum = lngSum + Val(.EditText)
                    Else
                        lngSum = lngSum + Val(.TextMatrix(lngRow, lngCol + 1))
                    End If
                End If
            Next
        Next
        If lngSum > mobj号序信息集.限约数 Then
            If Val(.EditText) < Val(.TextMatrix(.Row, .Col)) Or Val(.EditText) = 0 Then
                Exit Sub
            Else
                MsgBox "预约数(" & lngSum & ")不能超过限约数(" & mobj号序信息集.限约数 & ")！", vbInformation + vbOKOnly, gstrSysName
                Cancel = True: mblnValiedCanSave = False: Exit Sub
            End If
        End If
        .EditText = FormatEx(Val(.EditText), 0)
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub SetNewSN(ByVal lng限号数 As Long, ByVal lngCurAdd As Long, ByVal blnAdd As Boolean)
    '加号或者减号
    '参数
    '   lng限号数：当前限号数
    '   lngCurAdd：本次增加限号数，加号为正，减号为负
    '   blnAdd：是否加号操作
    Dim dtStart As Date, dtEnd As Date, intStep As Integer
    Dim colNumAll As Collection, objColAll As New Collection
    Dim obj号序信息 As 号序信息
    Dim lngRow As Long, lngCol As Long, lngCount As Long
    Dim dtOriginalStartTime As Date
    
    Err = 0: On Error GoTo ErrHandler
    m_IsDataChanged = True
    mobj号序信息集.限号数 = lng限号数
    If mblnClickedFunBtn Then
        Call cmdFun_Click(mintPreSelFun)
        Exit Sub
    End If
    If blnAdd Then  '加号
        If mobj号序信息集.Count > 0 And mobj号序信息集.是否分时段 And mobj号序信息集.是否序号控制 Then

            intStep = DateDiff("n", mobj号序信息集(mobj号序信息集.Count).开始时间, mobj号序信息集(mobj号序信息集.Count).终止时间)
            dtStart = mobj号序信息集(mobj号序信息集.Count).终止时间
            dtOriginalStartTime = Format(dtStart, "yyyy-mm-dd ") & Format(mobj上班时段.开始时间, "hh:mm:ss")
            dtEnd = Format(dtStart, "yyyy-mm-dd ") & Format(mobj上班时段.结束时间, "hh:mm:ss")
            If DateDiff("n", dtEnd, dtStart) > 0 Then dtEnd = DateAdd("d", 1, dtEnd)
            
            '减去预留时间
            Call 减去预留时间(dtStart, dtEnd, mobj上班时段.出诊预留时间, mobj上班时段.休息时段)
            If DateDiff("n", dtStart, dtEnd) > 0 Then
                For Each obj号序信息 In mobj号序信息集
                    objColAll.Add Array(obj号序信息.序号, obj号序信息.开始时间, obj号序信息.终止时间, IIf(obj号序信息.是否预约, 1, 0))
                Next
                Set colNumAll = CalculatTimeInterval(0, mobj号序信息集.是否序号控制, intStep, objColAll.Count + lngCurAdd, _
                    dtStart, dtEnd, mobj上班时段.休息时段, objColAll.Count + 1, , Format(dtOriginalStartTime, "yyyy-MM-dd hh:mm:ss"))
                AddRange objColAll, colNumAll
                ShowTimeIntervals mobj号序信息集.是否序号控制, objColAll
                
                lngCount = mobj号序信息集.Count
                For lngRow = 0 To vsTimeWork.Rows - 1 Step 2
                    For lngCol = 1 To vsTimeWork.Cols - 1 Step 1


                        If vsTimeWork.TextMatrix(lngRow, lngCol) <> "" Then
                            If lngCount <= 0 Then
                                vsTimeWork.Cell(flexcpForeColor, lngRow, lngCol, lngRow + 1, lngCol) = vbMagenta
                            End If
                            lngCount = lngCount - 1
                        End If
                    Next
                Next
            End If
        End If
    Else  '减号
        If mobj号序信息集.Count > 0 And mobj号序信息集.是否分时段 And mobj号序信息集.是否序号控制 Then
            For Each obj号序信息 In mobj号序信息集
                objColAll.Add Array(obj号序信息.序号, obj号序信息.开始时间, obj号序信息.终止时间, IIf(obj号序信息.是否预约, 1, 0))
            Next
            ShowTimeIntervals mobj号序信息集.是否序号控制, objColAll
            
            lngCount = lng限号数
            For lngRow = 0 To vsTimeWork.Rows - 1 Step 2
                For lngCol = 1 To vsTimeWork.Cols - 1 Step 1


                    If vsTimeWork.TextMatrix(lngRow, lngCol) <> "" Then
                        If lngCount <= 0 Then
                            If ValiedCanModify(Val(vsTimeWork.TextMatrix(lngRow, lngCol)), 0, True) Then
                                vsTimeWork.Cell(flexcpForeColor, lngRow, lngCol, lngRow + 1, lngCol) = vbRed
                                vsTimeWork.Cell(flexcpFontStrikethru, lngRow, lngCol, lngRow + 1, lngCol) = True
                            End If
                        End If
                        lngCount = lngCount - 1
                    End If
                Next
            Next
        End If
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function 减去预留时间(ByRef dtStartDate As Date, ByRef dtEndDate As Date, _
    ByVal lng预留时间 As Long, ByVal str休息时段 As String) As Boolean
    '减去预留时间
    Dim dtStart As Date, dtEnd As Date
    Dim i As Integer, varTemp As Variant
    
    Err = 0: On Error GoTo ErrHandler
    dtEndDate = DateAdd("n", -1 * lng预留时间, dtEndDate)
    varTemp = Split(str休息时段, ";")
    For i = 0 To UBound(varTemp)
        '如果休息时段的开始时间小于上班时段的开始时间，则表示是第二天，休息时段的开始时间和终止时间都要加一天
        dtStart = CDate(Format(dtStartDate, "yyyy-mm-dd ") & Split(varTemp(i), "-")(0))
        dtEnd = CDate(Format(dtStartDate, "yyyy-mm-dd ") & Split(varTemp(i), "-")(1))
        If DateDiff("n", dtStart, dtStartDate) > 0 Then dtStart = DateAdd("d", 1, dtStart): dtEnd = DateAdd("d", 1, dtEnd)
        '休息时段的终止时间小于休息时段的开始时间，则休息时段的终止时间加一天
        If DateDiff("n", dtEnd, dtStart) > 0 Then dtEnd = DateAdd("d", 1, dtEnd)
        '如果上班时段的终止时间在休息时段内，则上班时段的终止时间取休息时段的开始时间
        If DateDiff("n", dtEndDate, dtEnd) <= 0 And DateDiff("n", dtEndDate, dtStart) >= 0 Then dtEndDate = dtStart
    Next
    减去预留时间 = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function ValiedCanModify(ByVal lng序号 As Long, ByVal lng数量 As Long, _
    Optional ByVal blnDel As Boolean, Optional ByRef lng已预约 As Long) As Boolean
    '检查当前时段是否可以修改
    Dim i As Long, arrSN As Variant
    
    Err = 0: On Error GoTo ErrHandler
    If mcllFixedSN Is Nothing Then ValiedCanModify = True: Exit Function
    If mcllFixedSN.Count = 0 Then ValiedCanModify = True: Exit Function
    
    For i = 1 To mcllFixedSN.Count
        arrSN = mcllFixedSN(i) '(序号,数量)
        If blnDel Then
            '删除主要是启用时段，且启用序号控制
            If arrSN(0) >= lng序号 And lng数量 < arrSN(1) Then
                lng已预约 = arrSN(1)
                Exit Function
            End If
        ElseIf arrSN(0) = lng序号 And lng数量 < arrSN(1) Then
            lng已预约 = arrSN(1)
            Exit Function
        End If
    Next
    ValiedCanModify = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

'注意！不要删除或修改下列被注释的行！
'MemberInfo=1,0,0,0
Public Property Get EditMode() As gRegistPlanEditMode
    EditMode = m_EditMode
End Property

Public Property Let EditMode(ByVal New_EditMode As gRegistPlanEditMode)
    m_EditMode = New_EditMode
    PropertyChanged "EditMode"
    
    picFunBack.Visible = m_CanReCalic And (EditMode = ED_RegistPlan_Edit Or m_EditMode = ED_RegistPlan_NumLimitModify)
    Call UserControl_Resize
    SetEnabled UserControl.Controls, m_EditMode = ED_RegistPlan_Edit Or m_EditMode = ED_RegistPlan_NumLimitModify
    SetEnabledBackColor UserControl.Controls
    
    vsTimeWork.Editable = flexEDNone
    If mobj号序信息集 Is Nothing Then Exit Property
    SetFunVisible mobj号序信息集.是否序号控制
    If mobj号序信息集.是否序号控制 = False And mobj号序信息集.是否分时段 Then
        vsTimeWork.Editable = IIf(m_EditMode = ED_RegistPlan_Edit Or m_EditMode = ED_RegistPlan_NumLimitModify, flexEDKbdMouse, flexEDNone)
        vsTimeWork.Editable = IIf(vsTimeWork.Editable = flexEDKbdMouse And mobj号序信息集.预约控制 <> 1, flexEDKbdMouse, flexEDNone)
    End If
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,false
Public Property Get IsDataChanged() As Boolean
    IsDataChanged = m_IsDataChanged
End Property

Public Property Let IsDataChanged(ByVal New_IsDataChanged As Boolean)
    m_IsDataChanged = New_IsDataChanged
    PropertyChanged "IsDataChanged"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=7,0,0,5
Public Property Get 诊疗频次() As Integer
    诊疗频次 = m_诊疗频次
End Property

Public Property Let 诊疗频次(ByVal New_诊疗频次 As Integer)
    m_诊疗频次 = New_诊疗频次
    PropertyChanged "诊疗频次"
    txtUpd.Text = IIf(m_诊疗频次 = 0, m_def_诊疗频次, m_诊疗频次)
End Property

