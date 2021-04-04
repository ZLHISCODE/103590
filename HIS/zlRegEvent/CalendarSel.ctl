VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#16.3#0"; "Codejock.Calendar.v16.3.1.ocx"
Begin VB.UserControl CalendarSel 
   ClientHeight    =   6210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10560
   ScaleHeight     =   6210
   ScaleWidth      =   10560
   Begin VB.PictureBox picWeekList 
      BorderStyle     =   0  'None
      Height          =   4065
      Left            =   4380
      ScaleHeight     =   4065
      ScaleWidth      =   4665
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   60
      Visible         =   0   'False
      Width           =   4665
      Begin zl9RegEvent.DayCalendar btnDay 
         Height          =   1935
         Left            =   2460
         TabIndex        =   19
         ToolTipText     =   "按住Ctrl键可一次选择多个未设置安排的日期"
         Top             =   2040
         Visible         =   0   'False
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   3413
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlignmentY      =   1
         MultiSelectionMode=   -1  'True
         BorderStyle     =   0   'False
      End
      Begin zl9RegEvent.WeekCalendar btnWeek 
         Height          =   1365
         Left            =   60
         TabIndex        =   17
         Top             =   2280
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   2408
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlignmentX      =   1
         AlignmentY      =   1
         BorderStyle     =   0   'False
      End
      Begin VB.CheckBox chkValied 
         Caption         =   "周六不出诊"
         Height          =   285
         Index           =   0
         Left            =   75
         TabIndex        =   7
         Top             =   690
         Width           =   1305
      End
      Begin VB.CheckBox chkValied 
         Caption         =   "周日不出诊"
         Height          =   285
         Index           =   1
         Left            =   1485
         TabIndex        =   8
         Top             =   690
         Width           =   1305
      End
      Begin VB.PictureBox picLoop 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   60
         ScaleHeight     =   975
         ScaleWidth      =   3825
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1050
         Width           =   3825
         Begin VB.OptionButton optLoop 
            Caption         =   "不限制"
            Height          =   225
            Index           =   0
            Left            =   0
            TabIndex        =   9
            Top             =   -15
            Width           =   870
         End
         Begin VB.OptionButton optLoop 
            Caption         =   "月内轮循"
            Height          =   225
            Index           =   1
            Left            =   1395
            TabIndex        =   10
            Top             =   -15
            Value           =   -1  'True
            Width           =   1110
         End
         Begin VB.TextBox txtSkip 
            Height          =   285
            Left            =   405
            TabIndex        =   15
            Text            =   "3"
            Top             =   630
            Width           =   390
         End
         Begin MSComCtl2.UpDown updSkip 
            Height          =   285
            Left            =   765
            TabIndex        =   16
            Top             =   630
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "txtSkip"
            BuddyDispid     =   196615
            OrigLeft        =   750
            OrigTop         =   630
            OrigRight       =   1005
            OrigBottom      =   915
            Max             =   100
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.DTPicker dtpLoopDay 
            Height          =   300
            Left            =   1185
            TabIndex        =   13
            Top             =   255
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   158728195
            CurrentDate     =   42377
         End
         Begin VB.ComboBox cboDays 
            Height          =   300
            Left            =   1185
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   255
            Width           =   1260
         End
         Begin VB.Label lblLoopDate 
            AutoSize        =   -1  'True
            Caption         =   "开始轮循日期"
            Height          =   180
            Left            =   0
            TabIndex        =   11
            Top             =   315
            Width           =   1080
         End
         Begin VB.Label lblLoopSkipDays 
            AutoSize        =   -1  'True
            Caption         =   "间隔        天"
            Height          =   180
            Left            =   0
            TabIndex        =   14
            Top             =   675
            Width           =   1260
         End
      End
      Begin VB.OptionButton optRule 
         Caption         =   "单日"
         Height          =   240
         Index           =   1
         Left            =   900
         TabIndex        =   3
         Top             =   0
         Width           =   660
      End
      Begin VB.OptionButton optRule 
         Caption         =   "星期"
         Height          =   240
         Index           =   0
         Left            =   60
         TabIndex        =   2
         Top             =   0
         Value           =   -1  'True
         Width           =   810
      End
      Begin VB.OptionButton optRule 
         Caption         =   "特定日期"
         Height          =   240
         Index           =   4
         Left            =   900
         TabIndex        =   6
         Top             =   270
         Width           =   1050
      End
      Begin VB.OptionButton optRule 
         Caption         =   "双日"
         Height          =   240
         Index           =   2
         Left            =   1890
         TabIndex        =   4
         Top             =   0
         Width           =   660
      End
      Begin VB.OptionButton optRule 
         Caption         =   "轮循"
         Height          =   240
         Index           =   3
         Left            =   60
         TabIndex        =   5
         Top             =   270
         Width           =   810
      End
      Begin VB.Line lnDown 
         BorderColor     =   &H80000005&
         X1              =   0
         X2              =   4425
         Y1              =   585
         Y2              =   585
      End
      Begin VB.Line lnTop 
         BorderColor     =   &H80000000&
         X1              =   -210
         X2              =   4605
         Y1              =   570
         Y2              =   570
      End
   End
   Begin XtremeCalendarControl.DatePicker dtpDays 
      Height          =   3030
      Left            =   15
      TabIndex        =   0
      Top             =   60
      Width           =   4245
      _Version        =   1048579
      _ExtentX        =   7488
      _ExtentY        =   5345
      _StockProps     =   64
      AutoSize        =   0   'False
      ShowTodayButton =   0   'False
      ShowNoneButton  =   0   'False
      HighlightToday  =   0   'False
      ShowNonMonthDays=   0   'False
      Show3DBorder    =   0
      MaxSelectionCount=   1
      AskDayMetrics   =   -1  'True
   End
End
Attribute VB_Name = "CalendarSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum gShowStyle
    Show_Plan_Day = 0
    Show_Plan_Week = 1
    Show_Plan_Rule = 2
End Enum

'缺省属性值:
Const m_def_BackColor = 0
Const m_def_ForeColor = 0
Const m_def_Enabled = 0
Const m_def_BackStyle = 0
Const m_def_BorderStyle = 0
Const m_def_ShowStyle = 0
'属性变量:
Dim m_BackColor As Long
Dim m_ForeColor As Long
Dim m_Enabled As Boolean
Dim m_Font As Font
Dim m_BackStyle As Integer
Dim m_BorderStyle As Integer
Dim m_ShowStyle As gShowStyle
'事件声明:
Public Event SelectedChangeBefore(ByVal OldDate As String, NewDate As String, Cancel As Boolean)
Public Event SelectedChanged(ByVal OldDate As String, NewDate As String)

Private Const mDateBefore As String = "2016-01-"
Private mdtMinDay As Date
Private mobj出诊安排 As 出诊安排

Private mblnNotClick As Boolean
Private mbytKeyShift As Byte '1-vbShiftMask 2-VbCtrlMask 4-vbAltMask
Private mcolPicture As Collection '图片集合，标记节假日和停诊安排的。
                                '集合关键字(Value取日期，格式"yyyymmdd")：
                                '   K+Value - 未选择的且未设置安排的, KS+Value - 当前选择的且未设置安排的
                                '   B+Value - 未选择的且设置了安排的, BS+Value - 当前选择的且设置了安排的

Public Function LoadData(ByRef obj出诊安排 As 出诊安排) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载出诊安排
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    Set mobj出诊安排 = obj出诊安排
    If mobj出诊安排 Is Nothing Then Set mobj出诊安排 = New 出诊安排
    
    LoadData = InitData
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub SetBoldItem(ByVal blnWeek As Boolean)
    Dim i As Integer
    
    If blnWeek Then
        For i = 0 To btnWeek.Count - 1
            btnWeek.ItemBold(i) = False
        Next
    Else
        For i = 0 To btnDay.Count - 1
            btnDay.ItemBold(i) = False
        Next
    End If
    
    If Not mobj出诊安排.已保存出诊安排 Is Nothing Then
        If mobj出诊安排.已保存出诊安排.排班规则 = IIf(blnWeek, 1, 6) Or m_ShowStyle = Show_Plan_Week Then
            If mobj出诊安排.已保存出诊安排.排班规则 = mobj出诊安排.排班规则 Then
                For i = 1 To mobj出诊安排.已保存出诊安排.Count
                    If blnWeek Then
                        If mobj出诊安排.已保存出诊安排(i).是否删除 = False Then
                            btnWeek.ItemBold(GetWeekIndex(mobj出诊安排.已保存出诊安排(i).出诊日期)) = True
                        End If
                    Else
                        btnDay.ItemBold(Val(mobj出诊安排.已保存出诊安排(i).出诊日期) - 1) = True
                    End If
                Next
            End If
        End If
    End If
    If Not mobj出诊安排.未保存出诊安排 Is Nothing Then
        For i = 1 To mobj出诊安排.未保存出诊安排.Count
            If blnWeek Then
                btnWeek.ItemBold(GetWeekIndex(mobj出诊安排.未保存出诊安排(i).出诊日期)) = True
            Else
                btnDay.ItemBold(Val(mobj出诊安排.未保存出诊安排(i).出诊日期) - 1) = True
            End If
        Next
    End If
End Sub

Private Function InitData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化数据
    '编制:刘兴洪
    '日期:2016-01-12 15:36:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim obj出诊记录集 As 出诊记录集
    
    Err = 0: On Error GoTo Errhand:
    
    If m_ShowStyle = Show_Plan_Day Then
        Call AdjustWeekSeat
        dtpDays.SetRange Format(mobj出诊安排.开始时间, "yyyy-mm-dd"), _
            Format(mobj出诊安排.终止时间, "yyyy-mm-dd")
        mdtMinDay = Format(mobj出诊安排.开始时间, "yyyy-mm-dd")
    End If
    
    btnWeek.Tag = ""
    If mobj出诊安排.更新合作单位 Then
        Enabled = False '不允许调整规则
        If m_ShowStyle = Show_Plan_Rule Then
            '不允许多选
            btnDay.MultiSelectionMode = False
            btnDay.ToolTipText = ""
        End If
    End If
    
    mblnNotClick = True
    dtpDays.ClearSelection
    dtpDays.RedrawControl
    
    optRule(0).Value = True
    optLoop(1).Value = True
    cboDays.ListIndex = 0
    dtpLoopDay.Value = Now
    txtSkip.Text = "1"
    
    btnWeek.ClearAll
    btnDay.ClearAll
    
    If mobj出诊安排.Count = 0 Then
        '新增出诊记录集
         Set obj出诊记录集 = New 出诊记录集
        '设置缺省值
        If m_ShowStyle = Show_Plan_Rule And mobj出诊安排.排班规则 <> 1 Then
            mobj出诊安排.排班规则 = 1
            mobj出诊安排.缺省出诊日期 = "周一"
        End If
        With obj出诊记录集
            .出诊日期 = mobj出诊安排.缺省出诊日期
        End With
        mobj出诊安排.AddItem obj出诊记录集, GetPlanKey(obj出诊记录集.出诊日期)
    End If
    
    '加载数据
    Select Case m_ShowStyle
    Case Show_Plan_Rule
        '1-星期排班;2-单日排班;3-双日排班;4-月内轮循;5-轮循不限制;6-特定日期
        Select Case mobj出诊安排.排班规则
        Case 1
            Call SetBoldItem(True)
            If mobj出诊安排(1).出诊日期 = "" Then
                btnWeek.ItemValue(0) = True
                mobj出诊安排(1).出诊日期 = GetWeekName(0)
            Else
                btnWeek.ItemValue(GetWeekIndex(mobj出诊安排(1).出诊日期)) = True
            End If
        Case 2, 3
            optRule(IIf(mobj出诊安排.排班规则 = 2, 1, 2)).Value = True
            '周六不出诊
            If mobj出诊安排.周六不出诊 Then chkValied(0).Value = vbChecked
            '周日不出诊
            If mobj出诊安排.周日不出诊 Then chkValied(1).Value = vbChecked
        Case 4, 5
            optRule(3).Value = True
            If mobj出诊安排.排班规则 = 4 Then
                optLoop(1).Value = True
                zlControl.CboLocate cboDays, Val(Format(mobj出诊安排.开始时间, "dd")), True
                txtSkip.Text = Val(mobj出诊安排(1).出诊日期)
            Else
                optLoop(0).Value = True
                dtpLoopDay.Value = Format(mobj出诊安排.开始时间, "yyyy-mm-dd")
            End If
            txtSkip.Text = Val(mobj出诊安排(1).出诊日期)
            '周六不出诊
            If mobj出诊安排.周六不出诊 Then chkValied(0).Value = vbChecked
            '周日不出诊
            If mobj出诊安排.周日不出诊 Then chkValied(1).Value = vbChecked
            
            dtpLoopDay.Visible = optLoop(0).Value And m_ShowStyle = Show_Plan_Rule
            cboDays.Visible = Not optLoop(0).Value And m_ShowStyle = Show_Plan_Rule
        Case 6
            optRule(4).Value = True
            Call SetBoldItem(False)
            If mobj出诊安排(1).出诊日期 = "" Then
                btnDay.ItemValue(0) = True
                mobj出诊安排(1).出诊日期 = btnDay.ItemCaption(0) & "日"
            Else
                btnDay.ItemValue(Val(mobj出诊安排(1).出诊日期) - 1) = True
            End If
        End Select
    Case Show_Plan_Week
        Call SetBoldItem(True)
        If mobj出诊安排(1).出诊日期 = "" Then
            btnWeek.ItemValue(0) = True
            mobj出诊安排(1).出诊日期 = GetWeekName(0)
        Else
            btnWeek.ItemValue(GetWeekIndex(mobj出诊安排(1).出诊日期)) = True
        End If
    Case Show_Plan_Day
        If mobj出诊安排(1).出诊日期 = "" Then
            mobj出诊安排(1).出诊日期 = mobj出诊安排.开始时间
            dtpDays.Select CDate(mobj出诊安排(1).出诊日期)
            dtpDays.EnsureVisibleSelection
            dtpDays.RedrawControl
            Call dtpDays_SelectionChanged
        ElseIf IsDate(mobj出诊安排(1).出诊日期) Then
            dtpDays.Select CDate(mobj出诊安排(1).出诊日期)
            dtpDays.EnsureVisibleSelection
            dtpDays.RedrawControl
            Call dtpDays_SelectionChanged
        End If
    End Select
    mblnNotClick = False
    InitData = True
    Exit Function
Errhand:
    mblnNotClick = False
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub AdjustWeekSeat()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调整星期的位置
    '编制:刘兴洪
    '日期:2016-01-08 15:18:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, sngTop As Single, sngLeft As Single
    Dim sngRuleTop As Single, sngRuleLeft As Single

    Err = 0: On Error GoTo Errhand:
    sngRuleTop = 0: sngRuleLeft = 0
    Select Case m_ShowStyle
    Case Show_Plan_Rule '规则
        picWeekList.Visible = True
        For i = 0 To optLoop.UBound
            optLoop(i).Visible = True
        Next
        For i = 0 To optRule.UBound
            optRule(i).Visible = True
        Next
        sngRuleTop = optRule(4).Top + optRule(4).Height + 100
        sngRuleLeft = optRule(0).Left
        btnWeek.Visible = optRule(0).Value
        chkValied(0).Visible = Not optRule(0).Value And Not optRule(4).Value
        chkValied(1).Visible = Not optRule(0).Value And Not optRule(4).Value
        picLoop.Visible = optRule(3).Value
        dtpLoopDay.Visible = optLoop(0).Value
        picLoop.Top = chkValied(1).Top + chkValied(1).Height + 50
        cboDays.Visible = Not optLoop(0).Value
        lnTop.Visible = True
        lnDown.Visible = True
        dtpDays.Visible = False
        btnDay.Visible = optRule(4).Value
    Case Show_Plan_Day  '按日期
        btnWeek.Visible = False
        picWeekList.Visible = False
        Set dtpDays.Container = Me '放在控件中
        dtpDays.Visible = True
        dtpDays.Left = ScaleLeft
        dtpDays.Width = ScaleWidth
        If Not mobj出诊安排 Is Nothing Then
            If mobj出诊安排.模板类型 = 2 Then
                dtpDays.Top = -550
            Else
                dtpDays.Top = ScaleTop
            End If
        Else
            dtpDays.Top = ScaleTop
        End If
        dtpDays.Height = ScaleHeight - dtpDays.Top
    Case Show_Plan_Week  '按星期
        For i = 0 To optLoop.UBound
            optLoop(i).Visible = False
        Next
        For i = 0 To optRule.UBound
            optRule(i).Visible = False
        Next
        picWeekList.Visible = True
        btnWeek.Visible = True
        dtpDays.Visible = False
        chkValied(0).Visible = False
        chkValied(1).Visible = False
        picLoop.Visible = False
        dtpLoopDay.Visible = False
        lnTop.Visible = False
        lnDown.Visible = False
    End Select
    picLoop.Top = chkValied(0).Top + chkValied(0).Height + 50
    picLoop.Left = chkValied(0).Left
    Call picWeekList_Resize
    Call UserControl_Resize
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitFace()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化界面
    '编制:刘兴洪
    '日期:2016-01-08 14:44:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    
    On Error GoTo Errhand
    dtpDays.HighlightToday = False
    dtpDays.ShowNonMonthDays = False
    With dtpDays.PaintManager
        .ControlBackColor = &HFFEFE3
        .DayBackColor = &HFFEFE3
        .DaysOfWeekBackColor = &HFFEFE3
    End With
    
    With cboDays
        .Clear
        For i = 1 To 30
            .AddItem "第" & i & "天"
            .ItemData(.NewIndex) = i
        Next
        mblnNotClick = True
        .ListIndex = 0
        mblnNotClick = False
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub btnWeek_Click(idxWeek As Integer, Value As Boolean)
    Dim i As Integer
    Dim blnCancel As Boolean
    Dim strOldDate As String, strNewDate As String
    
    On Error GoTo Errhand
    If btnWeek.Tag = "" Then btnWeek.Tag = idxWeek
    If mblnNotClick Then Exit Sub
    
    '至少选择一个
    If Value = False Then
        btnWeek.ItemValue(idxWeek) = True
    End If
    
    strOldDate = GetWeekName(Val(btnWeek.Tag)): strNewDate = GetWeekName(idxWeek)
    If btnWeek.Tag = "" Or Val(btnWeek.Tag) <> idxWeek Then
        RaiseEvent SelectedChangeBefore(strOldDate, strNewDate, blnCancel)
        If blnCancel Then
            mblnNotClick = True
            btnWeek.ItemValue(idxWeek) = False
            btnWeek.ItemValue(Val(btnWeek.Tag)) = True
            mblnNotClick = False
            Exit Sub
        End If
        btnWeek.Tag = idxWeek
        Call ChangeCurPlan(mobj出诊安排, GetWeekName(idxWeek))
        RaiseEvent SelectedChanged(strOldDate, strNewDate)
        Call SetBoldItem(True)
    End If
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub chkValied_Click(index As Integer)
    On Error GoTo Errhand
    If mblnNotClick Then Exit Sub
    If index = 0 Then '周六不出诊
        mobj出诊安排.周六不出诊 = (chkValied(index).Value = vbChecked)
    Else '周日不出诊
        mobj出诊安排.周日不出诊 = (chkValied(index).Value = vbChecked)
    End If
    '手动调整为已修改，因为虽然规则改变了，但是记录集可能还是一样
    If mobj出诊安排.Count > 0 Then mobj出诊安排(1).是否修改 = True
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cboDays_Click()
    Dim blnCancel As Boolean
    
    On Error GoTo Errhand
    If mblnNotClick Then Exit Sub
    mobj出诊安排.开始时间 = CDate(mDateBefore & cboDays.ItemData(cboDays.ListIndex))
    '手动调整为已修改，因为虽然规则改变了，但是记录集可能还是一样
    If mobj出诊安排.Count > 0 Then mobj出诊安排(1).是否修改 = True
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub dtpDays_DayMetrics(ByVal Day As Date, ByVal Metrics As XtremeCalendarControl.IDatePickerDayMetrics)
    Dim blnSelected As Boolean, strKey As String
    Dim strStopTxt As String, blnNotSelect As Boolean
    
    On Error GoTo Errhand
    If mobj出诊安排 Is Nothing Then Exit Sub
    If mobj出诊安排.Count = 0 Then Exit Sub
    
    If DateDiff("d", mobj出诊安排.开始时间, Day) >= 0 And DateDiff("d", Day, mobj出诊安排.终止时间) >= 0 Then
        Metrics.ForeColor = vbBlack
    Else
        '当前出诊安排时间范围外的用灰色显示，不能选择
        Metrics.ForeColor = &HC0C0C0
        blnNotSelect = True
    End If
    
    '已保存的加粗显示
    If CurDayIsSavedPlan(Day) Then
        strStopTxt = CurDayIsNotVisit(Day)
        If strStopTxt <> "" Then
            If IsDate(mobj出诊安排(1).出诊日期) Then blnSelected = DateDiff("d", Day, mobj出诊安排(1).出诊日期) = 0
            strKey = "B" & IIf(blnSelected, "S", "") & Format(Day, "yyyymmdd")
            If CollExitsValue(mcolPicture, strKey) = False Then Call AddPictureToColl(Day, strStopTxt, blnSelected, True, blnNotSelect)
            If CollExitsValue(mcolPicture, strKey) Then Set Metrics.Picture = mcolPicture(strKey)
        Else
            Metrics.Font.Bold = True
        End If
    Else
        strStopTxt = CurDayIsNotVisit(Day)
        If strStopTxt <> "" Then
            If IsDate(mobj出诊安排(1).出诊日期) Then blnSelected = DateDiff("d", Day, mobj出诊安排(1).出诊日期) = 0
            strKey = "K" & IIf(blnSelected, "S", "") & Format(Day, "yyyymmdd")
            If CollExitsValue(mcolPicture, strKey) = False Then Call AddPictureToColl(Day, strStopTxt, blnSelected, False, blnNotSelect)
            If CollExitsValue(mcolPicture, strKey) Then Set Metrics.Picture = mcolPicture(strKey)
        Else
            Metrics.Font.Bold = False
        End If
    End If
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function CurDayIsSavedPlan(ByVal Day As Date) As Boolean
    Dim i As Integer
    
    '当前日期是否保存了安排
    With mobj出诊安排
        If Not .已保存出诊安排 Is Nothing Then
            For i = 1 To .已保存出诊安排.Count
                If .已保存出诊安排(i).是否删除 = False Then
                    If IsDate(.已保存出诊安排(i).出诊日期) Then
                        If DateDiff("d", Day, .已保存出诊安排(i).出诊日期) = 0 Then
                            CurDayIsSavedPlan = True
                            Exit Function
                        End If
                    End If
                End If
            Next
        End If
        
        If Not .未保存出诊安排 Is Nothing Then
            For i = 1 To .未保存出诊安排.Count
                If IsDate(.未保存出诊安排(i).出诊日期) Then
                    If DateDiff("d", Day, .未保存出诊安排(i).出诊日期) = 0 Then
                        CurDayIsSavedPlan = True
                        Exit Function
                    End If
                End If
            Next
        End If
    End With
End Function

Private Function CurDayIsNotVisit(ByVal Day As Date) As String
    '当前日期是否节假日或有停止安排
    '返回：停止原因或节假日名称
    Dim i As Integer
    
    '当前日期是否保存了安排
    With mobj出诊安排
        If Not .停诊记录 Is Nothing Then
            For i = 1 To .停诊记录.Count
                If DateDiff("d", Day, .停诊记录(i).开始时间) <= 0 And DateDiff("d", Day, .停诊记录(i).终止时间) >= 0 Then
                    CurDayIsNotVisit = .停诊记录(i).停诊原因
                    Exit Function
                End If
            Next
        End If
    End With
End Function

Private Sub AddPictureToColl(ByVal Day As Date, ByVal strSubTxt As String, _
    ByVal blnSelected As Boolean, ByVal blnHavePlan As Boolean, ByVal blnNotSelect As Boolean)
    '添加图片到集合
    Dim strKey As String, strTxt As String
    Dim objFont As StdFont, objSubFont As StdFont
    Dim lngBackColor As Long, lngForeColor As Long
    Dim objPic As IPictureDisp

    'mcolPicture:图片集合，标记节假日和停诊安排的。
    '集合关键字(Value取日期，格式"yyyymmdd")：
    '   K+Value - 未选择的且未设置安排的, KS+Value - 当前选择的且未设置安排的
    '   B+Value - 未选择的且设置了安排的, BS+Value - 当前选择的且设置了安排的
    strKey = IIf(blnHavePlan, "B", "K") & IIf(blnSelected, "S", "") & Format(Day, "yyyymmdd")

    If mcolPicture Is Nothing Then Set mcolPicture = New Collection
    If CollExitsValue(mcolPicture, strKey) = False Then
        Set objFont = New StdFont
        objFont.Name = "宋体"
        objFont.Size = 9
        objFont.Bold = blnHavePlan
        
        Set objSubFont = New StdFont
        objSubFont.Size = 9
        
        lngBackColor = IIf(blnSelected, dtpDays.PaintManager.SelectedDayBackColor, dtpDays.PaintManager.DayBackColor)
        lngForeColor = IIf(blnNotSelect, &HC0C0C0, vbBlack)
        
        strTxt = Val(Format(Day, "dd"))
        strTxt = IIf(Len(strTxt) = 1, " ", "") & strTxt
        Set objPic = GetTempImage(strTxt, 400, 400, lngBackColor, lngForeColor, objFont, pictxtAlignCenterCenter, _
            Left(strSubTxt, 1), &H5B60F2, objSubFont, pictxtAlignLeftTop)
        If Not objPic Is Nothing Then mcolPicture.Add objPic, strKey
    End If
End Sub

Private Sub dtpDays_SelectionChanged()
    Dim dtCurDate As Date
    Dim blnCancel As Boolean
    Dim strOldDate As String, strNewDate As String
    
    On Error GoTo Errhand
    If mblnNotClick Then Exit Sub
    
    '最多只能选一个
    If dtpDays.Selection.BlocksCount > 0 Then
        dtCurDate = dtpDays.Selection.Blocks(0).DateBegin
    Else
        dtpDays.Select mdtMinDay
    End If
    
    strOldDate = mobj出诊安排(1).出诊日期: strNewDate = Format(dtCurDate, "yyyy-mm-dd")
    RaiseEvent SelectedChangeBefore(strOldDate, strNewDate, blnCancel)
    If blnCancel Then
        dtpDays.ClearSelection
        dtpDays.Select CDate(strOldDate)
        Exit Sub
    End If
    
    Call ChangeCurPlan(mobj出诊安排, strNewDate)
    RaiseEvent SelectedChanged(strOldDate, strNewDate)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub dtpLoopDay_Change()
    On Error GoTo Errhand
    If mblnNotClick Then Exit Sub
    mobj出诊安排.开始时间 = dtpLoopDay.Value
    '手动调整为已修改，因为虽然规则改变了，但是记录集可能还是一样
    If mobj出诊安排.Count > 0 Then mobj出诊安排(1).是否修改 = True
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub optLoop_Click(index As Integer)
    Dim blnCancel As Boolean
    Dim strOldDate As String, strNewDate As String
    
    On Error GoTo Errhand
    If mblnNotClick Then Exit Sub
    strOldDate = Val(txtSkip.Text) & "天": strNewDate = Val(txtSkip.Text) & "天"
    RaiseEvent SelectedChangeBefore(strOldDate, strNewDate, blnCancel)
    If blnCancel Then
        mblnNotClick = True
        optLoop(IIf(index = 0, 1, 0)).Value = True
        mblnNotClick = False
        Exit Sub
    End If
    
    dtpLoopDay.Visible = optLoop(0).Value And m_ShowStyle = Show_Plan_Rule
    cboDays.Visible = Not optLoop(0).Value And m_ShowStyle = Show_Plan_Rule
    If index = 0 Then
        mobj出诊安排.排班规则 = 5
        mobj出诊安排.开始时间 = dtpLoopDay.Value
    Else '月内轮询
        mobj出诊安排.排班规则 = 4
        mobj出诊安排.开始时间 = CDate(mDateBefore & cboDays.ItemData(cboDays.ListIndex))
    End If
    Call ChangeCurPlan(mobj出诊安排, strNewDate, True)
    '手动调整为已修改，因为虽然规则改变了，但是记录集可能还是一样
    If mobj出诊安排.Count > 0 Then mobj出诊安排(1).是否修改 = True
    RaiseEvent SelectedChanged(strOldDate, strNewDate)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

 Private Sub optRule_Click(index As Integer)
    Dim i As Integer, str出诊日期 As String
    Dim byt原排班规则 As Byte
    Dim blnCancel As Boolean
    
    On Error GoTo Errhand
    Call AdjustWeekSeat
    btnWeek.Tag = ""
    If mblnNotClick Then Exit Sub
    
    '排班规则:1-星期排班;2-单日排班;3-双日排班;4-月内轮循;5-轮循不限制;6-特定日期
    byt原排班规则 = mobj出诊安排.排班规则
    mblnNotClick = True
    Select Case index
    Case 0
        btnWeek.ClearAll
        btnWeek.ItemValue(0) = True
        mobj出诊安排.排班规则 = 1
        str出诊日期 = "周一"
        
        Call SetBoldItem(True)
    Case 1
        mobj出诊安排.排班规则 = 2
        str出诊日期 = "单日"
        '周六不出诊
        If mobj出诊安排.周六不出诊 Then chkValied(0).Value = vbChecked
        '周日不出诊
        If mobj出诊安排.周日不出诊 Then chkValied(1).Value = vbChecked
    Case 2
        mobj出诊安排.排班规则 = 3
        str出诊日期 = "双日"
        '周六不出诊
        If mobj出诊安排.周六不出诊 Then chkValied(0).Value = vbChecked
        '周日不出诊
        If mobj出诊安排.周日不出诊 Then chkValied(1).Value = vbChecked
    Case 3
        mobj出诊安排.排班规则 = 4
        txtSkip.Text = "1"
        optLoop(1).Value = True
        mobj出诊安排.开始时间 = CDate(mDateBefore & cboDays.ItemData(cboDays.ListIndex))
        str出诊日期 = Val(txtSkip.Text) & "天"
        '周六不出诊
        If mobj出诊安排.周六不出诊 Then chkValied(0).Value = vbChecked
        '周日不出诊
        If mobj出诊安排.周日不出诊 Then chkValied(1).Value = vbChecked
    Case 4
        mobj出诊安排.排班规则 = 6
        btnDay.ClearAll
        btnDay.ItemValue(0) = True
        str出诊日期 = 1 & "日"
        
        Call SetBoldItem(False)
    End Select
    Call ChangeCurPlan(mobj出诊安排, str出诊日期, byt原排班规则 <> mobj出诊安排.排班规则)
    RaiseEvent SelectedChanged(str出诊日期, str出诊日期)
    mblnNotClick = False
    Exit Sub
Errhand:
    mblnNotClick = False
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub picWeekList_Resize()
    Err = 0: On Error Resume Next
    With picWeekList
        chkValied(0).Top = optRule(4).Top + optRule(4).Height + 100
        chkValied(1).Top = chkValied(0).Top
        picLoop.Left = chkValied(0).Left
        picLoop.Width = .ScaleWidth
        picLoop.Top = chkValied(0).Top + chkValied(0).Height + 50
        lnTop.X1 = -100
        lnTop.X2 = .ScaleWidth + 200
        lnDown.X1 = lnTop.X1
        lnDown.X2 = lnTop.X2
        btnDay.Left = .ScaleLeft  ' chkValied(0).Left
        btnDay.Width = .ScaleWidth - btnDay.Left
        btnDay.Top = lnDown.Y1 + 50
        btnDay.Height = .ScaleHeight - btnDay.Top
        btnWeek.Left = .ScaleLeft
        btnWeek.Width = .ScaleWidth - btnWeek.Left * 2
        If m_ShowStyle = Show_Plan_Rule Then
            btnWeek.Top = chkValied(1).Top
            btnWeek.Height = .ScaleHeight - btnWeek.Top
        Else
            btnWeek.Top = .ScaleTop
            btnWeek.Height = .ScaleHeight
        End If
    End With
End Sub

Private Sub txtSkip_Change()
    Dim blnCancel As Boolean
    Dim strOldDate As String, strNewDate As String
    
    On Error GoTo Errhand
    If mblnNotClick Then Exit Sub
    strOldDate = mobj出诊安排(1).出诊日期: strNewDate = Val(txtSkip.Text) & "天"
    RaiseEvent SelectedChangeBefore(strOldDate, strNewDate, blnCancel)
    If blnCancel Then
        mblnNotClick = True
        txtSkip.Text = Val(strOldDate)
        mblnNotClick = False
        Exit Sub
    End If
    
    If optLoop(0).Value Then
        mobj出诊安排.排班规则 = 5
        mobj出诊安排.开始时间 = dtpLoopDay.Value
    Else '月内轮询
        mobj出诊安排.排班规则 = 4
        mobj出诊安排.开始时间 = CDate(mDateBefore & cboDays.ItemData(cboDays.ListIndex))
    End If
    Call ChangeCurPlan(mobj出诊安排, strNewDate, True)
    '手动调整为已修改，因为虽然规则改变了，但是记录集可能还是一样
    If mobj出诊安排.Count > 0 Then mobj出诊安排(1).是否修改 = True
    RaiseEvent SelectedChanged(strOldDate, strNewDate)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub txtSkip_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub UserControl_Initialize()
    Call InitFace
    Set mcolPicture = New Collection
    Set mobj出诊安排 = New 出诊安排
End Sub

Private Sub UserControl_Resize()
    Err = 0: On Error Resume Next
    With UserControl
        dtpDays.Left = .ScaleLeft - 300
        dtpDays.Width = .ScaleWidth + 600
        If Not mobj出诊安排 Is Nothing Then
            If mobj出诊安排.模板类型 = 2 Then
                dtpDays.Top = -550
            Else
                dtpDays.Top = .ScaleTop
            End If
        Else
            dtpDays.Top = .ScaleTop
        End If
        dtpDays.Height = .ScaleHeight - dtpDays.Top
        
        picWeekList.Left = .ScaleLeft
        picWeekList.Width = .ScaleWidth
        picWeekList.Top = .ScaleTop
        picWeekList.Height = .ScaleHeight
    End With
End Sub

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8,0,0,0
Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
    SetBackColor Controls, m_BackColor
    dtpDays.PaintManager.ListControlBackColor = m_BackColor
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8,0,0,0
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
    
    SetEnabled UserControl.Controls, New_Enabled
    btnWeek.Enabled = True
    btnDay.Enabled = True
    dtpDays.Enabled = True
    dtpDays.PaintManager.SelectedDayBackColor = IIf(dtpDays.Enabled, &HFFFFC0, vbRed)
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=6,0,0,0
Public Property Get Font() As Font
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    PropertyChanged "Font"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=7,0,0,0
Public Property Get BackStyle() As Integer
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=7,0,0,0
Public Property Get BorderStyle() As Integer
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=7,0,0,0
Public Property Get ShowStyle() As gShowStyle
    ShowStyle = m_ShowStyle
End Property

Public Property Let ShowStyle(ByVal New_ShowStyle As gShowStyle)
    m_ShowStyle = New_ShowStyle
    PropertyChanged "ShowStyle"
    Call AdjustWeekSeat
End Property

Public Property Let KeyShift(ByVal Shift As Byte)
    mbytKeyShift = Shift
End Property

'为用户控件初始化属性
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    m_Enabled = m_def_Enabled
    Set m_Font = Ambient.Font
    m_BackStyle = m_def_BackStyle
    m_BorderStyle = m_def_BorderStyle
    m_ShowStyle = m_def_ShowStyle
    Call AdjustWeekSeat
End Sub

'从存贮器中加载属性值
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_ShowStyle = PropBag.ReadProperty("ShowStyle", m_def_ShowStyle)
    
    Call AdjustWeekSeat
    SetBackColor Controls, m_BackColor
End Sub

Private Sub UserControl_Terminate()
    Set mcolPicture = Nothing
    Set mobj出诊安排 = Nothing
End Sub

'将属性值写到存储器
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("ShowStyle", m_ShowStyle, m_def_ShowStyle)
End Sub

Private Sub btnDay_Click(idxDay As Integer, Value As Boolean, Text As String)
    Dim strTemp As String, i As Integer
    Dim blnOldSaved As Boolean, blnNewSaved As Boolean
    Dim blnCancel As Boolean
    Dim strOldDate As String, strNewDate As String
    Dim obj出诊记录集 As 出诊记录集
    
    On Error GoTo Errhand
    If mblnNotClick Then Exit Sub
    strOldDate = mobj出诊安排(1).出诊日期: strNewDate = btnDay.ItemCaption(idxDay) & "日"
    
    mblnNotClick = True
    If Value = True Then '选择
        Set obj出诊记录集 = mobj出诊安排(1).Clone
        
        '上一个选择是否是已保存的
        If mobj出诊安排.已保存出诊安排.Exits("K" & strOldDate) Then
            If mobj出诊安排.已保存出诊安排("K" & strOldDate).是否删除 = False Then blnOldSaved = True
        End If
        If mobj出诊安排.未保存出诊安排.Exits("K" & strOldDate) Then blnOldSaved = True
        
        '本次选择是否是已保存的
        If mobj出诊安排.已保存出诊安排.Exits("K" & strNewDate) Then
            If mobj出诊安排.已保存出诊安排("K" & strNewDate).是否删除 = False Then blnNewSaved = True
        End If
        If mobj出诊安排.未保存出诊安排.Exits("K" & strNewDate) Then blnNewSaved = True
        
        RaiseEvent SelectedChangeBefore(strOldDate, strNewDate, blnCancel)
        If blnCancel Then
            btnDay.ItemValue(idxDay) = False
            mblnNotClick = False
            Exit Sub
        End If
        
        If blnOldSaved = False And blnNewSaved = False And mbytKeyShift = vbCtrlMask Then
            obj出诊记录集.出诊日期 = strNewDate
            mobj出诊安排.AddItem obj出诊记录集, "K" & strNewDate
        Else
            For i = 0 To btnDay.Count - 1
                btnDay.ItemValue(i) = False
            Next
            btnDay.ItemValue(idxDay) = True
            Call ChangeCurPlan(mobj出诊安排, strNewDate)
        End If
        RaiseEvent SelectedChanged(strOldDate, strNewDate)
        Call SetBoldItem(False)
    Else '取消选择
        If mobj出诊安排.已保存出诊安排.Exits("K" & strNewDate) Or mobj出诊安排.未保存出诊安排.Exits("K" & strNewDate) _
            Or mobj出诊安排.Count = 1 Then
            '选择项为已保存记录集或未保存记录集的元素，或当前选择只有一个时，则不允许取消选择
            btnDay.ItemValue(idxDay) = True
        Else
            '移除元素
            If mobj出诊安排.Exits("K" & strNewDate) Then mobj出诊安排.Remove "K" & strNewDate
            RaiseEvent SelectedChanged(strOldDate, strNewDate)
            Call SetBoldItem(False)
        End If
    End If
    mblnNotClick = False
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
