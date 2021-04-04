VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.UserControl UserDatePicker 
   ClientHeight    =   4695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8715
   ScaleHeight     =   4695
   ScaleWidth      =   8715
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   360
      ScaleHeight     =   4455
      ScaleWidth      =   7875
      TabIndex        =   0
      Top             =   120
      Width           =   7875
      Begin VSFlex8Ctl.VSFlexGrid vsfMonth 
         Height          =   2985
         Index           =   0
         Left            =   60
         TabIndex        =   1
         Top             =   540
         Width           =   3015
         _cx             =   5318
         _cy             =   5265
         Appearance      =   2
         BorderStyle     =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "微软雅黑"
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
      Begin VSFlex8Ctl.VSFlexGrid vsfMonth 
         Height          =   2955
         Index           =   1
         Left            =   3180
         TabIndex        =   2
         Top             =   540
         Width           =   3705
         _cx             =   6535
         _cy             =   5212
         Appearance      =   2
         BorderStyle     =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "微软雅黑"
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
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   2
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
      Begin XtremeSuiteControls.ShortcutCaption sccTitle 
         CausesValidation=   0   'False
         Height          =   315
         Index           =   1
         Left            =   3180
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   210
         Width           =   3015
         _Version        =   589884
         _ExtentX        =   5318
         _ExtentY        =   564
         _StockProps     =   6
         Caption         =   "2016年7月"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.14
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         Alignment       =   1
      End
      Begin XtremeSuiteControls.ShortcutCaption sccTitle 
         CausesValidation=   0   'False
         Height          =   315
         Index           =   0
         Left            =   60
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   210
         Width           =   3015
         _Version        =   589884
         _ExtentX        =   5318
         _ExtentY        =   564
         _StockProps     =   6
         Caption         =   "2016年6月"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.14
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         Alignment       =   1
      End
      Begin VB.Image imgWork 
         Height          =   210
         Left            =   2340
         Picture         =   "UserDatePicker.ctx":0000
         Top             =   3750
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Image imgHolidy 
         Height          =   210
         Left            =   1980
         Picture         =   "UserDatePicker.ctx":0562
         Top             =   3780
         Visible         =   0   'False
         Width           =   210
      End
   End
End
Attribute VB_Name = "UserDatePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_HolidayStart As Date
Private m_TitleBackColor As OLE_COLOR
Private m_WeekBackColor As OLE_COLOR

Private Const m_def_TitleBackColor = &HEAA064
Private Const m_def_WeekBackColor = vbButtonFace

'事件声明
Public Event DayMetrics(Day As Date, Metrics As UserDatePickerDayMetrics)

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "控件可用状态。"
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
    UserControl.Enabled = vNewValue
    PropertyChanged "Enabled"
End Property

Property Get HolidayStart() As Date
Attribute HolidayStart.VB_Description = "节假日开始时间，用于定位显示的月份。"
    HolidayStart = m_HolidayStart
End Property

Property Let HolidayStart(ByVal vNewValue As Date)
    m_HolidayStart = vNewValue
    PropertyChanged "HolidayStart"
    Call RedrawControl
End Property

Property Let TitleBackColor(ByVal vNewValue As OLE_COLOR)
Attribute TitleBackColor.VB_Description = "年月标题行背景颜色。"
    m_TitleBackColor = vNewValue
    PropertyChanged "TitleBackColor"
    '
    '
End Property

Property Get TitleBackColor() As OLE_COLOR
    TitleBackColor = m_TitleBackColor
End Property

Property Let WeekBackColor(ByVal vNewValue As OLE_COLOR)
Attribute WeekBackColor.VB_Description = "星期行背景颜色。"
    m_WeekBackColor = vNewValue
    PropertyChanged "WeekBackColor"
    vsfMonth(0).Cell(flexcpBackColor, 0, 0, 0, vsfMonth(0).Cols - 1) = vNewValue
    vsfMonth(1).Cell(flexcpBackColor, 0, 0, 0, vsfMonth(1).Cols - 1) = vNewValue
End Property

Property Get WeekBackColor() As OLE_COLOR
    WeekBackColor = m_WeekBackColor
End Property

Public Sub RedrawControl()
    Dim i As Integer, j As Integer, k As Integer
    Dim datFirstMouthStart As Date, datSecondMouthStart As Date
    
    '确定开始显示时间，开始时间在本月已过半时，显示上一月和本月，否则显示本月和下一月
    datFirstMouthStart = m_HolidayStart
    If Day(m_HolidayStart) <= 15 Then datFirstMouthStart = DateAdd("m", -1, m_HolidayStart)
    datFirstMouthStart = Format(datFirstMouthStart, "yyyy/mm/01")
    datSecondMouthStart = Format(DateAdd("m", 1, datFirstMouthStart), "yyyy/mm/01")
    
    Call InitGrid(vsfMonth(0), Format(datFirstMouthStart, "yyyy年mm月"))
    Call InitGrid(vsfMonth(1), Format(datSecondMouthStart, "yyyy年mm月"))
    
    Call InitData(vsfMonth(0), datFirstMouthStart)
    Call InitData(vsfMonth(1), datSecondMouthStart)
    
    Call RaiseDayMetricsEvent(vsfMonth(0))
    Call RaiseDayMetricsEvent(vsfMonth(1))
End Sub

Private Sub InitData(ByVal vsfGrid As VSFlexGrid, ByVal datStart As Date)
    Dim i As Integer, j As Integer, blnExit As Boolean
    Dim intFirstDayWeek As Integer, datCurrent As Date
    
    '计算第一天星期几
    'Weekday以vbMonday为一周的第一天，则返回值：1-星期一,2-星期二,3-星期三,4-星期四,5-星期五,6-星期六,7-星期日
    intFirstDayWeek = Weekday(datStart, vbMonday)
    datCurrent = datStart: blnExit = False
    With vsfGrid
        .Cell(flexcpText, 1, 0, .Rows - 1, .Cols - 1) = ""
        For i = 1 To .Rows - 1
            If blnExit Then Exit For
            For j = 0 To .Cols - 1
                If i = 1 Then
                    If j >= intFirstDayWeek - 1 Then
                        .TextMatrix(i, j) = Day(datCurrent)
                        datCurrent = DateAdd("d", 1, datCurrent)
                    End If
                Else
                    .TextMatrix(i, j) = Day(datCurrent)
                    If Month(DateAdd("d", 1, datCurrent)) > Month(datCurrent) _
                        Or Year(DateAdd("d", 1, datCurrent)) > Year(datCurrent) Then
                        blnExit = True: Exit For
                    End If
                    datCurrent = DateAdd("d", 1, datCurrent)
                End If
            Next
        Next
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
    End With
End Sub

Private Sub InitGrid(ByVal vsfGrid As VSFlexGrid, ByVal strTitle As String)
    Dim i As Integer, j As Integer
    
    With vsfGrid
        .Rows = 7: .Cols = 7
        .FixedRows = 2: .FixedCols = 0
        
        .GridLines = flexGridFlat
        .GridColor = .BackColor
        .GridLinesFixed = flexGridNone
        
        .BackColorFixed = .BackColor
        .BackColorBkg = .BackColor
        .SheetBorder = .BackColor
        
        .ScrollBars = flexScrollBarNone
        .HighLight = flexHighlightNever
        .FocusRect = flexFocusHeavy
        
        .MergeCellsFixed = flexMergeRestrictRows
        .MergeRow(0) = True
        
        sccTitle(vsfGrid.Index).Caption = strTitle
        .Cell(flexcpBackColor, 0, 0, 0, .Cols - 1) = vbButtonFace
        .CellBorderRange 0, 0, 0, .Cols - 1, &H80000000, 0, 0, 0, 1, 0, 0
        For j = 0 To .Cols - 1
            .TextMatrix(0, j) = Choose(j + 1, "一", "二", "三", "四", "五", "六", "日")
        Next
       
        .RowHeight(0) = 300
        .Cell(flexcpPictureAlignment, 1, 0, .Rows - 1, .Cols - 1) = flexAlignLeftTop
    End With
End Sub

Private Sub UserControl_Initialize()
    m_TitleBackColor = m_def_TitleBackColor
    m_WeekBackColor = m_def_WeekBackColor
    Call RedrawControl
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    HolidayStart = PropBag.ReadProperty("HolidayStart", Now)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    TitleBackColor = PropBag.ReadProperty("TitleBackColor", m_def_TitleBackColor)
    WeekBackColor = PropBag.ReadProperty("WeekBackColor", m_def_WeekBackColor)
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    picBack.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
End Sub

Private Sub picBack_Resize()
    Dim i As Integer

    On Error Resume Next
    sccTitle(0).Move 0, 0, picBack.ScaleWidth / 2
    With vsfMonth(0)
        .Left = 0
        .Top = sccTitle(0).Top + sccTitle(0).Height
        .Width = picBack.ScaleWidth / 2 - .Left
        .Height = picBack.ScaleHeight - .Top
        .ColWidth(-1) = .Width / .Cols
        For i = 1 To .Rows - 1
            .RowHeight(i) = (.Height - 800) / (.Rows - 1)
        Next
    End With
    sccTitle(1).Move vsfMonth(0).Width + 20, 0, picBack.ScaleWidth / 2
    With vsfMonth(1)
        .Left = vsfMonth(0).Width + 20
        .Top = sccTitle(1).Top + sccTitle(1).Height
        .Width = picBack.ScaleWidth - .Left
        .Height = picBack.ScaleHeight - .Top
        .ColWidth(-1) = .Width / .Cols - 2
        For i = 1 To .Rows - 1
            .RowHeight(i) = (.Height - 800) / (.Rows - 1)
        Next
    End With
End Sub

Private Sub RaiseDayMetricsEvent(ByVal vsf As VSFlexGrid)
    Dim objMetrics As UserDatePickerDayMetrics
    Dim datCurrent As Date
    Dim i As Integer, j As Integer
    
    On Error Resume Next
    '先初始化，还原所有的属性值
    With vsf
        .Cell(flexcpPicture, 1, 0, .Rows - 1, .Cols - 1) = Nothing
        .Cell(flexcpBackColor, 1, 0, .Rows - 1, .Cols - 1) = .BackColor
        .Cell(flexcpForeColor, 1, 0, .Rows - 1, .Cols - 1) = .ForeColor
        .Cell(flexcpFontBold, 1, 0, .Rows - 1, .Cols - 1) = .FontBold
        
        For i = 1 To .Rows - 1
            For j = 0 To .Cols - 1
                If .TextMatrix(i, j) <> "" Then
                    Set objMetrics = New UserDatePickerDayMetrics
                    datCurrent = Format(sccTitle(vsf.Index).Caption, "yyyy/mm/") & .TextMatrix(i, j)
                    RaiseEvent DayMetrics(datCurrent, objMetrics)
                    If objMetrics.BackColor <> 0 Then .Cell(flexcpBackColor, i, j) = objMetrics.BackColor
                    If objMetrics.ForeColor <> 0 Then .Cell(flexcpForeColor, i, j) = objMetrics.ForeColor
                    If objMetrics.FontBold Then .Cell(flexcpFontBold, i, j) = objMetrics.FontBold
                    If objMetrics.IsHoliday Then
                        .Cell(flexcpPicture, i, j) = imgHolidy.Picture
                    End If
                    If objMetrics.IsWorkFromHoliday Then
                        .Cell(flexcpPicture, i, j) = imgWork.Picture
                    End If
                End If
            Next
        Next
    End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("HolidayStart", m_HolidayStart, Now)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("TitleBackColor", m_TitleBackColor, m_def_TitleBackColor)
    Call PropBag.WriteProperty("WeekBackColor", m_WeekBackColor, m_def_WeekBackColor)
End Sub

Private Sub vsfMonth_BeforeRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If vsfMonth(Index).TextMatrix(NewRow, NewCol) = "" Then Cancel = True
End Sub

