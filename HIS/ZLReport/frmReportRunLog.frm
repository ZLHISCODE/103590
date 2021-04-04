VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmReportRunLog 
   Caption         =   "报表运行日志"
   ClientHeight    =   7170
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11010
   Icon            =   "frmReportRunLog.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7170
   ScaleWidth      =   11010
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtDoEvents 
      Height          =   375
      Left            =   -1500
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   4200
      Index           =   1
      Left            =   240
      ScaleHeight     =   4200
      ScaleWidth      =   6945
      TabIndex        =   0
      Top             =   1920
      Width           =   6945
      Begin VSFlex8Ctl.VSFlexGrid vsfLog 
         Height          =   1935
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   3180
         _cx             =   5609
         _cy             =   3413
         Appearance      =   0
         BorderStyle     =   0
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
         BackColorSel    =   16772055
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483638
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   270
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
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
   End
   Begin MSComCtl2.DTPicker dtp 
      Height          =   300
      Index           =   0
      Left            =   3960
      TabIndex        =   2
      Top             =   120
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   270794755
      CurrentDate     =   41694
   End
   Begin MSComCtl2.DTPicker dtp 
      Height          =   300
      Index           =   1
      Left            =   6510
      TabIndex        =   3
      Top             =   330
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   270794755
      CurrentDate     =   41694
   End
   Begin XtremeCommandBars.ImageManager imgPublic 
      Left            =   2280
      Top             =   360
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmReportRunLog.frx":6852
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmReportRunLog.frx":7148
      Left            =   960
      Top             =   240
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmReportRunLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobjMain As Object
Private mblnStartUp As Boolean
Private mlngReportRunKey As Long
Private mstrCaption As String

Private Function CommandBarInit(ByRef cbsMain As Object, Optional ByVal blnEnableCustomization As Boolean, Optional ByVal objIcons As ImageManagerIcons) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    cbsMain.VisualTheme = xtpThemeOffice2003
    
    With cbsMain.Options
        .ShowExpandButtonAlways = blnEnableCustomization
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '放在VisualTheme后有效
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 32, 32
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization blnEnableCustomization
    
    cbsMain.Options.LargeIcons = True
    
    CommandBarInit = True
    
End Function

Private Function NewToolBar(objBar As CommandBar, _
                                ByVal xtpType As XTPControlType, _
                                ByVal lngID As Long, _
                                ByVal strCaption As String, _
                                Optional ByVal blnBeginGroup As Boolean, _
                                Optional ByVal lngIcon As Long = -1, _
                                Optional ByVal bytStyle As Byte = xtpButtonIconAndCaption, _
                                Optional ByVal strToolTipText As String, _
                                Optional ByVal intBefore As Integer) As CommandBarControl
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objControl As CommandBarControl
    
    With objBar.Controls
        Set objControl = .Add(xtpType, lngID, strCaption, intBefore)
        objControl.id = lngID
        objControl.IconId = IIF(lngIcon = -1, lngID, lngIcon)
        objControl.BeginGroup = blnBeginGroup
        
        If strToolTipText <> "" Then objControl.ToolTipText = strToolTipText

        If objControl.type = xtpControlButton Or objControl.type = xtpControlPopup Then
            objControl.Style = bytStyle
        End If
        
    End With
    
    Set NewToolBar = objControl
    
End Function

Private Function DockPannelInit(ByRef dkpMain As Object) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.UseSplitterTracker = False '实时拖动
    dkpMain.Options.AlphaDockingContext = True
    dkpMain.Options.CloseGroupOnButtonClick = True
    dkpMain.Options.HideClient = True

    DockPannelInit = True
    
End Function

Private Function InitCommandBar() As Boolean
    '******************************************************************************************************************
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objFindKey As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom

    '------------------------------------------------------------------------------------------------------------------
    '初始设置
    Call CommandBarInit(cbsMain)
    Set cbsMain.Icons = imgPublic.Icons
    cbsMain.Options.LargeIcons = False

    '------------------------------------------------------------------------------------------------------------------
    '菜单定义:包括公共部份，请对xtpControlPopup类型的命令ID重新赋值

    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.Visible = False


    '------------------------------------------------------------------------------------------------------------------
    '工具栏定义:包括公共部份

    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched

    Set objControl = NewToolBar(objBar, xtpControlLabel, 1502, "开始时间", True)
    Set cbrCustom = NewToolBar(objBar, xtpControlCustom, 2003, "开始时间")
    cbrCustom.handle = dtp(0).hwnd
    
    Set objControl = NewToolBar(objBar, xtpControlLabel, 1503, "结束时间")
    Set cbrCustom = NewToolBar(objBar, xtpControlCustom, 2004, "结束时间")
    cbrCustom.handle = dtp(1).hwnd
    
    Set objControl = NewToolBar(objBar, xtpControlButton, 1005, "刷新")
    Set objControl = NewToolBar(objBar, xtpControlButton, 1006, "退出")
    
End Function

Private Sub InitDockPannel()
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************

    Dim objPane As Pane

    Set objPane = dkpMain.CreatePane(1, 100, 100, DockLeftOf, Nothing)
    objPane.Title = "日志列表"
    objPane.Options = PaneNoCaption
    
    dkpMain.SetCommandBars cbsMain
    Call DockPannelInit(dkpMain)

End Sub

Public Function ShowMe(ByVal objMain As Object, ByVal lngReportKey As Long, ByVal strCaption As String)
    mlngReportRunKey = lngReportKey
    mstrCaption = strCaption
    Me.Show 1, objMain
End Function

Private Function InitData()
    On Error GoTo ErrHand
    Me.Caption = mstrCaption
    With vsfLog
        .Rows = 2
        .Cols = 4
        .ColWidth(0) = 1350
        .ColWidth(1) = 2700
        .ColWidth(2) = 2700
        .ColWidth(3) = 1350
        .ExtendLastCol = True
        .AllowUserResizing = flexResizeColumns
        .TextMatrix(0, 0) = "执行人员"
        .TextMatrix(0, 1) = "执行开始时间"
        .TextMatrix(0, 2) = "执行结束时间"
        .TextMatrix(0, 3) = "执行耗时"
        .ColKey(0) = "执行人员"
        .ColKey(1) = "执行开始时间"
        .ColKey(2) = "执行结束时间"
        .ColKey(3) = "执行耗时"
        .BackColorSel = &H8000000D
    End With
    
    Call InitCommandBar
    Call InitDockPannel
    
    dtp(0).Value = Format(DateAdd("D", -30, Now), "YYYY-MM-DD 00:00:01")
    dtp(1).Value = Format(Now, "YYYY-MM-DD 23:59:59")
    
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function RefreshData() As Boolean
    On Error GoTo ErrHand
    
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim i As Integer
    
    strSQL = "Select a.Id,to_Char(a.执行开始时间,'YYYY-MM-DD HH24:MI:SS') 执行开始时间, to_Char(a.执行结束时间,'YYYY-MM-DD HH24:MI:SS') 执行结束时间, " & _
             "    a.执行人员, c.编号 " & vbNewLine & _
             "From Zlrptrunhistory A, Zlreports C " & vbNewLine & _
             "Where a.报表id = C.id And A.执行开始时间 Between [2] And [3] And A.报表id =[1] " & vbNewLine & _
             "Order By A.执行开始时间 Desc"
    Set rsData = OpenSQLRecord(strSQL, Me.Caption, mlngReportRunKey, dtp(0).Value, dtp(1).Value)
    vsfLog.Rows = 1
    If rsData.BOF = False Then
        With vsfLog
            For i = 1 To rsData.RecordCount
                If .Rows < i + 1 Then .Rows = .Rows + 1
                .TextMatrix(i, .ColIndex("执行人员")) = Nvl(rsData("执行人员").Value)
                .TextMatrix(i, .ColIndex("执行开始时间")) = Nvl(rsData("执行开始时间").Value)
                .TextMatrix(i, .ColIndex("执行结束时间")) = Nvl(rsData("执行结束时间").Value)
                .TextMatrix(i, .ColIndex("执行耗时")) = DateDiff("S", CDate(rsData("执行开始时间").Value), CDate(rsData("执行结束时间").Value)) & "秒"
                rsData.MoveNext
            Next
        End With
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.id
    Case 1005
        Call txtDoEvents.SetFocus
        Call RefreshData
    Case 1006
        Unload Me
    End Select
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.id
    Case 1
        Item.handle = picPane(1).hwnd
    End Select
End Sub

Private Sub Form_Activate()
    If mblnStartUp = False Then Exit Sub
    mblnStartUp = False
    Call RefreshData
    
End Sub

Private Sub Form_Load()
    mblnStartUp = True
    Call InitData
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    picPane(1).Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    Select Case Index
    Case 1
        vsfLog.Move 15, 15, picPane(Index).ScaleWidth - 30, picPane(Index).ScaleHeight - 30
    End Select
End Sub
