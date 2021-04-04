VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMspPollLog 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7785
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   10875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   4
      Left            =   8490
      ScaleHeight     =   255
      ScaleWidth      =   1275
      TabIndex        =   6
      Top             =   465
      Width           =   1305
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   1
         Left            =   -30
         TabIndex        =   7
         Top             =   -30
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   120389635
         CurrentDate     =   41401
      End
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   3
      Left            =   6480
      ScaleHeight     =   240
      ScaleWidth      =   1245
      TabIndex        =   4
      Top             =   390
      Width           =   1275
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   0
         Left            =   -30
         TabIndex        =   5
         Top             =   -30
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   120389635
         CurrentDate     =   41401
      End
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   2
      Left            =   5100
      ScaleHeight     =   240
      ScaleWidth      =   1155
      TabIndex        =   2
      Top             =   390
      Width           =   1185
      Begin VB.ComboBox cboPeiord 
         Height          =   300
         Left            =   -30
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   -30
         Width           =   1215
      End
   End
   Begin VB.PictureBox picBack 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   4140
      Index           =   1
      Left            =   105
      ScaleHeight     =   4140
      ScaleWidth      =   4740
      TabIndex        =   0
      Top             =   480
      Width           =   4740
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   3240
         Index           =   0
         Left            =   180
         TabIndex        =   1
         Top             =   90
         Width           =   3030
         _cx             =   5345
         _cy             =   5715
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
         GridLines       =   8
         GridLinesFixed  =   8
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   -15
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmMspRunLog.frx":0000
      Left            =   375
      Top             =   15
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmMspPollLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



'######################################################################################################################
'变量定义
Private Enum Command
    初始控件
    读注册表
    新增事件
    修改事件
    删除日志
    刷新数据
    详细信息
End Enum

Private mblnReading As Boolean
Private mclsMsgLog As clsMsgLog
Private mlngModualCode As Long
Private mstrSQL As String
Private mclsVsf(0) As zlVSFlexGrid.clsVsf
Private mblnStartUp As Boolean
Private mblnDataChanged As Boolean
Private mblnStartService As Boolean

Public Event AfterClose(ByVal lngModual As Long)
Public Event AfterLoad(ByVal intIndex As Integer, ByVal strContent As String)

'######################################################################################################################
'接口方法
Public Function ShowForm()
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mblnStartUp = True
    Call Form_Activate
End Function

'######################################################################################################################
'私有方法

Private Property Let DataChanged(ByVal blnData As Boolean)
    mblnDataChanged = blnData
End Property

Private Property Get DataChanged() As Boolean
    DataChanged = mblnDataChanged
End Property

Private Function ExecuteCommand(ByVal enmCommand As Command, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim rs As zlDataSQLite.SQLiteRecordset
    Dim rsSQL As New ADODB.Recordset
    Dim strTmp As String
    Dim strSQL As String
    Dim intRow As Integer
    Dim varTmp As Variant
    Dim rsCondition As ADODB.Recordset
    Dim strEditMode As String
    
    On Error GoTo errHand
            
    mblnReading = True
    
    Select Case enmCommand
    '------------------------------------------------------------------------------------------------------------------
    Case Command.初始控件
        Set mclsMsgLog = New clsMsgLog

        Call InitGrid
        Call InitCommandBar
        Call InitDockPannel

        With cboPeiord
            .Clear
            .AddItem "今  天"
            .AddItem "昨  天"
            .AddItem "前三天"
            .AddItem "本  周"
            .AddItem "前一周"
            .AddItem "前半月"
            .AddItem "本  月"
            .AddItem "前一月"
            .AddItem "前二月"
            .AddItem "本  季"
            .AddItem "前三月"
            .AddItem "本半年"
            .AddItem "前半年"
            .AddItem "自定义"
        End With
        If cboPeiord.ListCount > 0 And cboPeiord.ListIndex = -1 Then cboPeiord.ListIndex = 0

        dtp(0).Value = Format(GetBasePeriod(cboPeiord.Text, 1), dtp(0).CustomFormat)
        dtp(1).Value = Format(GetBasePeriod(cboPeiord.Text, 2), dtp(1).CustomFormat)
        
    '------------------------------------------------------------------------------------------------------------------
    Case Command.删除日志
        
        With vsf(0)
            If MsgBox("您确认要删除“" & varParam(0) & "”的日志吗？", vbQuestion + vbYesNo + vbDefaultButton2, "") = vbYes Then
                
                If mclsMsgLog.OpenLogFile(App.Path & "\Data\zlMsgSender.db") = True Then
                    
                    If varParam(0) = "全部" Then
                        strEditMode = "DeleteAll"
                    Else
                        strEditMode = "Delete"
                        Set rsCondition = zlCommFun.CreateCondition
                        Call zlCommFun.SetCondition(rsCondition, "开始时间", Format(GetBasePeriod(varParam(0), 1), "yyyy-MM-dd") & " 00:00:00")
                        Call zlCommFun.SetCondition(rsCondition, "结束时间", Format(GetBasePeriod(varParam(0), 2), "yyyy-MM-dd") & " 23:59:59")
                    End If
                    
                    If mclsMsgLog.EditRunLog(strEditMode, rsCondition) Then
                        mclsMsgLog.CloseLogFile
                        Call ExecuteCommand(Command.刷新数据)
                    End If

                    
                End If
                
            End If
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case Command.刷新数据
        
        mclsVsf(0).ClearGrid
                        
        If mclsMsgLog.OpenLogFile(App.Path & "\Data\zlMsgSender.db") = True Then
            
            Set rsCondition = zlCommFun.CreateCondition
            Call zlCommFun.SetCondition(rsCondition, "开始时间", Format(dtp(0).Value, "yyyy-MM-dd") & " 00:00:00")
            Call zlCommFun.SetCondition(rsCondition, "结束时间", Format(dtp(1).Value, "yyyy-MM-dd") & " 23:59:59")
            
            rs = mclsMsgLog.GetRunLog("Filter", rsCondition)
            If rs.DataSet.BOF = False Then Call mclsVsf(0).LoadDataSource(rs.DataSet.DataSource)
            
            DataChanged = False
            
            mclsMsgLog.CloseLogFile
        End If
    End Select
    
    
    GoTo EndHand

    '出错处理
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
    '------------------------------------------------------------------------------------------------------------------
EndHand:
    mblnReading = False
End Function

Private Function InitGrid() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    '初始网格控件
    Set mclsVsf(0) = New zlVSFlexGrid.clsVsf
    With mclsVsf(0)
        Call .Initialize(Me.Controls, vsf(0), True, False)
        Call .ClearColumn
        
        Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[序号]", False)
            
        Call .AppendColumn("时间", 1890, flexAlignLeftCenter, flexDTString, , "Log_Time", True, False)
        Call .AppendColumn("类型", 600, flexAlignCenterCenter, flexDTString, , "Log_Type", True, False)
        Call .AppendColumn("信息", 3000, flexAlignLeftCenter, flexDTString, , "Log_Desc", True, False)
        
        .IndicatorMode = 2
        .IndicatorCol = .ColIndex("序号")
        .ConstCol = .ColIndex("序号")
                
    End With
            
    InitGrid = True
    
End Function

Private Function InitCommandBar() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    Dim objFindKey As CommandBarControl
    
    On Error GoTo errHand
    
    '------------------------------------------------------------------------------------------------------------------
    '初始设置
    Call zlCommFun.CommandBarInit(cbsMain)
    cbsMain.VisualTheme = xtpThemeNativeWinXP
    Set cbsMain.Icons = zlCommFun.GetPubIcons
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
    
    Set objFindKey = zlCommFun.NewToolBar(objBar, xtpControlPopup, 2, "操作", , , xtpButtonIconAndCaption)
    objFindKey.IconId = conMenu_Edit_NewItem
        
'    Set objControl = objFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_NewItem, "运行日志(&R)")
'    objControl.IconId = 10
'    Set objControl = objFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Modify, "发送日志(&S)")
'    objControl.IconId = 10
            
    
    Set objControl = objFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Delete, "&1.清除一月前的日志"): objControl.Parameter = "一月前": objControl.IconId = 1
    Set objControl = objFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Delete, "&2.清除一周前的日志"): objControl.Parameter = "一周前": objControl.IconId = 1
    Set objControl = objFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Delete, "&3.清除一天前的日志"): objControl.Parameter = "一天前": objControl.IconId = 1
    Set objControl = objFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Delete, "&9.清除全部日志"): objControl.Parameter = "全部": objControl.IconId = 1

            
    Set objControl = objFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)")
    objControl.BeginGroup = True
    
'    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_NewItem, "运行日志", True)
'    objControl.IconId = 10
'    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Modify, "发送日志")
'    objControl.IconId = 10
    
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlLabel, 1, "时间范围：", , , xtpButtonCaption)
    objControl.BeginGroup = True
    Set cbrCustom = zlCommFun.NewToolBar(objBar, xtpControlCustom, conMenu_View_Location, "")
    cbrCustom.Handle = picBack(2).hwnd
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlLabel, 1, "从", , , xtpButtonCaption)
    Set cbrCustom = zlCommFun.NewToolBar(objBar, xtpControlCustom, conMenu_View_Location, "")
    cbrCustom.Handle = picBack(3).hwnd
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlLabel, 1, "到", , , xtpButtonCaption)
    Set cbrCustom = zlCommFun.NewToolBar(objBar, xtpControlCustom, conMenu_View_Location, "")
    cbrCustom.Handle = picBack(4).hwnd
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Filter, "过滤")
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Refresh, "刷新", True)
    
'    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_File_Close, "关闭")
'    objControl.Flags = xtpFlagRightAlign
    
    '------------------------------------------------------------------------------------------------------------------
    '命令的快键绑定:公共部份主界面已处理

    With cbsMain.KeyBindings
        .Add 0, vbKeyF5, conMenu_View_Refresh           '刷新
    End With
        
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Function

Private Sub InitDockPannel()
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim objPane As Pane

    Set objPane = dkpMain.CreatePane(1, 100, 300, DockTopOf, Nothing)
    objPane.Title = "日志"
    objPane.Options = PaneNoCaption
        
    dkpMain.SetCommandBars cbsMain
    Call zlCommFun.DockPannelInit(dkpMain)

End Sub

Public Function ShowConetneMenu(Optional ByVal bytPlace As Byte = 1) As CommandBar
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim cbrPopupBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim cbrPopupItem As CommandBarControl
    Dim cbrPopupItem2 As CommandBarControl
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrControl2 As CommandBarControl
    
    '弹出菜单处理
    
    On Error GoTo errHand
    
    Set cbrPopupBar = cbsMain.Add("弹出菜单", xtpBarPopup)
    
    Select Case bytPlace
    Case 1  '
        
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Delete, "&1.清除一月前的日志"): cbrPopupItem.Parameter = "一月前": cbrPopupItem.IconId = 1
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Delete, "&2.清除一周前的日志"): cbrPopupItem.Parameter = "一周前": cbrPopupItem.IconId = 1
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Delete, "&3.清除一天前的日志"): cbrPopupItem.Parameter = "一天前": cbrPopupItem.IconId = 1
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Delete, "&9.清除全部日志"): cbrPopupItem.Parameter = "全部": cbrPopupItem.IconId = 1
        
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)")
        cbrPopupItem.BeginGroup = True
    
    End Select
    
    Set ShowConetneMenu = cbrPopupBar
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Sub cboPeiord_Click()
    If mblnReading Then Exit Sub
    
    If cboPeiord.Text <> "自定义" Then
        dtp(0).Value = Format(GetBasePeriod(cboPeiord.Text, 1), dtp(0).CustomFormat)
        dtp(1).Value = Format(GetBasePeriod(cboPeiord.Text, 2), dtp(1).CustomFormat)
        
        Call ExecuteCommand(Command.刷新数据)
    Else
        DataChanged = True
    End If
    
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)

    Select Case Control.id
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Refresh, conMenu_View_Filter
        Call ExecuteCommand(Command.刷新数据)
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete
        
        Call ExecuteCommand(Command.删除日志, Control.Parameter)
        
    Case conMenu_File_Close
    '--------------------------------------------------------------------------------------------------------------
        Unload Me
        RaiseEvent AfterClose(mlngModualCode)
    End Select
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
        
    Select Case Control.id
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Filter
        
        Control.Enabled = DataChanged
    End Select
    
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.id
    Case 1
        Item.Handle = picBack(1).hwnd
    End Select
End Sub

Private Sub dtp_Change(Index As Integer)
    '更改时间段名称为“自定义“
    mblnReading = True
    
    Select Case Index
    Case 0, 1
        Call zlControl.CboLocate(cboPeiord, "自定义")
    End Select
    
    mblnReading = False
    
    DataChanged = True
End Sub

Private Sub Form_Activate()
    If mblnStartUp = False Then Exit Sub
    DoEvents
    mblnStartUp = False
    
    Call ExecuteCommand(Command.刷新数据)
End Sub

Private Sub Form_Load()
    mblnStartUp = True
    mlngModualCode = 1001
    
    Call ExecuteCommand(Command.初始控件)
    Call ExecuteCommand(Command.读注册表)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mclsMsgLog = Nothing
End Sub

Private Sub picBack_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 1
        vsf(0).Move 15, 15, picBack(Index).Width - 30, picBack(Index).Height - 30
    End Select
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsf(Index).AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsf_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbrPopupBar As CommandBar
    
    Select Case Button
    '------------------------------------------------------------------------------------------------------------------
    Case 2          '弹出菜单处理
        Call zlCommFun.SendLMouseButton(vsf(Index).hwnd, X, Y)
        Select Case Index
        Case 0
            If mclsVsf(Index).MoveColumn = False Then
                Call ShowConetneMenu(1).ShowPopup
            End If
        End Select
        
    End Select
End Sub
