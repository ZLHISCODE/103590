VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Begin VB.Form frmMipClientRunlog 
   Caption         =   "运行日志查阅"
   ClientHeight    =   8775
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13425
   Icon            =   "frmMipClientRunlog.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   13425
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picBack 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   4140
      Index           =   1
      Left            =   105
      ScaleHeight     =   4140
      ScaleWidth      =   4740
      TabIndex        =   6
      Top             =   495
      Width           =   4740
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   3240
         Index           =   0
         Left            =   180
         TabIndex        =   7
         Top             =   60
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
         GridColor       =   12632256
         GridColorFixed  =   12632256
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
         GridLinesFixed  =   1
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
         AutoSizeMode    =   1
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
         WordWrap        =   -1  'True
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
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   2
      Left            =   5100
      ScaleHeight     =   240
      ScaleWidth      =   1155
      TabIndex        =   4
      Top             =   405
      Width           =   1185
      Begin VB.ComboBox cboPeiord 
         Height          =   300
         Left            =   -30
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   -30
         Width           =   1215
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
      TabIndex        =   2
      Top             =   405
      Width           =   1275
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   0
         Left            =   -30
         TabIndex        =   3
         Top             =   -30
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   130744323
         CurrentDate     =   41401
      End
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   4
      Left            =   8490
      ScaleHeight     =   255
      ScaleWidth      =   1275
      TabIndex        =   0
      Top             =   480
      Width           =   1305
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   1
         Left            =   -30
         TabIndex        =   1
         Top             =   -30
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   130744323
         CurrentDate     =   41401
      End
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmMipClientRunlog.frx":6852
      Left            =   375
      Top             =   30
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmMipClientRunlog"
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
Private mclsMipRunLog As clsMipRunLog
Private mlngModualCode As Long
Private mstrSQL As String
Private mclsVsf(0) As zlVSFlexGrid.clsVsf
Private mblnStartUp As Boolean
Private mblnDataChanged As Boolean
Private mblnStartService As Boolean
Private mstrLogFile As String

Public Event AfterClose(ByVal lngModual As Long)
Public Event AfterLoad(ByVal intIndex As Integer, ByVal strContent As String)

'######################################################################################################################
'接口方法
Public Function ShowForm(ByVal objParentForm As Object, ByVal strLogFile As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mblnStartUp = True
    mstrLogFile = strLogFile
    
    Set mclsMipRunLog = New clsMipRunLog
    Call mclsMipRunLog.Initialize(mstrLogFile)

    Me.Show , objParentForm
        
    Call ExecuteCommand(Command.刷新数据)
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
    Dim blnMuliSelect As Boolean
    
    On Error GoTo errHand
            
    mblnReading = True
    
    Select Case enmCommand
    '------------------------------------------------------------------------------------------------------------------
    Case Command.初始控件
                
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
        
            blnMuliSelect = False
            For intRow = 1 To .Rows - 1
                If Val(Abs(.TextMatrix(intRow, .ColIndex("选择")))) = 1 Then
                    blnMuliSelect = True
                    Exit For
                End If
            Next
            
            If blnMuliSelect = True Then
                If MsgBox("您确认要删除已经勾选的日志吗？", vbQuestion + vbYesNo + vbDefaultButton2, "") = vbYes Then
                    
                    If mclsMipRunLog.OpenRunLogFile() = True Then
                        Set rsCondition = zlCommFun.CreateCondition
                        For intRow = 1 To .Rows - 1
                            If Val(Abs(.TextMatrix(intRow, .ColIndex("选择")))) = 1 Then
                                Call zlCommFun.SetCondition(rsCondition, "ID", .TextMatrix(intRow, .ColIndex("ID")))
                                Call mclsMipRunLog.EditRunLog("DeleteID", rsCondition)
                            End If
                        Next

                        Call ExecuteCommand(Command.刷新数据)
                        
                    End If
                    
                End If
            
            Else
                If MsgBox("您确认要删除当前行的日志吗？", vbQuestion + vbYesNo + vbDefaultButton2, "") = vbYes Then
                    
                    If mclsMipRunLog.OpenRunLogFile() = True Then
                        
                        Set rsCondition = zlCommFun.CreateCondition
                        Call zlCommFun.SetCondition(rsCondition, "ID", .TextMatrix(.Row, .ColIndex("ID")))
                                                                        
                        If mclsMipRunLog.EditRunLog("DeleteID", rsCondition) Then
                            mclsMipRunLog.CloseRunLogFile
                            Call ExecuteCommand(Command.刷新数据)
                        End If
    
                    End If
                    
                End If
            End If
            
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case Command.刷新数据
        
        mclsVsf(0).ClearGrid
                        
        If mclsMipRunLog.OpenRunLogFile() = True Then
            
            Set rsCondition = zlCommFun.CreateCondition
            Call zlCommFun.SetCondition(rsCondition, "开始时间", Format(dtp(0).Value, "yyyy-MM-dd") & " 00:00:00")
            Call zlCommFun.SetCondition(rsCondition, "结束时间", Format(dtp(1).Value, "yyyy-MM-dd") & " 23:59:59")
            
            rs = mclsMipRunLog.GetRunLog("Filter", rsCondition)
            If rs.DataSet.BOF = False Then Call mclsVsf(0).LoadDataSource(rs.DataSet.DataSource)
            
            vsf(0).AutoSize 1, vsf(0).Cols - 1
            
            DataChanged = False
            
            mclsMipRunLog.CloseRunLogFile
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
        
        Call .Initialize(Me.Controls, vsf(0), True, True, gfrmMipResource.ils16)
        
        Call .ClearColumn
        
        Call .AppendColumn("", 255, flexAlignCenterTop, flexDTString, "", "[序号]", False)
        Call .AppendColumn("", 300, flexAlignCenterTop, flexDTBoolean, "", "[选择]", False)
        Call .AppendColumn("", 255, flexAlignCenterTop, flexDTString, "", "[图标]", False)
        Call .AppendColumn("ID", 0, flexAlignLeftTop, flexDTString, , "ID", True, False, , True)
        Call .AppendColumn("时间", 1890, flexAlignLeftTop, flexDTString, , "Log_Time", True, False)
        Call .AppendColumn("类型", 600, flexAlignLeftTop, flexDTString, , "Log_Type", True, False)
        Call .AppendColumn("信息", 3000, flexAlignLeftTop, flexDTString, , "Log_Desc", True, False)
        
        .IndicatorMode = 2
        .IndicatorCol = .ColIndex("序号")
        .ConstCol = .ColIndex("序号")
                
        Call .InitializeEdit(True, False, False)
        Call .InitializeEditColumn(.ColIndex("选择"), True, vbVsfEditCheck)
        mclsVsf(0).AppendRows = True
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
'    cbsMain.VisualTheme = xtpThemeNativeWinXP
    Set cbsMain.Icons = frmMipResource.imgPublic.Icons
    cbsMain.Options.LargeIcons = True
    
    '------------------------------------------------------------------------------------------------------------------
    '菜单定义:包括公共部份，请对xtpControlPopup类型的命令ID重新赋值

    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap

    '------------------------------------------------------------------------------------------------------------------
    '文件
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    objMenu.id = conMenu_FilePopup
    Set objControl = zlCommFun.NewCommandBar(objMenu, xtpControlButton, conMenu_File_Exit, "退出(&X)", True, , , "退出运行日志查阅功能")
    '------------------------------------------------------------------------------------------------------------------
    '分类
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    objMenu.id = conMenu_EditPopup
    Set objControl = zlCommFun.NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_SelAll, "全选(&A)", , , , "将当前列表中的所有数据置为勾选状态")
    Set objControl = zlCommFun.NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_ClsAll, "全清(&C)", , , , "将当前列表中的所有数据置为非勾选状态")
    Set objControl = zlCommFun.NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Delete, "清除(&D)", True, , , "清除当前行或者勾选中的运行日志")
        
    '------------------------------------------------------------------------------------------------------------------
    '查看
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    objMenu.id = conMenu_ViewPopup
    Set objPopup = zlCommFun.NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)")
    Set objControl = zlCommFun.NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", , , , "显示/隐藏工具栏按钮")
    Set objControl = zlCommFun.NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", , , , "显示/隐藏工具栏按钮上的文字内容")
    Set objControl = zlCommFun.NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", , , , "设置工具栏按钮图标为大图标或小图标")
    
    Set objControl = zlCommFun.NewCommandBar(objMenu, xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)", , , , "显示/隐藏状态栏")
    
    Set objControl = zlCommFun.NewCommandBar(objMenu, xtpControlButton, conMenu_View_Refresh, "刷新(&R)", True, , , "按当前设置的条件重新刷新数据")
    
    '------------------------------------------------------------------------------------------------------------------
    '帮助
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    objMenu.id = conMenu_HelpPopup
    Set objControl = zlCommFun.NewCommandBar(objMenu, xtpControlButton, conMenu_Help_Help, "帮助主题(&H)", , , , "显示关于运行日志查阅的操作说明")
    Set objControl = zlCommFun.NewCommandBar(objMenu, xtpControlButton, conMenu_Help_About, "关于(&A)…", True, , , "显示有关运行日志的相关说明")
    
    
    '------------------------------------------------------------------------------------------------------------------
    '工具栏定义:包括公共部份

    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_SelAll, "全选", True, , , , , "将当前列表中的所有数据置为勾选状态")
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_ClsAll, "全清", , , , , , "将当前列表中的所有数据置为非勾选状态")
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Delete, "清除", True, , , , , "清除当前行或者勾选中的运行日志")
        
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlLabel, 1, "时间范围：", , , xtpButtonCaption, , , "设置查询过滤的时间范围")
    objControl.BeginGroup = True
    Set cbrCustom = zlCommFun.NewToolBar(objBar, xtpControlCustom, conMenu_View_Location, "")
    cbrCustom.Handle = picBack(2).hWnd
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlLabel, 1, "从", , , xtpButtonCaption)
    Set cbrCustom = zlCommFun.NewToolBar(objBar, xtpControlCustom, conMenu_View_Location, "", , , , , , "设置查询过滤的开始时间")
    cbrCustom.Handle = picBack(3).hWnd
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlLabel, 1, "到", , , xtpButtonCaption)
    Set cbrCustom = zlCommFun.NewToolBar(objBar, xtpControlCustom, conMenu_View_Location, "", , , , , , "设置查询过滤的结束时间")
    cbrCustom.Handle = picBack(4).hWnd
        
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Refresh, "刷新", True, , , , , "按当前设置的条件重新刷新数据")
        
    cbsMain.StatusBar.Visible = True
    cbsMain.StatusBar.IdleText = "准备"
    Call cbsMain.StatusBar.AddPane(0)
    Call cbsMain.StatusBar.SetPaneText(0, cbsMain.StatusBar.IdleText)
    Call cbsMain.StatusBar.SetPaneStyle(0, SBPS_STRETCH)
    Call cbsMain.StatusBar.AddPane(ID_INDICATOR_CAPS)
    Call cbsMain.StatusBar.AddPane(ID_INDICATOR_NUM)
    Call cbsMain.StatusBar.AddPane(ID_INDICATOR_SCRL)

    '------------------------------------------------------------------------------------------------------------------
    '命令的快键绑定:公共部份主界面已处理

    With cbsMain.KeyBindings
        .Add 0, vbKeyF5, conMenu_View_Refresh               '刷新
        .Add 0, vbKeyDelete, conMenu_Edit_Delete            '清除
        .Add FCONTROL, vbKeyA, conMenu_Edit_SelAll          '全选
        .Add FSHIFT, vbKeyDelete, conMenu_Edit_ClsAll       '全清
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
            
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_SelAll, "全选(&A)")
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_ClsAll, "全清(&C)")
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Delete, "清除(&D)")
        cbrPopupItem.BeginGroup = True
        
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
    Dim lngLoop As Long
    Dim objControl As Object
    
    Select Case Control.id
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_SelAll
        With vsf(0)
            .Cell(flexcpText, 1, .ColIndex("选择"), .Rows - 1, .ColIndex("选择")) = 1
        End With
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_ClsAll
        
        With vsf(0)
            .Cell(flexcpText, 1, .ColIndex("选择"), .Rows - 1, .ColIndex("选择")) = 0
        End With
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Refresh
        
        Call ExecuteCommand(Command.刷新数据)
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete
        
        Call ExecuteCommand(Command.删除日志, Control.Parameter)
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_ToolBar_Button     '工具栏
    
        For lngLoop = 2 To cbsMain.Count
            cbsMain(lngLoop).Visible = Not cbsMain(lngLoop).Visible
        Next
        cbsMain.RecalcLayout
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_ToolBar_Text      '按钮文字
    
        For lngLoop = 2 To cbsMain.Count
            For Each objControl In cbsMain(lngLoop).Controls
                If objControl.Type = xtpControlButton Then
                    objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                End If
            Next
        Next
        cbsMain.RecalcLayout
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_ToolBar_Size      '大图标
    
        cbsMain.Options.LargeIcons = Not cbsMain.Options.LargeIcons
        cbsMain.RecalcLayout
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_StatusBar         '状态栏
    
        cbsMain.StatusBar.Visible = Not cbsMain.StatusBar.Visible
        cbsMain.RecalcLayout
        
    Case conMenu_File_Close
    '--------------------------------------------------------------------------------------------------------------
        Unload Me
        RaiseEvent AfterClose(mlngModualCode)
    End Select
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
'    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
        
    Select Case Control.id
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Filter
        
        Control.Enabled = DataChanged
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_ToolBar_Button            '工具栏
        If cbsMain.Count >= 2 Then
            Control.Checked = cbsMain(2).Visible
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_ToolBar_Text              '图标文字
        If cbsMain.Count >= 2 Then
            Control.Checked = Not (cbsMain(2).Controls(1).Style = xtpButtonIcon)
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_ToolBar_Size              '大图标
        Control.Checked = cbsMain.Options.LargeIcons
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_StatusBar                 '状态栏
        Control.Checked = cbsMain.StatusBar.Visible
                
    End Select
    
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.id
    Case 1
        Item.Handle = picBack(1).hWnd
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

Private Sub Form_Load()
    mblnStartUp = True
    mlngModualCode = 1001
    
    Call ExecuteCommand(Command.初始控件)
    Call ExecuteCommand(Command.读注册表)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If Not (mclsMipRunLog Is Nothing) Then
        Set mclsMipRunLog = Nothing
    End If
    
    If Not (mclsVsf(0) Is Nothing) Then
        Set mclsVsf(0) = Nothing
    End If
    
End Sub

Private Sub picBack_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 1
        vsf(0).Move 15, 15, picBack(Index).Width - 30, picBack(Index).Height - 30
        mclsVsf(0).AppendRows = True
    End Select
End Sub

Private Sub vsf_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Call mclsVsf(Index).AfterEdit(Row, Col)
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsf(Index).AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsf_AfterScroll(Index As Integer, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mclsVsf(0).AppendRows = True
End Sub

Private Sub vsf_AfterSort(Index As Integer, ByVal Col As Long, Order As Integer)
    mclsVsf(0).AppendRows = True
End Sub

Private Sub vsf_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    mclsVsf(0).AppendRows = True
End Sub

Private Sub vsf_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbrPopupBar As CommandBar
    
    Select Case Button
    '------------------------------------------------------------------------------------------------------------------
    Case 2          '弹出菜单处理
        Call zlCommFun.SendLMouseButton(vsf(Index).hWnd, X, Y)
        Select Case Index
        Case 0
            If mclsVsf(Index).MoveColumn = False Then
                Call ShowConetneMenu(1).ShowPopup
            End If
        End Select
        
    End Select
End Sub

Private Sub vsf_SetupEditWindow(Index As Integer, ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    Call mclsVsf(Index).EditSelAll
End Sub

Private Sub vsf_StartEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf(Index).BeforeEdit(Row, Col, Cancel)
End Sub

Private Sub vsf_ValidateEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf(Index).ValidateEdit(Col, Cancel)
End Sub
