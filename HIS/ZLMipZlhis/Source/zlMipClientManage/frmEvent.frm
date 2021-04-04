VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmEvent 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12465
   LinkTopic       =   "Form1"
   ScaleHeight     =   7695
   ScaleWidth      =   12465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   1800
      Index           =   1
      Left            =   1620
      ScaleHeight     =   1800
      ScaleWidth      =   2865
      TabIndex        =   0
      Top             =   1365
      Width           =   2865
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1215
         Index           =   0
         Left            =   270
         TabIndex        =   1
         Top             =   240
         Width           =   1860
         _cx             =   3281
         _cy             =   2143
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   -15
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmEvent.frx":0000
      Left            =   360
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmEvent"
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
    删除事件
    刷新数据
    刷新指定事件
    移除指定事件
    刷新事件消息
End Enum

Private mlngModualCode As Long
Private mstrSQL As String
Private mclsVsf(0) As zlVSFlexGrid.clsVsf
Private mblnStartUp As Boolean
Private mobjFindKey As CommandBarControl
Private mstrFindKey As String
Private mblnDataChanged As Boolean


Public Event AfterClose(ByVal lngModual As Long)
Public Event AfterLoad(ByVal intIndex As Integer, ByVal strContent As String)
Private WithEvents mfrmEventEdit As frmEventEdit
Attribute mfrmEventEdit.VB_VarHelpID = -1
Private mfrmEventMsg As frmEventMsg

'######################################################################################################################
'接口方法
Public Function ShowForm()
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Call Form_Activate
End Function

'######################################################################################################################
'私有方法
Private Function ExecuteCommand(ByVal enmCommand As Command, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strTmp As String
    Dim strSQL As String
    Dim intRow As Integer
    Dim varTmp As Variant

    On Error GoTo errHand
            
    Select Case enmCommand
    '------------------------------------------------------------------------------------------------------------------
    Case Command.初始控件
        
        Call InitGrid
        Call InitCommandBar
        Call InitDockPannel
    '------------------------------------------------------------------------------------------------------------------
    Case Command.新增事件
    
        If mfrmEventEdit Is Nothing Then
            Set mfrmEventEdit = New frmEventEdit
            Call mfrmEventEdit.InitDialog(Me, mlngModualCode)
        End If
    
        Call mfrmEventEdit.NewData
        DoEvents
        Me.SetFocus
    '------------------------------------------------------------------------------------------------------------------
    Case Command.修改事件
    
        If mfrmEventEdit Is Nothing Then
            Set mfrmEventEdit = New frmEventEdit
            Call mfrmEventEdit.InitDialog(Me, mlngModualCode)
        End If
        
        With vsf(0)
            Call mfrmEventEdit.ModifyData(.TextMatrix(.Row, .ColIndex("id")))
        End With
        DoEvents
        Me.SetFocus
    '------------------------------------------------------------------------------------------------------------------
    Case Command.删除事件
        
        If mfrmEventEdit Is Nothing Then
            Set mfrmEventEdit = New frmEventEdit
            Call mfrmEventEdit.InitDialog(Me, mlngModualCode)
        End If
        
        With vsf(0)
            If MsgBox("您确认要删除当前业务事件吗？", vbQuestion + vbYesNo + vbDefaultButton2, ParamInfo.系统名称) = vbYes Then
                Call mfrmEventEdit.DeleteData(.TextMatrix(.Row, .ColIndex("id")))
            End If
        End With
        
        DoEvents
        Me.SetFocus
        
    '------------------------------------------------------------------------------------------------------------------
    Case Command.刷新数据
        
        With vsf(0)
            mclsVsf(0).SaveKey = Trim(.TextMatrix(.Row, .ColIndex("id")))
                    
            ExecuteCommand = mclsVsf(0).LoadDataSource(gclsBusiness.EventRead)
            
            Call mclsVsf(0).RestoreRow(mclsVsf(0).SaveKey, .ColIndex("id"))
        End With
        
        Call ExecuteCommand(Command.刷新事件消息)
        
    '------------------------------------------------------------------------------------------------------------------
    Case Command.刷新指定事件
        
        ExecuteCommand = LoadCustomData(Trim(varParam(0)))
        Call ExecuteCommand(Command.刷新事件消息)
        
    '------------------------------------------------------------------------------------------------------------------
    Case Command.移除指定事件
        
        With vsf(0)
            
            intRow = mclsVsf(0).FindRow(Trim(varParam(0)), .ColIndex("id"))
            
            If intRow > 0 Then
                If .Rows > 2 Then
                    .RemoveItem .Row
                    mclsVsf(0).AppendRows = True
                Else
                    Call mclsVsf(0).ClearGrid
                End If
            End If
        End With
    
    '------------------------------------------------------------------------------------------------------------------
    Case Command.刷新事件消息
        
        With vsf(0)
            Call mfrmEventMsg.ReadData(.TextMatrix(.Row, .ColIndex("id")))
        End With
        
    End Select
    
    
    GoTo EndHand

    '出错处理
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
    Call zlComLib.SaveErrLog
    
    '------------------------------------------------------------------------------------------------------------------
EndHand:
End Function

Private Function LoadCustomData(ByVal strDataKey As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intRow As Integer
    Dim rsData As ADODB.Recordset
    Dim rsCondition As ADODB.Recordset
    
    Set rsCondition = zlCommFun.CreateCondition
    Call zlCommFun.SetCondition(rsCondition, "id", strDataKey)
    
    Set rsData = gclsBusiness.EventRead("id", rsCondition)
    If rsData.BOF = True Then Exit Function
    
    With vsf(0)
        
        intRow = mclsVsf(0).FindRow(strDataKey, .ColIndex("id"))
        
        If intRow > 0 Then
            '已加载
            .Row = intRow
        Else
            '未加载
            If Trim(.TextMatrix(.Rows - 1, .ColIndex("id"))) <> "" Then .Rows = .Rows + 1
            .Row = .Rows - 1
        End If
        
        Call mclsVsf(0).LoadGridRow(.Row, rsData)
    End With
    
    mclsVsf(0).AppendRows = True
    
    LoadCustomData = True
    
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
        Call .Initialize(Me.Controls, vsf(0), True, False, gclsBusiness.GetImageList(16))
        Call .ClearColumn
        
        Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[图标]", False)
        Call .AppendColumn("id", 0, flexAlignLeftCenter, flexDTString, , "id", True)
        Call .AppendColumn("名称", 1800, flexAlignLeftCenter, flexDTString, , "title", True)
        Call .AppendColumn("编码", 600, flexAlignLeftCenter, flexDTString, , "code", True)
        Call .AppendColumn("类型", 1200, flexAlignLeftCenter, flexDTString, , "kind", True)
        Call .AppendColumn("应用", 750, flexAlignLeftCenter, flexDTString, , "app", True)
        Call .AppendColumn("设备", 750, flexAlignLeftCenter, flexDTString, , "device", True)
        Call .AppendColumn("注释", 1800, flexAlignLeftCenter, flexDTString, , "note", True)
        
        .AppendRows = True
        
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
        
    Set objControl = objFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_NewItem, "新增事件(&N)")
    Set objControl = objFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Modify, "修改事件(&M)")
    Set objControl = objFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Delete, "删除事件(&D)")
    
    Set objControl = objFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_NewParent, "新增消息(&A)")
    objControl.BeginGroup = True
    objControl.IconId = conMenu_Edit_NewItem
    
    Set objControl = objFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_ModifyParent, "修改消息(&U)")
    objControl.IconId = conMenu_Edit_Modify
    
    Set objControl = objFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_DeleteParent, "删除消息(&D)")
    objControl.IconId = conMenu_Edit_Delete
    
    Set objControl = objFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)")
    objControl.BeginGroup = True
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_NewItem, "新增", True)
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Modify, "修改")
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Delete, "删除")
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Refresh, "刷新", True)
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_File_Close, "关闭")
    objControl.Flags = xtpFlagRightAlign
    
    '------------------------------------------------------------------------------------------------------------------
    '命令的快键绑定:公共部份主界面已处理

    With cbsMain.KeyBindings
        .Add 0, vbKeyF5, conMenu_View_Refresh           '刷新
        .Add 0, vbKeyDelete, conMenu_Edit_Delete

        .Add FCONTROL, vbKeyN, conMenu_Edit_NewItem     '新增
    End With
        
    Exit Function
    
errHand:
    If zlComLib.ErrCenter = 1 Then
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

    Set objPane = dkpMain.CreatePane(1, 100, 100, DockLeftOf, Nothing)
    objPane.Title = "事件"
    objPane.Options = PaneNoCaption
    
    Set objPane = dkpMain.CreatePane(2, 300, 100, DockRightOf, objPane)
    objPane.Title = "详细"
    objPane.Options = PaneNoCaption
            
    dkpMain.SetCommandBars cbsMain
    Call zlCommFun.DockPannelInit(dkpMain)

End Sub

'######################################################################################################################
'对象事件
Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)

    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_NewItem               '增加
        
        Call ExecuteCommand(Command.新增事件)
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Modify               '修改
        
        Call ExecuteCommand(Command.修改事件)
                
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete

        Call ExecuteCommand(Command.删除事件)
                
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_NewParent, conMenu_Edit_ModifyParent, conMenu_Edit_DeleteParent
        
        Call mfrmEventMsg.Execute(Control)
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Refresh               '刷新

        Call ExecuteCommand(Command.刷新数据)
        
    Case conMenu_File_Close
    '--------------------------------------------------------------------------------------------------------------
        Unload Me
        RaiseEvent AfterClose(mlngModualCode)
    End Select
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    With vsf(0)
        Select Case Control.ID
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_NewItem               '增加
                    
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Modify               '修改
            
            Control.Enabled = (Trim(.TextMatrix(.Row, .ColIndex("id"))) <> "" And Control.Visible)
                    
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Delete
    
            Control.Enabled = (Trim(.TextMatrix(.Row, .ColIndex("id"))) <> "" And Control.Visible)
                                
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_NewParent, conMenu_Edit_ModifyParent, conMenu_Edit_DeleteParent
            
            Call mfrmEventMsg.Update(Control)
            
        End Select
    End With
End Sub

Private Sub dkpMain_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    Select Case Pane
    Case dkpMain.Panes(1)
        Select Case Action
        Case PaneActionPinned, PaneActionPinning, PaneActionExpanded, PaneActionExpanding, PaneActionCollapsed, PaneActionCollapsing
            Cancel = False
        Case Else
            Cancel = True
        End Select
    End Select
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = picPane(1).hWnd
    Case 2
        Set mfrmEventMsg = New frmEventMsg
        Call mfrmEventMsg.InitForm(Me, mlngModualCode)
        Item.Handle = mfrmEventMsg.hWnd
    End Select
End Sub

Private Sub Form_Activate()
    If mblnStartUp = False Then Exit Sub
    DoEvents
    mblnStartUp = False
    
    Call ExecuteCommand(Command.刷新数据)
End Sub

Private Sub Form_Load()
    mblnStartUp = True
    mlngModualCode = 1002
    
    Call ExecuteCommand(Command.初始控件)
    Call ExecuteCommand(Command.读注册表)

    Call zlComLib.RestoreWinState(Me, App.ProductName)
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, gblnShowInTaskBar)
'    Call zlDataBase.ShowReportMenu(Me, ParamInfo.系统号, ParamInfo.模块号, UserInfo.模块权限)
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Call zlCommFun.SetPaneRange(dkpMain, 1, 100, 15, 400, Me.ScaleHeight)
    
    dkpMain.RecalcLayout
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mclsVsf(0) = Nothing
    Set mobjFindKey = Nothing
    If Not (mfrmEventEdit Is Nothing) Then
        Unload mfrmEventEdit
        Set mfrmEventEdit = Nothing
    End If
    
    If Not (mfrmEventMsg Is Nothing) Then
        Unload mfrmEventMsg
        Set mfrmEventMsg = Nothing
    End If
            
End Sub

Private Sub mfrmEventEdit_AfterDeleteData(ByVal DataKey As String)
    
    Call ExecuteCommand(Command.移除指定事件, DataKey)
    
End Sub

Private Sub mfrmEventEdit_AfterModifyData(ByVal DataKey As String)
    
    Call ExecuteCommand(Command.刷新指定事件, DataKey)
    
End Sub

Private Sub mfrmEventEdit_AfterNewData(ByVal DataKey As String)
    
    Call ExecuteCommand(Command.刷新指定事件, DataKey)
                
End Sub

Private Sub mfrmEventEdit_Backward(DataKey As String, Cancel As Boolean)
    Dim intRow As Integer
    
    With vsf(0)
    
        intRow = mclsVsf(0).FindRow(DataKey, .ColIndex("id"))
        If intRow > 0 And .Row <> intRow Then .Row = intRow
        
        If .Row < .Rows - 1 Then
            .Row = .Row + 1
            .ShowCell .Row, .Col
            DataKey = .TextMatrix(.Row, .ColIndex("id"))
        End If
    End With
            
End Sub

Private Sub mfrmEventEdit_Forward(DataKey As String, Cancel As Boolean)
    
    Dim intRow As Integer
    
    With vsf(0)
        
        intRow = mclsVsf(0).FindRow(DataKey, .ColIndex("id"))
        If intRow > 0 And .Row <> intRow Then .Row = intRow
                
        If .Row > 1 Then
            .Row = .Row - 1
            .ShowCell .Row, .Col
            DataKey = .TextMatrix(.Row, .ColIndex("id"))
        End If
    End With
    
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 1
        vsf(0).Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
        mclsVsf(0).AppendRows = True
    End Select
End Sub

Private Sub vsf_AfterMoveColumn(Index As Integer, ByVal Col As Long, Position As Long)
    mclsVsf(Index).AppendRows = True
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsf(Index).AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
    
    If OldRow <> NewRow Then
        Call ExecuteCommand(Command.刷新事件消息)
    End If
End Sub

Private Sub vsf_AfterScroll(Index As Integer, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mclsVsf(Index).AppendRows = True
End Sub

Private Sub vsf_AfterSort(Index As Integer, ByVal Col As Long, Order As Integer)
    With vsf(Index)
        Call mclsVsf(Index).RestoreRow(mclsVsf(Index).SaveKey, .ColIndex("id"))
        .ShowCell .Row, .Col
    End With
End Sub

Private Sub vsf_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    mclsVsf(Index).AppendRows = True
End Sub

Private Sub vsf_BeforeSort(Index As Integer, ByVal Col As Long, Order As Integer)
    With vsf(Index)
        mclsVsf(Index).SaveKey = Trim(.TextMatrix(.Row, .ColIndex("id")))
    End With
End Sub

Private Sub vsf_BeforeUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
   Call mclsVsf(Index).BeforeResizeColumn(Col, Cancel)
End Sub

Private Sub vsf_DblClick(Index As Integer)
    Dim objMenu As CommandBarControl
    
    Set objMenu = cbsMain.FindControl(, conMenu_Edit_Modify, False)
    If Not (objMenu Is Nothing) Then
        If objMenu.Enabled = True Then
            Call cbsMain_Execute(objMenu)
        End If
    End If
End Sub

Private Sub vsf_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call vsf_DblClick(Index)
    End If
End Sub

Private Sub vsf_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    mclsVsf(Index).MoveColumn = (vsf(Index).MouseRow = 0)
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

Public Function ShowConetneMenu(Optional ByVal bytPlace As Byte = 1) As CommandBar
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim cbrPopupBar As CommandBar
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
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_NewItem, "新增事件(&N)")
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Modify, "修改事件(&M)")
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Delete, "删除事件(&D)")
        
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_NewParent, "新增消息(&A)")
        cbrPopupItem.IconId = conMenu_Edit_NewItem
        cbrPopupItem.BeginGroup = True
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_ModifyParent, "修改消息(&U)")
        cbrPopupItem.IconId = conMenu_Edit_Modify
        
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_DeleteParent, "删除消息(&D)")
        cbrPopupItem.IconId = conMenu_Edit_Delete
        
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)")
        cbrPopupItem.BeginGroup = True
    
    End Select
    
    Set ShowConetneMenu = cbrPopupBar
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

