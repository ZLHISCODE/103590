VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Begin VB.Form frmStation 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8025
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12735
   LinkTopic       =   "Form1"
   ScaleHeight     =   8025
   ScaleWidth      =   12735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txtLocation 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   5025
      TabIndex        =   2
      Top             =   180
      Width           =   1575
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   1980
      Index           =   2
      Left            =   3450
      ScaleHeight     =   1980
      ScaleWidth      =   2700
      TabIndex        =   0
      Top             =   975
      Width           =   2700
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1215
         Index           =   0
         Left            =   120
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
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmStation.frx":0000
      Left            =   375
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmStation"
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
    设置消息用户
    应用消息用户
    刷新站点
End Enum

Private mlngModualCode As Long
Private mstrSQL As String
Private mclsVsf(0) As zlVSFlexGrid.clsVsf
Private mblnStartUp As Boolean
Private mobjFindKey As CommandBarControl
Private mstrFindKey As String
Private mblnDataChanged As Boolean
Private WithEvents mfrmStationUser As frmStationUser
Attribute mfrmStationUser.VB_VarHelpID = -1

Public Event AfterClose()
Public Event AfterLoad(ByVal intIndex As Integer, ByVal strContent As String)

'######################################################################################################################
'接口方法
Public Function Execute()
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
'    Call Form_Activate
    Call ExecuteCommand(Command.刷新站点)
    
End Function

'######################################################################################################################
Private Function ExecuteCommand(ByVal enmCommand As Command, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim rsPara As ADODB.Recordset
    Dim rsCondition As ADODB.Recordset
    Dim strTmp As String
    Dim strSQL As String
    Dim intRow As Integer
    Dim varTmp As Variant
    Dim blnMuliSelect As Boolean
    
    On Error GoTo errHand
    
    
    
    Select Case enmCommand
    '------------------------------------------------------------------------------------------------------------------
    Case Command.初始控件
        
        Call InitGrid
        Call InitCommandBar
        Call InitDockPannel
    '------------------------------------------------------------------------------------------------------------------
    Case Command.设置消息用户
        
        With vsf(0)
            If .TextMatrix(.Row, .ColIndex("消息用户")) = "<默认>" Then
                Call mfrmStationUser.ShowDialog(Me, .TextMatrix(.Row, .ColIndex("工作站")), "", "")
            Else
                Call mfrmStationUser.ShowDialog(Me, .TextMatrix(.Row, .ColIndex("工作站")), .TextMatrix(.Row, .ColIndex("消息用户")), .TextMatrix(.Row, .ColIndex("消息密码")))
            End If
            
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case Command.应用消息用户
        
        Set rsPara = zlCommFun.CreateParameter
        
        With vsf(0)
            If MsgBox("您确认要将当前工作站(" & .TextMatrix(.Row, .ColIndex("工作站")) & ")的消息用户同步到已经勾选的工作站吗？", vbQuestion + vbYesNo + vbDefaultButton2, ParamInfo.系统名称) = vbYes Then
                For intRow = 1 To .Rows - 1
                    If Val(Abs(.TextMatrix(intRow, .ColIndex("选择")))) = 1 And .TextMatrix(intRow, .ColIndex("工作站")) <> "" Then
                        Call zlCommFun.SetParameter(rsPara, "来源工作站", .TextMatrix(.Row, .ColIndex("工作站")))
                        Call zlCommFun.SetParameter(rsPara, "目标工作站", .TextMatrix(intRow, .ColIndex("工作站")))
                        If gclsBusiness.ClientsEdit("Copy", rsPara) Then
                            .TextMatrix(intRow, .ColIndex("消息用户")) = .TextMatrix(.Row, .ColIndex("消息用户"))
                        End If
                    End If
                Next
            End If
 
        End With
'        Call ExecuteCommand(Command.刷新站点)
    '------------------------------------------------------------------------------------------------------------------
    Case Command.刷新站点
        

        With vsf(0)
            mclsVsf(0).SaveKey = Trim(.TextMatrix(.Row, .ColIndex("id")))
            
            If Trim(txtLocation.Text) = "" Then
                ExecuteCommand = mclsVsf(0).LoadDataSource(gclsBusiness.ClientRead())
            Else
                Set rsCondition = zlCommFun.CreateCondition
                Call zlCommFun.SetCondition(rsCondition, "FilterStyle", mstrFindKey)
                Call zlCommFun.SetCondition(rsCondition, "FilterText", Trim(txtLocation.Text))
                ExecuteCommand = mclsVsf(0).LoadDataSource(gclsBusiness.ClientRead("FilterData", rsCondition))
            End If
      
            Call mclsVsf(0).RestoreRow(mclsVsf(0).SaveKey, .ColIndex("id"))
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

Private Function InitGrid() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    On Error GoTo errHand
    
    Set mclsVsf(0) = New zlVSFlexGrid.clsVsf
    With mclsVsf(0)
        Call .Initialize(Me.Controls, vsf(0), True, True, gclsBusiness.GetImageList(16))
        Call .ClearColumn
        
        Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[序号]", False)
        Call .AppendColumn("", 300, flexAlignCenterCenter, flexDTBoolean, "", "[选择]", False)
        Call .AppendColumn("", 300, flexAlignCenterCenter, flexDTString, "", "[图标]", False)
        Call .AppendColumn("id", 0, flexAlignLeftCenter, flexDTString, , "id", True)
        Call .AppendColumn("工作站", 2100, flexAlignLeftCenter, flexDTString, , "工作站", True)
        Call .AppendColumn("IP", 1500, flexAlignLeftCenter, flexDTString, , "IP", True)
        Call .AppendColumn("操作系统", 3000, flexAlignLeftCenter, flexDTString, , "操作系统", True)
        Call .AppendColumn("部门", 1500, flexAlignLeftCenter, flexDTString, , "部门", True)
        Call .AppendColumn("消息用户", 1080, flexAlignLeftCenter, flexDTString, , "消息用户", True)
        Call .AppendColumn("消息密码", 0, flexAlignLeftCenter, flexDTString, , "消息密码", True, , , True)
        Call .AppendColumn("说明", 1500, flexAlignLeftCenter, flexDTString, , "说明", True)
        
        .IndicatorMode = 2
        .IndicatorCol = .ColIndex("序号")
        .ConstCol = .ColIndex("序号")
        
        Call .InitializeEdit(True, False, False)
        Call .InitializeEditColumn(.ColIndex("选择"), True, vbVsfEditCheck)
        
    End With
            
    InitGrid = True
    
    Exit Function

errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
    
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
    cbsMain.VisualTheme = xtpThemeWhidbey
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
        
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_File_Parameter, "配置", True)
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_SelAll, "全选", True)
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_ClsAll, "全清")
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Modify, "消息用户设置", True)
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_ApplyTo, "消息用户应用")
        
    mstrFindKey = zlDataBase.GetPara("定位依据", ParamInfo.系统号, mlngModualCode, "工作站")
    If mstrFindKey = "" Then mstrFindKey = "工作站"

    Set mobjFindKey = zlCommFun.NewToolBar(objBar, xtpControlPopup, conMenu_View_LocationItem, mstrFindKey, True, , xtpButtonIconAndCaption)
    mobjFindKey.IconId = conMenu_View_Find
    mobjFindKey.flags = xtpFlagRightAlign
    mobjFindKey.Style = xtpButtonIconAndCaption
    Set objControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&1.工作站"): objControl.Parameter = "工作站"
    objControl.IconId = 1
    Set objControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&2.IP"): objControl.Parameter = "IP"
    objControl.IconId = 1
    Set objControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&2.部门"): objControl.Parameter = "部门"
    objControl.IconId = 1

    Set cbrCustom = zlCommFun.NewToolBar(objBar, xtpControlCustom, 0, "")
    cbrCustom.Handle = txtLocation.hWnd
        
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Refresh, "刷新")

    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_File_Close, "关闭")
    objControl.flags = xtpFlagRightAlign
    
    '------------------------------------------------------------------------------------------------------------------
    '命令的快键绑定:公共部份主界面已处理

    With cbsMain.KeyBindings
        .Add 0, vbKeyF5, conMenu_View_Refresh           '刷新
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
    objPane.Title = "信息"
    objPane.Options = PaneNoCaption
    
    dkpMain.SetCommandBars cbsMain
    Call zlCommFun.DockPannelInit(dkpMain)

End Sub


Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)

    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Parameter
        Call frmStationParameter.ShowConfigDialog(Me)
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
    Case conMenu_Edit_Modify
    
        Call ExecuteCommand(Command.设置消息用户)
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_ApplyTo
        
        Call ExecuteCommand(Command.应用消息用户)
                
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Refresh               '刷新
                
        Call ExecuteCommand(Command.刷新站点)
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_LocationItem
        mstrFindKey = Control.Parameter
        mobjFindKey.Caption = mstrFindKey
        cbsMain.RecalcLayout
        
    Case conMenu_File_Close
    '--------------------------------------------------------------------------------------------------------------
'        Unload Me
        RaiseEvent AfterClose
    End Select
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim intRow As Integer
    Dim blnMuliSelect As Boolean
    
    With vsf(0)
        Select Case Control.ID
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Delete
    
            Control.Enabled = (Trim(.TextMatrix(.Row, .ColIndex("id"))) <> "" And Control.Visible)
                    
        End Select
    End With
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = picPane(2).hWnd
    End Select
End Sub

Private Sub Form_Activate()
    If mblnStartUp = False Then Exit Sub
    DoEvents
    mblnStartUp = False
    
    
End Sub

Private Sub Form_Load()
    mblnStartUp = True
    mlngModualCode = 1003
    
    Call ExecuteCommand(Command.初始控件)
    Call ExecuteCommand(Command.读注册表)

    Call zlComLib.RestoreWinState(Me, App.ProductName)
'    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, gblnShowInTaskBar)
'    Call zlDataBase.ShowReportMenu(Me, ParamInfo.系统号, ParamInfo.模块号, UserInfo.模块权限)
    Set mfrmStationUser = New frmStationUser
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mclsVsf(0) = Nothing
    Set mobjFindKey = Nothing
    If Not (mfrmStationUser Is Nothing) Then
        Unload mfrmStationUser
        Set mfrmStationUser = Nothing
    End If
End Sub

Private Sub mfrmStationUser_AfterDataChanged(ByVal strStation As String, ByVal strMipUser As String, ByVal strMipUserPassword As String)
    
    Dim intRow As Integer
    
    With mclsVsf(0)
        intRow = .FindRow(strStation, .ColIndex("工作站"))
        
        If intRow > 0 Then
            '
            If strMipUser = "" Then
                .TextMatrix(intRow, .ColIndex("消息用户")) = "<默认>"
                .TextMatrix(intRow, .ColIndex("消息密码")) = ""
            Else
                .TextMatrix(intRow, .ColIndex("消息用户")) = strMipUser
                .TextMatrix(intRow, .ColIndex("消息密码")) = strMipUserPassword
            End If
            
        End If
    End With
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 2
        vsf(0).Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
        mclsVsf(0).AppendRows = True
    End Select
End Sub

Private Sub txtLocation_KeyPress(KeyAscii As Integer)
    Dim strText As String
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim rsCondition As ADODB.Recordset
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0

'        If txtLocation.Text <> "" Then
            txtLocation.Tag = ""
            
            Dim obj As CommandBarControl
            
            Set obj = cbsMain.FindControl(, conMenu_View_Refresh, True)
            If obj Is Nothing Then Exit Sub
            
            If obj.Enabled = True Then
                Call cbsMain_Execute(obj)
            End If

'        End If
'        txtLocation.Tag = ""
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub vsf_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Call mclsVsf(Index).AfterEdit(Row, Col)
End Sub

Private Sub vsf_AfterMoveColumn(Index As Integer, ByVal Col As Long, Position As Long)
    mclsVsf(Index).AppendRows = True
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsf(Index).AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
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

Private Sub vsf_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call mclsVsf(Index).KeyDown(KeyCode, Shift)
End Sub

Private Sub vsf_KeyPress(Index As Integer, KeyAscii As Integer)
    Call mclsVsf(Index).KeyPress(KeyAscii)
End Sub

Private Sub vsf_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Call mclsVsf(Index).KeyPressEdit(KeyAscii)
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
    '------------------------------------------------------------------------------------------------------------------
    Case 1  '

        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_SelAll, "全部勾选(&A)")
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_ClsAll, "全部不选(&U)")
        
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Modify, "用户设置(&M)")
        cbrPopupItem.BeginGroup = True
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_ApplyTo, "用户应用(&A)")
        
    End Select
    
    Set ShowConetneMenu = cbrPopupBar
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
End Function


Private Sub vsf_SetupEditWindow(Index As Integer, ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    Call mclsVsf(0).EditSelAll
End Sub

Private Sub vsf_StartEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf(0).BeforeEdit(Row, Col, Cancel)
End Sub

Private Sub vsf_ValidateEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf(0).ValidateEdit(Col, Cancel)
End Sub
