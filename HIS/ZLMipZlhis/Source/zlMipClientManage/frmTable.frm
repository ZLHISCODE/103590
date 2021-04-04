VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmTable 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11355
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   11355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txtLocation 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   0
      TabIndex        =   8
      Top             =   1575
      Width           =   1575
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   4245
      Index           =   2
      Left            =   330
      ScaleHeight     =   4245
      ScaleWidth      =   2370
      TabIndex        =   6
      Top             =   2085
      Width           =   2370
      Begin XtremeSuiteControls.TaskPanel tpl 
         Height          =   4770
         Left            =   345
         TabIndex        =   7
         Top             =   495
         Width           =   3210
         _Version        =   589884
         _ExtentX        =   5662
         _ExtentY        =   8414
         _StockProps     =   64
         Behaviour       =   1
         ItemLayout      =   2
         HotTrackStyle   =   3
      End
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   0
      Left            =   0
      ScaleHeight     =   240
      ScaleWidth      =   1950
      TabIndex        =   4
      Top             =   555
      Width           =   1980
      Begin VB.ComboBox cboOwner 
         Height          =   300
         Left            =   -30
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   -30
         Width           =   2010
      End
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   2685
      Index           =   0
      Left            =   5415
      ScaleHeight     =   2685
      ScaleWidth      =   4590
      TabIndex        =   2
      Top             =   150
      Width           =   4590
      Begin XtremeSuiteControls.TabControl tbc 
         Height          =   1590
         Left            =   450
         TabIndex        =   3
         Top             =   360
         Width           =   3165
         _Version        =   589884
         _ExtentX        =   5583
         _ExtentY        =   2805
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   2760
      Index           =   1
      Left            =   2235
      ScaleHeight     =   2760
      ScaleWidth      =   2865
      TabIndex        =   0
      Top             =   75
      Width           =   2865
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1215
         Index           =   0
         Left            =   270
         TabIndex        =   1
         Top             =   225
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
   Begin XtremeCommandBars.ImageManager ImageManager1 
      Left            =   1425
      Top             =   1155
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmTable.frx":0000
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
      Bindings        =   "frmTable.frx":1326
      Left            =   375
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmTable"
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
    初始数据
    新增信息
    修改信息
    删除信息
    刷新数据
    刷新附加数据
    刷新指定信息
    移除指定信息
End Enum

Private mlngModualCode As Long
Private mstrSQL As String
Private mclsVsf(0) As zlVSFlexGrid.clsVsf
Private mblnStartUp As Boolean
Private mobjFindKey As CommandBarControl
Private mstrFindKey As String
Private mblnDataChanged As Boolean
Private mrsCondition As ADODB.Recordset
Private mblnReading As Boolean
Private WithEvents mfrmTableEdit As frmTableEdit
Attribute mfrmTableEdit.VB_VarHelpID = -1
Private mfrmTableSQL As frmTableSQL
Private mfrmTableRelation As frmTableRelation
Private mstrBusiness As String
Private mstrCurrentGroup As String

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
    Call Form_Activate
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
        Call InitTabControl
        Call InitTaskPanel
        
    '------------------------------------------------------------------------------------------------------------------
    Case Command.初始数据
        
        mblnReading = True
        cboOwner.Clear
        cboOwner.AddItem "公共"
        cboOwner.ItemData(cboOwner.NewIndex) = 0
        If gclsBusiness.IsDBAUser Then
            
            Set rs = gclsBusiness.GetBusinessKind(True, "")
            If rs.BOF = False Then
                Do While Not rs.EOF
                    cboOwner.AddItem UCase(rs("业务").Value)
                    cboOwner.ItemData(cboOwner.NewIndex) = 1
                    rs.MoveNext
                Loop
            End If
            If cboOwner.ListCount > 0 And cboOwner.ListIndex = -1 Then cboOwner.ListIndex = 1
        Else
            Set rs = gclsBusiness.GetBusinessKind(True, gstrDbUser)
            If rs.BOF = False Then
                Do While Not rs.EOF
                    cboOwner.AddItem UCase(rs("业务").Value)
                    cboOwner.ItemData(cboOwner.NewIndex) = 1
                    rs.MoveNext
                Loop
            End If
            If cboOwner.ListCount > 0 And cboOwner.ListIndex = -1 Then cboOwner.ListIndex = 1
        End If
        If cboOwner.ListCount > 0 And cboOwner.ListIndex = -1 Then cboOwner.ListIndex = 0
        mblnReading = False
        
    '------------------------------------------------------------------------------------------------------------------
    Case Command.新增信息
    
        If mfrmTableEdit Is Nothing Then
            Set mfrmTableEdit = New frmTableEdit
            Call mfrmTableEdit.InitDialog(Me, mlngModualCode)
        End If
    
        Call mfrmTableEdit.NewData(mstrBusiness)
        DoEvents
        Me.SetFocus
    '------------------------------------------------------------------------------------------------------------------
    Case Command.修改信息
    
        If mfrmTableEdit Is Nothing Then
            Set mfrmTableEdit = New frmTableEdit
            Call mfrmTableEdit.InitDialog(Me, mlngModualCode)
        End If
        
        With vsf(0)
            Call mfrmTableEdit.ModifyData(mstrBusiness, .TextMatrix(.Row, .ColIndex("id")))
        End With
        DoEvents
        Me.SetFocus
    '------------------------------------------------------------------------------------------------------------------
    Case Command.删除信息
        
        If mfrmTableEdit Is Nothing Then
            Set mfrmTableEdit = New frmTableEdit
            Call mfrmTableEdit.InitDialog(Me, mlngModualCode)
        End If
        
        With vsf(0)
            If MsgBox("您确认要删除当前业务信息吗？", vbQuestion + vbYesNo + vbDefaultButton2, ParamInfo.系统名称) = vbYes Then
                Call mfrmTableEdit.DeleteData(mstrBusiness, .TextMatrix(.Row, .ColIndex("id")))
            End If
        End With
        
        DoEvents
        Me.SetFocus
            
    '------------------------------------------------------------------------------------------------------------------
    Case Command.刷新数据
        
        With vsf(0)
            mclsVsf(0).SaveKey = Trim(.TextMatrix(.Row, .ColIndex("id")))
            
            mclsVsf(0).ClearGrid
            
            Set mrsCondition = zlCommFun.CreateCondition
            Call zlCommFun.SetCondition(mrsCondition, "data_code", mstrBusiness)
            
            
            Call zlCommFun.SetCondition(mrsCondition, "FilterStyle", mstrFindKey)
            If Trim(txtLocation.Text) = "" Then
                Call zlCommFun.SetCondition(mrsCondition, "FilterText", "")
            Else
                Call zlCommFun.SetCondition(mrsCondition, "FilterText", Trim(txtLocation.Text))
            End If
            
            Select Case mstrCurrentGroup
            Case "T1"
                Call zlCommFun.SetCondition(mrsCondition, "tab_type", "1")
            Case "T2"
                Call zlCommFun.SetCondition(mrsCondition, "tab_type", "2")
            End Select
            
            ExecuteCommand = mclsVsf(0).LoadDataSource(gclsBusiness.TableRead("", mrsCondition))
            
            Call mclsVsf(0).RestoreRow(mclsVsf(0).SaveKey, .ColIndex("id"))
        End With
                
        Call ExecuteCommand(Command.刷新附加数据)
        
    '------------------------------------------------------------------------------------------------------------------
    Case Command.刷新指定信息
        
        ExecuteCommand = LoadCustomData(Trim(varParam(0)))
        
        Call ExecuteCommand(Command.刷新附加数据)
        
    '-----------------------------------------------------------------------------------------------------------------
    Case Command.刷新附加数据
        
        With vsf(0)
            Call mfrmTableSQL.RefreshData(.TextMatrix(.Row, .ColIndex("id")))
            Call mfrmTableRelation.RefreshData(mstrBusiness, .TextMatrix(.Row, .ColIndex("id")))
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case Command.移除指定信息
        
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
    
    Set rsData = gclsBusiness.TableRead("id", rsCondition)
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
    On Error GoTo errHand
    
    Set mclsVsf(0) = New zlVSFlexGrid.clsVsf
    With mclsVsf(0)
        Call .Initialize(Me.Controls, vsf(0), True, False, gclsBusiness.GetImageList(16))
        Call .ClearColumn
        Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[序号]", False)
        Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[图标]", False)
        Call .AppendColumn("id", 0, flexAlignLeftCenter, flexDTString, , "id", True)
        Call .AppendColumn("编码", 2100, flexAlignLeftCenter, flexDTString, , "tab_code", True)
        Call .AppendColumn("名称", 2700, flexAlignLeftCenter, flexDTString, , "tab_title", True)
        Call .AppendColumn("注释", 1800, flexAlignLeftCenter, flexDTString, , "tab_note", True)
        
        .IndicatorMode = 2
        .IndicatorCol = .ColIndex("序号")
        .ConstCol = .ColIndex("序号")
        .AppendRows = True
        
    End With
            
    InitGrid = True
    
    Exit Function

errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
    
End Function

'Private Function InitData() As Boolean
'    '******************************************************************************************************************
'    '功能：
'    '参数：
'    '返回：
'    '******************************************************************************************************************
'    Dim rsTmp As ADODB.Recordset
'
'    On Error GoTo errHand
'
''    mlngSys = 0
'
'    cboSystem.Clear
''    cboSystem.AddItem "所有系统共享"
''    cboSystem.ItemData(cboSystem.NewIndex) = 0
'
'    Set rsTmp = gclsBusiness.GetSystem(gblnDBA, UCase(UserInfo.用户名))
'    If rsTmp.BOF = False Then
'        Do While Not rsTmp.EOF
'            cboSystem.AddItem Right(Space(4) & rsTmp("编号").Value, 4) & " - " & rsTmp("名称").Value
'            cboSystem.ItemData(cboSystem.NewIndex) = rsTmp("编号").Value
'            rsTmp.MoveNext
'        Loop
'    End If
'
'    If cboSystem.ListCount > 0 Then
'        cboSystem.ListIndex = 0
'    End If
'
'    Exit Function
'    '------------------------------------------------------------------------------------------------------------------
'errHand:
'    If zlComLib.ErrCenter = 1 Then
'        Resume
'    End If
'End Function

Private Sub InitTabControl()
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************

    On Error GoTo errHand

    '------------------------------------------------------------------------------------------------------------------
    Call TabControlInit(tbc, xtpTabAppearancePropertyPage2003)
    With tbc
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .BoldSelected = True
        End With
        Set mfrmTableSQL = New frmTableSQL
        Set mfrmTableRelation = New frmTableRelation
        Call mfrmTableRelation.InitForm(Me)
        .InsertItem 0, "基本资料", mfrmTableSQL.hWnd, 0
        .InsertItem 1, "信息项目", mfrmTableRelation.hWnd, 0
        .Item(0).Selected = True
    End With

    Exit Sub

errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitTaskPanel()
    
    Dim tplGroup As TaskPanelGroup
    Dim tplItem As TaskPanelGroupItem
    
    With tpl
        .SetIconSize 24, 24
        Call .Icons.AddIcons(ImageManager1.Icons)
        .VisualTheme = xtpTaskPanelThemeNativeWinXP
        .Behaviour = xtpTaskPanelBehaviourToolbox
        .ItemLayout = xtpTaskItemLayoutImagesWithTextBelow
        
        .SetMargins 5, 5, 5, 5, 5
        .SetItemInnerMargins 0, 5, 0, 5
        .SelectItemOnFocus = True
        
                
        Set tplGroup = .Groups.Add(0, "类型")
        tplGroup.Expandable = False
        tplGroup.CaptionVisible = False
        
        Set tplItem = tplGroup.Items.Add(1, "系统内置", xtpTaskItemTypeLink, 3)
        tplItem.Tag = "T1"
        tplItem.Tooltip = "系统固定配置的消息"
        
        tplItem.Selected = True
        mstrCurrentGroup = tplItem.Tag
        
        Set tplItem = tplGroup.Items.Add(2, "用户定义", xtpTaskItemTypeLink, 2)
        tplItem.Tag = "T2"
        tplItem.Tooltip = "用户自己定义的消息"
        
        .Reposition
    
    End With
    
    Exit Sub

errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
    Call zlComLib.SaveErrLog
End Sub

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
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlLabel, conMenu_Object_DBUser, "业务", , , xtpButtonCaption)
    objControl.BeginGroup = True
    Set cbrCustom = zlCommFun.NewToolBar(objBar, xtpControlCustom, conMenu_View_Location, "")
    cbrCustom.Handle = picBack(0).hWnd
        
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_NewItem, "新增", True)
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Modify, "修改")
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Delete, "删除")
    
    mstrFindKey = zlDataBase.GetPara("定位依据", ParamInfo.系统号, mlngModualCode, "编码")
    If mstrFindKey = "" Then mstrFindKey = "编码"

    Set mobjFindKey = zlCommFun.NewToolBar(objBar, xtpControlPopup, conMenu_View_LocationItem, mstrFindKey, False, , xtpButtonIconAndCaption)
    mobjFindKey.IconId = conMenu_View_Find
    mobjFindKey.flags = xtpFlagRightAlign
    mobjFindKey.BeginGroup = True
    mobjFindKey.Style = xtpButtonIconAndCaption
    Set objControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&1.编码"): objControl.Parameter = "编码"
    objControl.IconId = 99999999
    Set objControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&2.名称"): objControl.Parameter = "名称"
    objControl.IconId = 99999999
        
    Set cbrCustom = zlCommFun.NewToolBar(objBar, xtpControlCustom, 0, "")
    cbrCustom.Handle = txtLocation.hWnd
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Refresh, "刷新")
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_File_Close, "关闭")
    objControl.flags = xtpFlagRightAlign
    
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
    objPane.Title = "分组"
    objPane.Options = PaneNoCaption
    
    Set objPane = dkpMain.CreatePane(2, 300, 100, DockRightOf, objPane)
    objPane.Title = "SQL"
    objPane.Options = PaneNoCaption
    
    Set objPane = dkpMain.CreatePane(3, 300, 100, DockBottomOf, objPane)
    objPane.Title = "SQL"
    objPane.Options = PaneNoCaption
    
    dkpMain.SetCommandBars cbsMain
    Call zlCommFun.DockPannelInit(dkpMain)

End Sub

Private Sub cboOwner_Click()
    If mblnReading Then Exit Sub
    
    mstrBusiness = cboOwner.List(cboOwner.ListIndex)
    If InStr(mstrBusiness, "-") > 0 Then
        mstrBusiness = Trim(Mid(mstrBusiness, 1, InStr(mstrBusiness, "-") - 1))
    Else
        mstrBusiness = "-"
    End If
    
    Call ExecuteCommand(Command.刷新数据)
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)

    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_NewItem               '增加
        
        Call ExecuteCommand(Command.新增信息)
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Modify               '修改
        
        Call ExecuteCommand(Command.修改信息)
                
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete

        Call ExecuteCommand(Command.删除信息)
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_NewParent, conMenu_Edit_ModifyParent, conMenu_Edit_DeleteParent
        
        Call mfrmTableRelation.Execute(Control)
                        
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
            
            Control.Enabled = (mstrCurrentGroup = "T2" And Control.Visible)
            
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Modify               '修改
            
            Control.Enabled = (mstrCurrentGroup = "T2" And Trim(.TextMatrix(.Row, .ColIndex("id"))) <> "" And Control.Visible)
                    
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Delete
    
            Control.Enabled = (mstrCurrentGroup = "T2" And Trim(.TextMatrix(.Row, .ColIndex("id"))) <> "" And Control.Visible)
            
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_NewParent, conMenu_Edit_ModifyParent, conMenu_Edit_DeleteParent
            
            Call mfrmTableRelation.Update(Control)
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_LocationItem
            Control.Checked = (mstrFindKey = Control.Parameter)
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
        Item.Handle = picPane(2).hWnd
    Case 2
        Item.Handle = picPane(1).hWnd
    Case 3
        Item.Handle = picPane(0).hWnd
    End Select
End Sub

Private Sub Form_Activate()
    If mblnStartUp = False Then Exit Sub
    DoEvents
    mblnStartUp = False
    
    Call ExecuteCommand(Command.初始数据)
    
    Call cboOwner_Click
    
'    Call ExecuteCommand(Command.刷新数据)
End Sub

Private Sub Form_Load()
    mblnStartUp = True
    mlngModualCode = 1001
    
    Call ExecuteCommand(Command.初始控件)
    Call ExecuteCommand(Command.读注册表)

    Call zlComLib.RestoreWinState(Me, App.ProductName)
        
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Call zlCommFun.SetPaneRange(dkpMain, 1, 100, 15, 100, Me.ScaleHeight)
'    Call zlCommFun.SetPaneRange(dkpMain, 2, 100, 15, 300, Me.ScaleHeight)
        
    dkpMain.RecalcLayout
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mclsVsf(0) = Nothing
    Set mobjFindKey = Nothing
    
    If Not (mfrmTableEdit Is Nothing) Then
        Unload mfrmTableEdit
        Set mfrmTableEdit = Nothing
    End If
    
    If Not (mfrmTableSQL Is Nothing) Then
        Unload mfrmTableSQL
        Set mfrmTableSQL = Nothing
    End If
    
    If Not (mfrmTableRelation Is Nothing) Then
        Unload mfrmTableRelation
        Set mfrmTableRelation = Nothing
    End If
    
    Set mrsCondition = Nothing
End Sub

Private Sub mfrmTableEdit_AfterDeleteData(ByVal DataKey As String)
    Call ExecuteCommand(Command.移除指定信息, DataKey)
End Sub

Private Sub mfrmTableEdit_AfterModifyData(ByVal DataKey As String)
    Call ExecuteCommand(Command.刷新指定信息, DataKey)
End Sub

Private Sub mfrmTableEdit_AfterNewData(ByVal DataKey As String)
    Call ExecuteCommand(Command.刷新指定信息, DataKey)
End Sub

Private Sub mfrmTableEdit_Backward(DataKey As String, Cancel As Boolean)
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

Private Sub mfrmTableEdit_Forward(DataKey As String, Cancel As Boolean)
    
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
    Case 0
        tbc.Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
    Case 1
        vsf(0).Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
        mclsVsf(0).AppendRows = True
    Case 2
        tpl.Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
    End Select
End Sub

Private Sub tpl_ItemClick(ByVal Item As XtremeSuiteControls.ITaskPanelGroupItem)
    mstrCurrentGroup = Item.Tag
    Call ExecuteCommand(Command.刷新数据)
End Sub

Private Sub txtLocation_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0

        txtLocation.Tag = ""
        
        Dim obj As CommandBarControl
        
        Set obj = cbsMain.FindControl(, conMenu_View_Refresh, True)
        If obj Is Nothing Then Exit Sub
        If obj.Enabled = True Then
            Call cbsMain_Execute(obj)
        End If
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub vsf_AfterMoveColumn(Index As Integer, ByVal Col As Long, Position As Long)
    mclsVsf(Index).AppendRows = True
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsf(Index).AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
    If OldRow <> NewRow Then
        Call ExecuteCommand(Command.刷新附加数据)
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
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_NewItem, "新增信息(&N)")
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Modify, "修改信息(&M)")
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Delete, "删除信息(&D)")
        
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_View_Refresh, "重新刷新(&R)")
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
