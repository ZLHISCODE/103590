VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmItem 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8265
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13185
   LinkTopic       =   "Form1"
   ScaleHeight     =   8265
   ScaleWidth      =   13185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txtLocation 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   0
      TabIndex        =   8
      Top             =   690
      Width           =   1575
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   0
      Left            =   7860
      ScaleHeight     =   240
      ScaleWidth      =   1950
      TabIndex        =   5
      Top             =   2505
      Width           =   1980
      Begin VB.ComboBox cboOwner 
         Height          =   300
         Left            =   -30
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   -30
         Width           =   2010
      End
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   1980
      Index           =   2
      Left            =   3450
      ScaleHeight     =   1980
      ScaleWidth      =   2700
      TabIndex        =   3
      Top             =   975
      Width           =   2700
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1215
         Index           =   0
         Left            =   120
         TabIndex        =   4
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
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   4245
      Index           =   1
      Left            =   495
      ScaleHeight     =   4245
      ScaleWidth      =   2370
      TabIndex        =   2
      Top             =   1680
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
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   2295
      Index           =   0
      Left            =   3525
      ScaleHeight     =   2295
      ScaleWidth      =   4245
      TabIndex        =   0
      Top             =   3285
      Width           =   4245
      Begin XtremeSuiteControls.TabControl tbc 
         Height          =   1590
         Left            =   450
         TabIndex        =   1
         Top             =   360
         Width           =   3165
         _Version        =   589884
         _ExtentX        =   5583
         _ExtentY        =   2805
         _StockProps     =   64
      End
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   7170
      Top             =   420
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItem.frx":0000
            Key             =   "file"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItem.frx":015A
            Key             =   "folder"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItem.frx":69BC
            Key             =   "folder_open"
         EndProperty
      EndProperty
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
   Begin XtremeCommandBars.ImageManager ImageManager1 
      Bindings        =   "frmItem.frx":D21E
      Left            =   2430
      Top             =   450
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmItem.frx":D232
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   375
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmItem"
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
    
    刷新项目数据
        
    新增信息
    修改信息
    删除信息
    内容配置
    刷新内容配置
    刷新触发条件
    刷新投递目标
    刷新指定信息
    移除指定信息
End Enum

Private mstrBusiness As String
Private mlngModualCode As Long
Private mstrSQL As String
Private mclsVsf(0) As zlVSFlexGrid.clsVsf
Private mblnStartUp As Boolean
Private mobjFindKey As CommandBarControl
Private mstrFindKey As String
Private mblnDataChanged As Boolean
Private mblnReading As Boolean
Private mrsCondition As ADODB.Recordset
Private mstrCurrentType As String

Private WithEvents mfrmItemEdit As frmItemEdit
Attribute mfrmItemEdit.VB_VarHelpID = -1
Private WithEvents mfrmItemConfigEdit As frmItemConfigEdit
Attribute mfrmItemConfigEdit.VB_VarHelpID = -1
Private mfrmItemTrigger As frmItemTrigger
Private mfrmItemConfig As frmItemConfig
Private mfrmItemDelivery As frmItemDelivery

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
    
    
    Set mrsCondition = zlCommFun.CreateCondition
    
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
    
        If mfrmItemEdit Is Nothing Then
            Set mfrmItemEdit = New frmItemEdit
            Call mfrmItemEdit.InitDialog(Me, mlngModualCode)
        End If

        Call mfrmItemEdit.NewData(mstrBusiness)
        DoEvents
        Me.SetFocus

    '------------------------------------------------------------------------------------------------------------------
    Case Command.修改信息
    
        If mfrmItemEdit Is Nothing Then
            Set mfrmItemEdit = New frmItemEdit
            Call mfrmItemEdit.InitDialog(Me, mlngModualCode)
        End If
        
        With vsf(0)
            Call mfrmItemEdit.ModifyData(mstrBusiness, .TextMatrix(.Row, .ColIndex("id")))
        End With
        DoEvents
        Me.SetFocus
    '------------------------------------------------------------------------------------------------------------------
    Case Command.删除信息
        
        If mfrmItemEdit Is Nothing Then
            Set mfrmItemEdit = New frmItemEdit
            Call mfrmItemEdit.InitDialog(Me, mlngModualCode)
        End If
        
        With vsf(0)
            If MsgBox("您确认要删除当前业务信息吗？", vbQuestion + vbYesNo + vbDefaultButton2, ParamInfo.系统名称) = vbYes Then
                Call mfrmItemEdit.DeleteData(mstrBusiness, .TextMatrix(.Row, .ColIndex("id")))
            End If
        End With
        
        DoEvents
        Me.SetFocus
    '------------------------------------------------------------------------------------------------------------------
    Case Command.内容配置
        
        If mfrmItemConfigEdit Is Nothing Then
            Set mfrmItemConfigEdit = New frmItemConfigEdit
            Call mfrmItemConfigEdit.InitDialog(Me, mlngModualCode)
        End If
        
        With vsf(0)
            Call mfrmItemConfigEdit.ConfigData(mstrBusiness, .TextMatrix(.Row, .ColIndex("id")))
        End With
        
        DoEvents
        Me.SetFocus
        
    '------------------------------------------------------------------------------------------------------------------
    Case Command.刷新项目数据
        
        With vsf(0)
            mclsVsf(0).SaveKey = Trim(.TextMatrix(.Row, .ColIndex("id")))
            mclsVsf(0).ClearGrid
            Set mrsCondition = zlCommFun.CreateCondition
            
            Call zlCommFun.SetCondition(mrsCondition, "FilterStyle", mstrFindKey)
            If Trim(txtLocation.Text) = "" Then
                Call zlCommFun.SetCondition(mrsCondition, "FilterText", "")
            Else
                Call zlCommFun.SetCondition(mrsCondition, "FilterText", Trim(txtLocation.Text))
            End If
            
            Select Case mstrCurrentType
            Case "T1"
                Call zlCommFun.SetCondition(mrsCondition, "item_type", "1")
            Case "T2"
                Call zlCommFun.SetCondition(mrsCondition, "item_type", "2")
            End Select
            
            Call zlCommFun.SetCondition(mrsCondition, "data_code", mstrBusiness)
            
            ExecuteCommand = mclsVsf(0).LoadDataSource(gclsBusiness.ItemRead("item_type", mrsCondition))
            
            Call mclsVsf(0).RestoreRow(mclsVsf(0).SaveKey, .ColIndex("id"))
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case Command.刷新内容配置
    
        With vsf(0)
            Call mfrmItemConfig.RefreshData(.TextMatrix(.Row, .ColIndex("id")))
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case Command.刷新触发条件
    
        With vsf(0)
            Call mfrmItemTrigger.RefreshData(.TextMatrix(.Row, .ColIndex("id")))
        End With
                
    '------------------------------------------------------------------------------------------------------------------
    Case Command.刷新投递目标
    
        With vsf(0)
            Call mfrmItemDelivery.RefreshData(.TextMatrix(.Row, .ColIndex("id")), (mstrCurrentType = "T1"))
        End With

        
    '------------------------------------------------------------------------------------------------------------------
    Case Command.刷新指定信息
        
        ExecuteCommand = LoadCustomData(Trim(varParam(0)))
        
        Call ExecuteCommand(Command.刷新触发条件)
        Call ExecuteCommand(Command.刷新内容配置)
        
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
        mstrCurrentType = tplItem.Tag
        
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
    
    Set rsData = gclsBusiness.ItemRead("id", rsCondition)
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
        Call .AppendColumn("项目标识", 2100, flexAlignLeftCenter, flexDTString, , "item_code", True)
        Call .AppendColumn("项目名称", 3000, flexAlignLeftCenter, flexDTString, , "item_title", True)
        Call .AppendColumn("触发类型", 1080, flexAlignLeftCenter, flexDTString, , "trigger_type", True)
        Call .AppendColumn("重发策略", 1500, flexAlignLeftCenter, flexDTString, , "again_policy_title", True)
        Call .AppendColumn("开始时间", 1080, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd", "start_date", True)
        Call .AppendColumn("结束时间", 1080, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd", "stop_date", True)
        Call .AppendColumn("备注说明", 1500, flexAlignLeftCenter, flexDTString, , "item_note", True)
        
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

Private Sub InitTabControl()
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************

    On Error GoTo errHand
    
'    ------------------------------------------------------------------------------------------------------------------
    Call TabControlInit(tbc, xtpTabAppearancePropertyPage2003)
    With tbc
'
'        With .PaintManager
'            .Appearance = xtpTabAppearancePropertyPage2003
'            .BoldSelected = True
''            .ClientFrame = xtpTabFrameSingleLine
'            .ShowIcons = True
'            .DisableLunaColors = False
'        End With
        
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .BoldSelected = True
        End With
        
        Set .Icons = zlCommFun.GetPubIcons
        
        Set mfrmItemTrigger = New frmItemTrigger
        Set mfrmItemConfig = New frmItemConfig
        Call mfrmItemConfig.InitForm(Me, 1002)
        
        Set mfrmItemDelivery = New frmItemDelivery
        Call mfrmItemDelivery.InitForm(Me)
        
        .InsertItem 0, "触发条件", mfrmItemTrigger.hWnd, 0
        .InsertItem 1, "内容构造", mfrmItemConfig.hWnd, 0
        .InsertItem 2, "投递目标", mfrmItemDelivery.hWnd, 0

        .Item(0).Selected = True
    End With

    Exit Sub

errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
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
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Compend, "内容")
        
    mstrFindKey = zlDataBase.GetPara("定位依据", ParamInfo.系统号, mlngModualCode, "项目标识")
    If mstrFindKey = "" Then mstrFindKey = "项目标识"

    Set mobjFindKey = zlCommFun.NewToolBar(objBar, xtpControlPopup, conMenu_View_LocationItem, mstrFindKey, False, , xtpButtonIconAndCaption)
    mobjFindKey.IconId = conMenu_View_Find
    mobjFindKey.flags = xtpFlagRightAlign
    mobjFindKey.BeginGroup = True
    mobjFindKey.Style = xtpButtonIconAndCaption
    Set objControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&1.项目标识"): objControl.Parameter = "项目标识"
    objControl.IconId = 99999999
    Set objControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&2.项目名称"): objControl.Parameter = "项目名称"
    objControl.IconId = 99999999
        
    Set cbrCustom = zlCommFun.NewToolBar(objBar, xtpControlCustom, 0, "")
    cbrCustom.Handle = txtLocation.hWnd
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Refresh, "刷新", True)
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
    objPane.Title = "信息"
    objPane.Options = PaneNoCaption
    
    Set objPane = dkpMain.CreatePane(2, 300, 100, DockRightOf, objPane)
    objPane.Title = "SQL"
    objPane.Options = PaneNoCaption
    
    Set objPane = dkpMain.CreatePane(3, 300, 200, DockBottomOf, objPane)
    objPane.Title = "关系"
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
    
    Call ExecuteCommand(Command.刷新项目数据)
    Call ExecuteCommand(Command.刷新触发条件)
    Call ExecuteCommand(Command.刷新内容配置)
    Call ExecuteCommand(Command.刷新投递目标)
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
    Case conMenu_Edit_Compend
        
        Call ExecuteCommand(Command.内容配置)
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Refresh               '刷新
        
        Call ExecuteCommand(Command.刷新项目数据)
        Call ExecuteCommand(Command.刷新触发条件)
        Call ExecuteCommand(Command.刷新内容配置)
        Call ExecuteCommand(Command.刷新投递目标)
        
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
'            If tvw.SelectedItem Is Nothing Then
'                Control.Enabled = False
'            Else
                Control.Enabled = (mstrCurrentType = "T2" And Control.Visible)
'            End If
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Modify               '修改
            
'            If tvw.SelectedItem Is Nothing Then
'                Control.Enabled = False
'            ElseIf mstrCurrentType = "T1" Then
'                Control.Enabled = False
'            Else
                Control.Enabled = (mstrCurrentType = "T2" And Trim(.TextMatrix(.Row, .ColIndex("id"))) <> "" And Control.Visible)
'            End If
                    
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Delete
            
'            If tvw.SelectedItem Is Nothing Then
'                Control.Enabled = False
'            ElseIf tvw.SelectedItem.Key = "T1" Then
'                Control.Enabled = False
'            Else
                Control.Enabled = (mstrCurrentType = "T2" And Trim(.TextMatrix(.Row, .ColIndex("id"))) <> "" And Control.Visible)
'            End If
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Compend
            
'            If tvw.SelectedItem Is Nothing Then
'                Control.Enabled = False
'            ElseIf tvw.SelectedItem.Key = "T1" Then
'                Control.Enabled = False
'            Else
                Control.Enabled = (mstrCurrentType = "T2" And Trim(.TextMatrix(.Row, .ColIndex("id"))) <> "" And Control.Visible)
'            End If
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
        Item.Handle = picPane(1).hWnd
    Case 2
        Item.Handle = picPane(2).hWnd
    Case 3
        Item.Handle = picPane(0).hWnd
    End Select
End Sub

Private Sub Form_Activate()
    If mblnStartUp = False Then Exit Sub
    DoEvents
    mblnStartUp = False
'    Call InitData
    Call ExecuteCommand(Command.初始数据)
    Call cboOwner_Click
End Sub

Private Sub Form_Load()
    mblnStartUp = True
    mlngModualCode = 1002
    
    Call ExecuteCommand(Command.初始控件)
    Call ExecuteCommand(Command.读注册表)

    Call zlComLib.RestoreWinState(Me, App.ProductName)
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, gblnShowInTaskBar)
    Call zlDataBase.ShowReportMenu(Me, ParamInfo.系统号, ParamInfo.模块号, UserInfo.模块权限)
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Call zlCommFun.SetPaneRange(dkpMain, 1, 100, 15, 100, Me.ScaleHeight)
    Call zlCommFun.SetPaneRange(dkpMain, 2, 15, 200, Me.ScaleWidth, 500)
    
    dkpMain.RecalcLayout
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mclsVsf(0) = Nothing
    Set mobjFindKey = Nothing
        
    If Not (mfrmItemEdit Is Nothing) Then
        Unload mfrmItemEdit
        Set mfrmItemEdit = Nothing
    End If
    
    If Not (mfrmItemTrigger Is Nothing) Then
        Unload mfrmItemTrigger
        Set mfrmItemTrigger = Nothing
    End If
    
    If Not (mfrmItemDelivery Is Nothing) Then
        Unload mfrmItemDelivery
        Set mfrmItemDelivery = Nothing
    End If
    
    If Not (mfrmItemConfig Is Nothing) Then
        Unload mfrmItemConfig
        Set mfrmItemConfig = Nothing
    End If
    
    If Not (mfrmItemConfigEdit Is Nothing) Then
        Unload mfrmItemConfigEdit
        Set mfrmItemConfigEdit = Nothing
    End If
    
    Set mrsCondition = Nothing
End Sub

Private Sub mfrmItemConfigEdit_AfterDeleteData(ByVal DataKey As String)
    '
End Sub

Private Sub mfrmItemConfigEdit_AfterModifyData(ByVal DataKey As String)
    Call ExecuteCommand(Command.刷新内容配置, DataKey)
End Sub

Private Sub mfrmItemConfigEdit_AfterNewData(ByVal DataKey As String)
    '
End Sub

Private Sub mfrmItemConfigEdit_Backward(DataKey As String, Cancel As Boolean)
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

Private Sub mfrmItemConfigEdit_Forward(DataKey As String, Cancel As Boolean)
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

Private Sub mfrmItemEdit_AfterDeleteData(ByVal DataKey As String)
    Call ExecuteCommand(Command.移除指定信息, DataKey)
End Sub

Private Sub mfrmItemEdit_AfterModifyData(ByVal DataKey As String)
    Call ExecuteCommand(Command.刷新指定信息, DataKey)
End Sub

Private Sub mfrmItemEdit_AfterNewData(ByVal DataKey As String)
    Call ExecuteCommand(Command.刷新指定信息, DataKey)
End Sub

Private Sub mfrmItemEdit_Backward(DataKey As String, Cancel As Boolean)
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

Private Sub mfrmItemEdit_Forward(DataKey As String, Cancel As Boolean)
    
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
'        tvw.Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
        tpl.Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
    Case 2
        vsf(0).Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
        mclsVsf(0).AppendRows = True
    End Select
End Sub

Private Sub tpl_ItemClick(ByVal Item As XtremeSuiteControls.ITaskPanelGroupItem)
    mstrCurrentType = Item.Tag
    Call ExecuteCommand(Command.刷新项目数据)
    Call ExecuteCommand(Command.刷新触发条件)
    Call ExecuteCommand(Command.刷新内容配置)
    Call ExecuteCommand(Command.刷新投递目标)
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
        Call ExecuteCommand(Command.刷新触发条件)
        Call ExecuteCommand(Command.刷新内容配置)
        Call ExecuteCommand(Command.刷新投递目标)
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
                Call ShowConetneMenu(2).ShowPopup
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
    Case 2  '
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_NewItem, "新增项目(&N)")
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Modify, "修改项目(&M)")
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Delete, "删除项目(&D)")
        
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Compend, "内容构造(&C)")
        cbrPopupItem.BeginGroup = True
        
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

