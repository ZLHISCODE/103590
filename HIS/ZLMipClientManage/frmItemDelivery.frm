VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Begin VB.Form frmItemDelivery 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   10080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   2685
      Index           =   0
      Left            =   5085
      ScaleHeight     =   2685
      ScaleWidth      =   4335
      TabIndex        =   2
      Top             =   2820
      Width           =   4335
      Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
         Height          =   1215
         Left            =   255
         TabIndex        =   3
         Top             =   240
         Width           =   3465
         _cx             =   6112
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
         BackColorFixed  =   14737632
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
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
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
         AutoSizeMode    =   1
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   6
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
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   2685
      Index           =   1
      Left            =   195
      ScaleHeight     =   2685
      ScaleWidth      =   4335
      TabIndex        =   0
      Top             =   1140
      Width           =   4335
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1215
         Index           =   0
         Left            =   255
         TabIndex        =   1
         Top             =   240
         Width           =   3465
         _cx             =   6112
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
      Bindings        =   "frmItemDelivery.frx":0000
      Left            =   690
      Top             =   180
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmItemDelivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum Command
    初始控件
    读注册表
    
    刷新投递数据
    
    新增投递目标
    修改投递目标
    删除投递目标
    显示投递目标
    
    刷新指定目录
        
    新增信息
    修改信息
    删除信息
    内容配置
    刷新附加数据
    刷新指定信息
    移除指定信息
End Enum

Private mlngModualCode As Long
Private mstrPrivs As String
Private mstrSQL As String
Private mclsVsf(0) As zlVSFlexGrid.clsVsf
Private mclsVsfDetail As zlVSFlexGrid.clsVsf

Private mblnStartUp As Boolean
Private mlngTmp As Long
Private mblnShowAll As Boolean
Private mblnShowStop As Boolean
Private mobjFindKey As CommandBarControl
Private mstrFindKey As String
Private mblnDataChanged As Boolean
Private mblnNew As Boolean
Private mfrmParent As Object
Private mstrDataKey As String
Private mblnSystem As Boolean

Private WithEvents mfrmItemDeliveryEdit As frmItemDeliveryEdit
Attribute mfrmItemDeliveryEdit.VB_VarHelpID = -1

Public Function InitForm(ByVal frmParent As Object) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Set mfrmParent = frmParent
        
        
    InitForm = True
    
End Function

Public Function RefreshData(ByVal strDataKey As String, Optional ByVal blnSystem As Boolean) As Boolean

    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Dim rsTmp As ADODB.Recordset
    Dim rsCondition As ADODB.Recordset
    
    mblnSystem = blnSystem
    mstrDataKey = strDataKey
    
    Set rsCondition = zlCommFun.CreateCondition
    Call zlCommFun.SetCondition(rsCondition, "item_id", strDataKey)
    
    Call mclsVsf(0).ClearGrid
    Set rsTmp = gclsBusiness.ItemDeliverRead("item_id", rsCondition)
    If rsTmp.BOF = False Then
        Call mclsVsf(0).LoadGrid(rsTmp)
    End If
    
    Call ExecuteCommand(Command.显示投递目标)
    
    RefreshData = True
    
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
        Call .Initialize(Me.Controls, vsf(0), True, False, GetImageList(16))
        Call .ClearColumn
        
        Call .AppendColumn("", 270, flexAlignLeftCenter, flexDTString, , "[序号]", False, False, False)
        Call .AppendColumn("ID", 0, flexAlignLeftCenter, flexDTString, , "id", True, , , True)
        Call .AppendColumn("目标标识", 1800, flexAlignLeftCenter, flexDTString, , "deliver_code", True)
        Call .AppendColumn("目标命名", 1500, flexAlignLeftCenter, flexDTString, , "deliver_title", True)
        .IndicatorMode = 2
        .IndicatorCol = .ColIndex("序号")
        .ConstCol = .ColIndex("序号")
        
        .AppendRows = True
        
    End With
            
    Set mclsVsfDetail = New zlVSFlexGrid.clsVsf
    With mclsVsfDetail
        Call .Initialize(Me.Controls, vsfDetail, True, False)
        Call .ClearColumn
        
        Call .AppendColumn("", 270, flexAlignCenterCenter, flexDTString, , "[序号]", False, False, False)
        Call .AppendColumn("id", 0, flexAlignLeftCenter, flexDTString, , "", True, , , True)
        Call .AppendColumn("上级id", 0, flexAlignLeftCenter, flexDTString, , "", True, , , True)
        Call .AppendColumn("状态", 0, flexAlignLeftCenter, flexDTString, , "", True, , , True)
        Call .AppendColumn("目标内容", 1500, flexAlignLeftCenter, flexDTString, , "标题", True)
                
        .IndicatorMode = 2
        .IndicatorCol = .ColIndex("序号")
        .ConstCol = .ColIndex("序号")
        .VsfObject.OutlineCol = .ColIndex("目标内容")
        .VsfObject.RowHidden(0) = False
        
        .VsfObject.MergeCells = flexMergeFree
        .VsfObject.MergeCol(.ColIndex("类型")) = True
        .AppendRows = False
    End With

        
    InitGrid = True
    
End Function

Private Function ExecuteCommand(ByVal enmCommand As Command, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim rsCondition As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim rsDeliver As ADODB.Recordset
    Dim strTmp As String
    Dim strSQL As String
    Dim intRow As Integer
    Dim varTmp As Variant

    On Error GoTo errHand
    
    Set rsCondition = zlCommFun.CreateCondition
    
    Select Case enmCommand
        
    '------------------------------------------------------------------------------------------------------------------
    Case Command.新增投递目标
    
        If mfrmItemDeliveryEdit Is Nothing Then
            Set mfrmItemDeliveryEdit = New frmItemDeliveryEdit
            Call mfrmItemDeliveryEdit.InitDialog(mfrmParent, 11)
        End If
                
        Call mfrmItemDeliveryEdit.NewData(mstrDataKey)
        
        DoEvents
        Me.SetFocus
    '------------------------------------------------------------------------------------------------------------------
    Case Command.修改投递目标
    
        If mfrmItemDeliveryEdit Is Nothing Then
            Set mfrmItemDeliveryEdit = New frmItemDeliveryEdit
            Call mfrmItemDeliveryEdit.InitDialog(mfrmParent)
        End If
        
        With vsf(0)
            Call mfrmItemDeliveryEdit.ModifyData(mstrDataKey, .TextMatrix(.Row, .ColIndex("ID")))
            DoEvents
            Me.SetFocus
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case Command.删除投递目标
        
        If mfrmItemDeliveryEdit Is Nothing Then
            Set mfrmItemDeliveryEdit = New frmItemDeliveryEdit
            Call mfrmItemDeliveryEdit.InitDialog(mfrmParent)
        End If
        
        With vsf(0)
            If MsgBox("您确认要删除当前消息分类目录吗？", vbQuestion + vbYesNo + vbDefaultButton2, ParamInfo.系统名称) = vbYes Then
                Call mfrmItemDeliveryEdit.DeleteData(.TextMatrix(.Row, .ColIndex("ID")))
            End If
            DoEvents
            Me.SetFocus
        End With

        Me.SetFocus
                
    '------------------------------------------------------------------------------------------------------------------
    Case Command.显示投递目标
        
        
        mclsVsfDetail.ClearGrid
        
        Set rsCondition = zlCommFun.CreateCondition
        
        With vsf(0)
            Call zlCommFun.SetCondition(rsCondition, "id", .TextMatrix(.Row, .ColIndex("ID")))
        End With
        
        Set rsTmp = gclsBusiness.ItemDeliverRead("id", rsCondition)
        If rsTmp.BOF = False Then
            
            Set rsDeliver = New ADODB.Recordset
            With rsDeliver
                .Fields.Append "id", adVarChar, 100, adFldKeyColumn
                .Fields.Append "上级id", adVarChar, 200
                .Fields.Append "状态", adTinyInt
                .Fields.Append "标题", adVarChar, 200
                .Fields.Append "排序1", adBigInt
                .Fields.Append "排序2", adVarChar, 200
                .Open
            End With
    
            If gclsBusiness.GetDeliveryTree(zlCommFun.NVL(rsTmp("deliver_object").Value), rsDeliver) Then
                Call UpdateTargetGrid(rsDeliver)
            End If
            
        End If
                
    '------------------------------------------------------------------------------------------------------------------
    Case Command.刷新投递数据

        With vsf(0)
            mclsVsf(0).SaveKey = Trim(.TextMatrix(.Row, .ColIndex("id")))

            Call zlCommFun.SetCondition(rsCondition, "item_id", mstrDataKey)

            ExecuteCommand = mclsVsf(0).LoadDataSource(gclsBusiness.ItemDeliverRead("item_id", rsCondition))

            Call mclsVsf(0).RestoreRow(mclsVsf(0).SaveKey, .ColIndex("id"))
        End With

        
    '------------------------------------------------------------------------------------------------------------------
    Case Command.刷新指定信息

        ExecuteCommand = LoadCustomData(Trim(varParam(0)))
'        Call ExecuteCommand(Command.刷新投递数据)
        Call ExecuteCommand(Command.显示投递目标)
        
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


Private Sub UpdateTargetGrid(ByVal rsDeliver As ADODB.Recordset)
    Dim intMaxOutlineLevel As Integer
    Dim intLoop As Integer
        
    rsDeliver.Filter = ""
    rsDeliver.Sort = "排序1"
    If rsDeliver.RecordCount > 0 Then
        rsDeliver.MoveFirst
        With mclsVsfDetail
            .ClearGrid
            Call .LoadGrid(rsDeliver)
            intMaxOutlineLevel = .ShowOutline(.ColIndex("id"), .ColIndex("上级id"), vsfDetail.BackColor)
        
'            Call UpdateCollapseState
            .VsfObject.ShowCell .VsfObject.Row, .VsfObject.ColIndex("目标内容")
            .VsfObject.AutoSize .VsfObject.ColIndex("目标内容"), .VsfObject.ColIndex("目标内容")
        End With
    End If
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
    
    Set rsData = gclsBusiness.ItemDeliverRead("id", rsCondition)
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

Private Sub InitDockPannel()
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim objPane As Pane

    Set objPane = dkpMain.CreatePane(1, 100, 300, DockLeftOf, Nothing)
    objPane.Title = "SQL"
    objPane.Options = PaneNoCaption
        
    Set objPane = dkpMain.CreatePane(2, 100, 300, DockRightOf, objPane)
    objPane.Title = "SQL"
    objPane.Options = PaneNoCaption
    
    dkpMain.SetCommandBars cbsMain
    Call zlCommFun.DockPannelInit(dkpMain)

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
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_NewItem, "新增", , , xtpButtonIconAndCaption)
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Modify, "修改", , , xtpButtonIconAndCaption)
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Delete, "删除", , , xtpButtonIconAndCaption)
        
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Refresh, "刷新", True, , xtpButtonIconAndCaption)
    
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
    
End Function


Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_Edit_NewItem
        Call ExecuteCommand(Command.新增投递目标)
    Case conMenu_Edit_Modify
        Call ExecuteCommand(Command.修改投递目标)
    Case conMenu_Edit_Delete
        Call ExecuteCommand(Command.删除投递目标)
    End Select
    
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    Select Case Control.ID
    Case conMenu_Edit_NewItem
        Control.Visible = Not mblnSystem
        Control.Enabled = (Control.Visible And mstrDataKey <> "")
    Case conMenu_Edit_Modify
        Control.Visible = Not mblnSystem
        With vsf(0)
            Control.Enabled = (Control.Visible And mstrDataKey <> "" And .TextMatrix(.Row, .ColIndex("ID")) <> "")
        End With
    Case conMenu_Edit_Delete
        Control.Visible = Not mblnSystem
        With vsf(0)
            Control.Enabled = (Control.Visible And mstrDataKey <> "" And .TextMatrix(.Row, .ColIndex("ID")) <> "")
        End With
    End Select
    
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = picPane(1).hWnd
    Case 2
        Item.Handle = picPane(0).hWnd
    End Select
End Sub

Private Sub Form_Load()
    Call InitGrid
    Call InitDockPannel
    Call InitCommandBar
End Sub

Private Sub Form_Resize()
    On Error Resume Next
        
    Call zlCommFun.SetPaneRange(dkpMain, 2, 300, 15, 300, Me.ScaleHeight)
    
    dkpMain.RecalcLayout
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mclsVsf(0) = Nothing
    
    Set mobjFindKey = Nothing
End Sub

Private Sub mfrmItemDeliveryEdit_AfterDeleteData(ByVal DataKey As String)
    Call ExecuteCommand(Command.移除指定信息, DataKey)
End Sub

Private Sub mfrmItemDeliveryEdit_AfterModifyData(ByVal DataKey As String)
    Call ExecuteCommand(Command.刷新指定信息, DataKey)
End Sub

Private Sub mfrmItemDeliveryEdit_AfterNewData(ByVal DataKey As String)
    Call ExecuteCommand(Command.刷新指定信息, DataKey)
End Sub

Private Sub mfrmItemDeliveryEdit_Backward(DataKey As String, Cancel As Boolean)
    Dim intRow As Integer
    
    With vsf(0)
    
        intRow = mclsVsf(0).FindRow(DataKey, .ColIndex("id"))
        If intRow > 0 And .Row <> intRow Then .Row = intRow
        
        If .Row < .Rows - 1 Then
            .Row = .Row + 1
            .ShowCell .Row, .Col
            If DataKey = .TextMatrix(.Row, .ColIndex("id")) Then
                Cancel = True
            Else
                DataKey = .TextMatrix(.Row, .ColIndex("id"))
            End If
        End If
    End With
            
End Sub

Private Sub mfrmItemDeliveryEdit_Forward(DataKey As String, Cancel As Boolean)
    
    Dim intRow As Integer
    
    With vsf(0)
        
        intRow = mclsVsf(0).FindRow(DataKey, .ColIndex("id"))
        If intRow > 0 And .Row <> intRow Then .Row = intRow
                
        If .Row > 1 Then
            .Row = .Row - 1
            .ShowCell .Row, .Col
            If DataKey = .TextMatrix(.Row, .ColIndex("id")) Then
                Cancel = True
            Else
                DataKey = .TextMatrix(.Row, .ColIndex("id"))
            End If
        End If
    End With
    
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 0
        vsfDetail.Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 15
    Case 1
        vsf(0).Move 0, 15, picPane(Index).Width, picPane(Index).Height - 15
        mclsVsf(0).AppendRows = True
    End Select
End Sub


Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsf(0).AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
    If OldRow <> NewRow Then Call ExecuteCommand(Command.显示投递目标)
End Sub

Private Sub vsf_AfterScroll(Index As Integer, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mclsVsf(0).AppendRows = True
End Sub

Private Sub vsf_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    mclsVsf(0).AppendRows = True
End Sub


Private Sub vsf_DblClick(Index As Integer)
    Dim objMenu As CommandBarControl
    
    If mblnSystem Then
        Set objMenu = cbsMain.FindControl(, conMenu_Edit_View, False)
    Else
        Set objMenu = cbsMain.FindControl(, conMenu_Edit_Modify, False)
    End If
    
    If Not (objMenu Is Nothing) Then
        If objMenu.Enabled = True Then
            Call cbsMain_Execute(objMenu)
        End If
    End If
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

        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_NewItem, "新增(&N)")
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Modify, "修改(&M)")
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)")
                        
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

Private Sub vsfDetail_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsfDetail.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

