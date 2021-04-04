VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Begin VB.Form frmTableRelation 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   10785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   4065
      Index           =   0
      Left            =   4980
      ScaleHeight     =   4065
      ScaleWidth      =   2940
      TabIndex        =   2
      Top             =   870
      Width           =   2940
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1215
         Index           =   0
         Left            =   150
         TabIndex        =   3
         Top             =   165
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
         BackColor       =   -2147483633
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772055
         ForeColorSel    =   0
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483633
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483638
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   0
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   0
         GridLinesFixed  =   5
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
      Height          =   4065
      Index           =   1
      Left            =   375
      ScaleHeight     =   4065
      ScaleWidth      =   3750
      TabIndex        =   0
      Top             =   840
      Width           =   3750
      Begin MSComctlLib.TreeView tvw 
         Height          =   2505
         Left            =   270
         TabIndex        =   1
         Top             =   150
         Width           =   2550
         _ExtentX        =   4498
         _ExtentY        =   4419
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   476
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         Appearance      =   0
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
      Bindings        =   "frmTableRelation.frx":0000
      Left            =   540
      Top             =   120
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmTableRelation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################
Private mlngModualCode As Long
Private mstrPrivs As String
Private mstrSQL As String
Private mclsVsf(0) As zlVSFlexGrid.clsVsf
Private mblnStartUp As Boolean
Private mlngTmp As Long
Private mblnShowAll As Boolean
Private mblnShowStop As Boolean
Private mobjFindKey As CommandBarControl
Private mstrFindKey As String
Private mblnDataChanged As Boolean
Private mblnNew As Boolean
Private mstrDataKey As String
Private mfrmParent As Object
Private mrsCondition As ADODB.Recordset
Private mstrBusiness As String

Private Enum Command
    初始控件
    读注册表
    新增关系
    修改关系
    删除关系
    刷新数据
    刷新关系
    刷新指定关系
    移除指定关系
End Enum
'
Private WithEvents mfrmTableRelationEdit As frmTableRelationEdit
Attribute mfrmTableRelationEdit.VB_VarHelpID = -1

'######################################################################################################################
'接口方法
Public Function InitForm(ByVal frmParent As Object) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Set mfrmParent = frmParent
    
    Call ExecuteCommand(Command.初始控件)
    
    InitForm = True
    
End Function

Public Function RefreshData(ByVal strBusiness As String, ByVal strDataKey As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mstrBusiness = strBusiness
    mstrDataKey = strDataKey
    
    Call ExecuteCommand(Command.刷新数据)
    
    RefreshData = True
    
End Function

Public Sub Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call cbsMain_Execute(Control)
End Sub

Public Sub Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call cbsMain_Update(Control)
End Sub

'######################################################################################################################
'私有方法
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
        
        
        Call .AppendColumn("引用条件：", 1800, flexAlignLeftCenter, flexDTString, , "条件", True)
        vsf(0).Cell(flexcpFontBold, 0, 0) = True
        
'        .IndicatorMode = 2
'        .IndicatorCol = .ColIndex("序号")
'        .ConstCol = .ColIndex("序号")
'        vsf(0).RowHidden(0) = True
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
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlLabel, conMenu_View_LocationItem, "业务信息表关系", , , xtpButtonIconAndCaption)
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_NewItem, "新增", True)
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Modify, "修改")
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Delete, "删除")
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Refresh, "刷新", True)
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
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
    objPane.Close
    picPane(0).Visible = False
    
    dkpMain.SetCommandBars cbsMain
    Call zlCommFun.DockPannelInit(dkpMain)

End Sub

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
        
        Set tvw.ImageList = gfrmPubResource.GetImageCtl
        
    '------------------------------------------------------------------------------------------------------------------
    Case Command.新增关系
    
        If mfrmTableRelationEdit Is Nothing Then
            Set mfrmTableRelationEdit = New frmTableRelationEdit
            Call mfrmTableRelationEdit.InitDialog(mfrmParent)
        End If
    
        Call mfrmTableRelationEdit.NewData(mstrBusiness, mstrDataKey, tvw.SelectedItem.Text)
        DoEvents
        Me.SetFocus
    '------------------------------------------------------------------------------------------------------------------
    Case Command.修改关系
    
        If mfrmTableRelationEdit Is Nothing Then
            Set mfrmTableRelationEdit = New frmTableRelationEdit
            Call mfrmTableRelationEdit.InitDialog(mfrmParent)
        End If
        
        strTmp = tvw.SelectedItem.Key
        If InStr(strTmp, "R_") > 0 Then
            Call mfrmTableRelationEdit.ModifyData(mstrBusiness, mstrDataKey, Mid(strTmp, InStr(strTmp, "R_") + 2, 32))
        End If

        DoEvents
        Me.SetFocus
    '------------------------------------------------------------------------------------------------------------------
    Case Command.删除关系
        
        If mfrmTableRelationEdit Is Nothing Then
            Set mfrmTableRelationEdit = New frmTableRelationEdit
            Call mfrmTableRelationEdit.InitDialog(mfrmParent)
        End If
        strTmp = tvw.SelectedItem.Key
        If InStr(strTmp, "R_") > 0 Then
            If MsgBox("您确认要删除当前业务信息表关系吗？", vbQuestion + vbYesNo + vbDefaultButton2, ParamInfo.系统名称) = vbYes Then
                Call mfrmTableRelationEdit.DeleteData(mstrBusiness, Mid(strTmp, InStr(strTmp, "R_") + 2, 32))
            End If
        End If

        DoEvents
        Me.SetFocus
                
    '------------------------------------------------------------------------------------------------------------------
    Case Command.刷新数据
        
        Dim objNode As Node
        Dim objRootNode As Node
        
        Set mrsCondition = zlCommFun.CreateCondition
        Call zlCommFun.SetCondition(mrsCondition, "id", mstrDataKey)
        
        tvw.Nodes.Clear
        Set rs = gclsBusiness.GetTableTree(mstrDataKey)
        If Not (rs Is Nothing) Then

            If rs.RecordCount > 0 Then
                rs.MoveFirst
                Do While Not rs.EOF
                    
                    If zlCommFun.NVL(rs("上级id").Value) = "" Then
                        Set objNode = tvw.Nodes.Add(, , "K_" & rs("id").Value, rs("名称").Value)
                        objNode.Expanded = True
                    Else
                        Set objNode = tvw.Nodes.Add("K_" & rs("上级id").Value, tvwChild, "K_" & rs("id").Value, rs("名称").Value)
                        objNode.Expanded = False
                    End If
                    
                    
                    If objRootNode Is Nothing Then Set objRootNode = objNode
                    
                    If Not (objNode.Parent Is Nothing) Then
                        If objNode.Parent <> objRootNode Then
                            objNode.ForeColor = RGB(192, 192, 192)
                        End If
                    End If
                    
                    
                    
                    'constitute
                    
                    If Val(rs("关系").Value) = 2 Then
                        objNode.Image = "constitute"
                    Else
                        objNode.Image = IIf(Val(rs("类型").Value) = 0, "folder", "file")
                    End If
                    
                    rs.MoveNext
                Loop
            End If
            
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case Command.刷新关系
        
        If InStr(tvw.SelectedItem.Key, "R_") > 0 Then
            Set mrsCondition = zlCommFun.CreateCondition
            Call zlCommFun.SetCondition(mrsCondition, "id", Mid(tvw.SelectedItem.Key, InStr(tvw.SelectedItem.Key, "R_") + 2, 32))
            
            Call mclsVsf(0).LoadGrid(gclsBusiness.TableRelationRead("Condition", mrsCondition))
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case Command.刷新指定关系
        
        ExecuteCommand = LoadCustomData(Trim(varParam(0)))
        
    '------------------------------------------------------------------------------------------------------------------
    Case Command.移除指定关系
        
        With vsf(0)
            intRow = mclsVsf(0).FindRow(Trim(varParam(0)), .ColIndex("id"))
            If intRow > 0 Then
                If .Rows > 2 Then
                    .RemoveItem .Row
'                    mclsVsf(0).AppendRows = True
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

Private Function ShowConetneMenu(Optional ByVal bytPlace As Byte = 1) As CommandBar
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
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_NewItem, "新增关系(&N)")
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Modify, "修改关系(&M)")
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Delete, "删除关系(&D)")

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
    
    Set rsData = gclsBusiness.TableRelationRead("id", rsCondition)
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
    
'    mclsVsf(0).AppendRows = True
    
    LoadCustomData = True
    
End Function


'######################################################################################################################
'控件事件
Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_NewItem, conMenu_Edit_NewParent              '增加
        
        Call ExecuteCommand(Command.新增关系)
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Modify, conMenu_Edit_ModifyParent              '修改
        
        Call ExecuteCommand(Command.修改关系)
                
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete, conMenu_Edit_DeleteParent

        Call ExecuteCommand(Command.删除关系)
            
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Refresh               '刷新

        Call ExecuteCommand(Command.刷新数据)
    
    End Select
            
End Sub


Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objParentNode As Node
    With vsf(0)
        Select Case Control.ID
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_NewItem                   '增加
            
            If tvw.SelectedItem Is Nothing Then
                Control.Enabled = False
            Else
                Set objParentNode = tvw.SelectedItem.Parent
                If objParentNode Is Nothing Then
                    Control.Enabled = True
                Else
                    Control.Enabled = (tvw.SelectedItem.Image = "file" And objParentNode.Parent Is Nothing)
                End If
            End If
            
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Modify, conMenu_Edit_Delete                 '修改，删除
                        
            If tvw.SelectedItem Is Nothing Then
                Control.Enabled = False
            Else
                Set objParentNode = tvw.SelectedItem.Parent
                If objParentNode Is Nothing Then
                    Control.Enabled = False
                Else
                    Control.Enabled = (tvw.SelectedItem.Image <> "file" And Not (objParentNode Is Nothing) And objParentNode.Parent Is Nothing)
                End If
            End If
            
        End Select
    End With
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
    Call InitCommandBar
    Call InitDockPannel
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Call zlCommFun.SetPaneRange(dkpMain, 2, 200, 15, 200, Me.ScaleHeight)
    
    dkpMain.RecalcLayout
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mclsVsf(0) = Nothing
    Set mobjFindKey = Nothing
    If Not (mfrmTableRelationEdit Is Nothing) Then
        Unload mfrmTableRelationEdit
        Set mfrmTableRelationEdit = Nothing
    End If
End Sub

Private Sub mfrmTableRelationEdit_AfterDeleteData(ByVal DataKey As String)
    Call ExecuteCommand(Command.移除指定关系, DataKey)
End Sub

Private Sub mfrmTableRelationEdit_AfterModifyData(ByVal DataKey As String)
    Call ExecuteCommand(Command.刷新关系, DataKey)
End Sub

Private Sub mfrmTableRelationEdit_AfterNewData(ByVal DataKey As String)
    Call ExecuteCommand(Command.刷新关系, DataKey)
End Sub

Private Sub mfrmTableRelationEdit_Backward(DataKey As String, Cancel As Boolean)
    Dim intRow As Integer
    
    With vsf(0)
    
        intRow = mclsVsf(0).FindRow(DataKey, .ColIndex("id"))
        If intRow > 0 And .Row <> intRow Then .Row = intRow
        
        intRow = GetNewRow(.Row, 1)
        
        If intRow >= 0 Then
            .Row = intRow
            .ShowCell .Row, .Col
            DataKey = .TextMatrix(.Row, .ColIndex("id"))
        End If
        
'        If .Row < .Rows - 1 Then
'            .Row = .Row + 1
'            .ShowCell .Row, .Col
'            DataKey = .TextMatrix(.Row, .ColIndex("id"))
'        End If
        
    End With
            
End Sub

Private Function GetNewRow(ByVal lngStartRow As Long, ByVal bytWay As Byte) As Long
    Dim lngRow As Long
    
    GetNewRow = -1
    
    With vsf(0)
        If bytWay = 1 Then
            For lngRow = lngStartRow + 1 To .Rows - 1
                
                If Val(.TextMatrix(lngRow, .ColIndex("标志"))) = 0 Then
                    GetNewRow = lngRow
                    Exit For
                End If
            Next
        Else
            For lngRow = lngStartRow - 1 To 1 Step -1
                
                If Val(.TextMatrix(lngRow, .ColIndex("标志"))) = 0 Then
                    GetNewRow = lngRow
                    Exit For
                End If
            Next
        End If
        
    End With
End Function

Private Sub mfrmTableRelationEdit_Forward(DataKey As String, Cancel As Boolean)
    
    Dim intRow As Integer
    
    With vsf(0)
        
        intRow = mclsVsf(0).FindRow(DataKey, .ColIndex("id"))
        If intRow > 0 And .Row <> intRow Then .Row = intRow
                        
        intRow = GetNewRow(.Row, 2)
        If intRow >= 0 Then
            .Row = intRow
            .ShowCell .Row, .Col
            DataKey = .TextMatrix(.Row, .ColIndex("id"))
        End If
        
'        If .Row > 1 Then
'            .Row = .Row - 1
'            .ShowCell .Row, .Col
'            DataKey = .TextMatrix(.Row, .ColIndex("id"))
'        End If
    End With
    
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 0
        vsf(0).Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
'        mclsVsf(0).AppendRows = True
    Case 1
        tvw.Move 0, 15, picPane(Index).Width, picPane(Index).Height - 15
    End Select
End Sub

Private Sub tvw_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbrPopupBar As CommandBar
    
    Select Case Button
    '------------------------------------------------------------------------------------------------------------------
    Case 2          '弹出菜单处理
        Call zlCommFun.SendLMouseButton(tvw.hWnd, X, Y)
            
        Call ShowConetneMenu(1).ShowPopup
        
    End Select
End Sub

Private Sub tvw_NodeClick(ByVal Node As MSComctlLib.Node)
    
    
    If InStr(Node.Key, "F_") <= 0 Then
        If InStr(Node.Key, "R_") > 0 Then
            If dkpMain.Panes(2).Selected = False Then dkpMain.Panes(2).Select
            If Node.Image = "folder" Then
                vsf(0).TextMatrix(0, 0) = "引用条件："
            Else
                vsf(0).TextMatrix(0, 0) = "组合条件："
            End If
            Call ExecuteCommand(Command.刷新关系)
        Else
            If dkpMain.Panes(2).Closed = False Then dkpMain.Panes(2).Close
        End If
    Else
        If dkpMain.Panes(2).Closed = False Then dkpMain.Panes(2).Close
    End If
    tvw.SetFocus
    
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsf(Index).AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsf_AfterScroll(Index As Integer, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
'    mclsVsf(Index).AppendRows = True
End Sub

Private Sub vsf_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
'    mclsVsf(Index).AppendRows = True
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
